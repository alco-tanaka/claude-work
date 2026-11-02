"""NE在庫スナップショット → スプレッドシート「NE在庫」シートへ書き込み
WEB在庫マスタ出力シートのJAN → JAN-SKU対照表でSKU変換 → NE在庫マスタを取得。
NEへは読み取りのみ（書き込みなし）。

Usage:
  python stock_snapshot.py             # 通常実行（シートへ書き込み）
  python stock_snapshot.py --dry-run   # 取得のみ（シートへ書き込まない）

環境変数:
  GOOGLE_CREDENTIALS_JSON  サービスアカウントキー(JSON文字列)
  STOCK_SPREADSHEET_ID     対象スプレッドシートID
  STOCK_BRANDS             任意。ブランドコードをカンマ区切りで指定すると絞り込む（例: 089,123）
  NE_ACCESS_TOKEN / NE_REFRESH_TOKEN / NE_TOKEN_OUTPUT  (ne_client.py と共通)
"""
from __future__ import annotations
import sys, os, json
from datetime import datetime, timedelta, timezone
from ne_client import ne_post

JST = timezone(timedelta(hours=9))

SHEET_MASTER = "WEB在庫マスタ出力"
SHEET_MAP = "JAN-SKU対照表"
SHEET_OUT = "NE在庫"
MASTER_HEADER_ROW = 4          # 見出し行（1〜3行目は出力条件）
OUT_HEADER = ["SKU", "JAN", "商品名", "総在庫", "引当数", "フリー在庫", "取得日時", "備考"]
BATCH = 200


def to_int(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 0


def norm_jan(v) -> str:
    s = str(v).strip()
    return s.zfill(13) if s.isdigit() else s


def fetch_all(endpoint, params, page_size=500):
    rows, offset = [], 0
    while True:
        d = ne_post(endpoint, {**params, "limit": page_size, "offset": offset})
        chunk = d.get("data", [])
        rows.extend(chunk)
        if len(chunk) < page_size:
            break
        offset += page_size
    return rows


def open_spreadsheet():
    import gspread
    info = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
    gc = gspread.service_account_from_dict(info)
    return gc.open_by_key(os.environ["STOCK_SPREADSHEET_ID"])


def read_targets(ss):
    """WEB在庫マスタ出力から (JAN, ブランド) の一覧を返す。"""
    vals = ss.worksheet(SHEET_MASTER).get_all_values()
    brands = {b.strip() for b in os.environ.get("STOCK_BRANDS", "").split(",") if b.strip()}
    jans, seen = [], set()
    for row in vals[MASTER_HEADER_ROW:]:
        if len(row) < 10 or not row[9].strip():
            continue
        if brands and row[0].strip() not in brands:
            continue
        jan = norm_jan(row[9])
        if jan not in seen:
            seen.add(jan)
            jans.append(jan)
    return jans


def read_jan_sku_map(ss):
    """JAN-SKU対照表（C列=JAN, D列=SKU）→ {JAN: SKU}。重複・桁不正は警告。"""
    vals = ss.worksheet(SHEET_MAP).get_all_values()
    mapping, warns = {}, []
    for i, row in enumerate(vals[1:], start=2):
        if len(row) < 4 or not row[2].strip() or not row[3].strip():
            continue
        jan, sku = norm_jan(row[2]), row[3].strip()
        if len(jan) != 13 or not jan.isdigit():
            warns.append(f"対照表{i}行目: JANが13桁の数字ではありません ({row[2]})")
        if jan in mapping and mapping[jan] != sku:
            warns.append(f"対照表{i}行目: JAN {jan} に複数のSKU ({mapping[jan]} / {sku})")
            continue
        mapping[jan] = sku
    return mapping, warns


def fetch_stock(skus):
    stock = {}
    for i in range(0, len(skus), BATCH):
        chunk = skus[i:i + BATCH]
        for r in fetch_all("/api_v1_master_stock/search", {
            "fields": "stock_goods_id,stock_quantity,stock_allocation_quantity,stock_free_quantity",
            "stock_goods_id-in": ",".join(chunk),
        }):
            stock[r["stock_goods_id"]] = r
    return stock


def fetch_names(skus):
    names = {}
    for i in range(0, len(skus), BATCH):
        chunk = skus[i:i + BATCH]
        for r in fetch_all("/api_v1_master_goods/search", {
            "fields": "goods_id,goods_name",
            "goods_id-in": ",".join(chunk),
        }):
            names[r["goods_id"]] = r.get("goods_name", "")
    return names


def main():
    dry = "--dry-run" in sys.argv
    ss = open_spreadsheet()

    jans = read_targets(ss)
    mapping, warns = read_jan_sku_map(ss)
    print(f"対象JAN: {len(jans)}件 / 対照表: {len(mapping)}件")

    pairs, unmapped = [], []
    for jan in jans:
        sku = mapping.get(jan)
        (pairs if sku else unmapped).append((jan, sku) if sku else jan)

    skus = sorted({sku for _, sku in pairs})
    stock = fetch_stock(skus)
    names = fetch_names(skus)
    now = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")

    out = [OUT_HEADER]
    for jan, sku in pairs:
        s = stock.get(sku)
        if s is None:
            out.append([sku, jan, names.get(sku, ""), "", "", "", now, "NE在庫マスタに未登録"])
        else:
            out.append([sku, jan, names.get(sku, ""),
                        to_int(s.get("stock_quantity")),
                        to_int(s.get("stock_allocation_quantity")),
                        to_int(s.get("stock_free_quantity")), now, ""])
    for jan in unmapped:
        out.append(["", jan, "", "", "", "", now, "JAN-SKU対照表に未登録"])

    not_in_ne = sum(1 for r in out[1:] if r[7] == "NE在庫マスタに未登録")
    print(f"取得: {len(out) - 1 - len(unmapped) - not_in_ne}件 / NE未登録: {not_in_ne}件 / 対照表未登録: {len(unmapped)}件")
    for w in warns[:20]:
        print(f"[警告] {w}")
    if len(warns) > 20:
        print(f"[警告] ほか{len(warns) - 20}件")

    if dry:
        print("--dry-run のためシートへは書き込みません")
        for r in out[:6]:
            print(r)
        return

    try:
        ws = ss.worksheet(SHEET_OUT)
    except Exception:
        ws = ss.add_worksheet(title=SHEET_OUT, rows=max(len(out) + 100, 1000), cols=len(OUT_HEADER))
    ws.clear()
    if ws.row_count < len(out):
        ws.add_rows(len(out) - ws.row_count)
    # RAW: JAN/SKUを文字列のまま保持（先頭の0を落とさない）
    ws.update(range_name="A1", values=out, value_input_option="RAW")
    print(f"「{SHEET_OUT}」シートに{len(out) - 1}行を書き込みました")


if __name__ == "__main__":
    main()
