"""調査用: 指定期間のNE出荷実績(入出庫履歴)をSKU単位で集計し、シート「出荷集計(臨時)」に出す。
NEは読み取りのみ。

環境変数:
  SHIP_SINCE  既定 2026-10-27 18:00:00（この日時以降）
  SHIP_UNTIL  任意。指定するとこの日時より前だけ集計（例: 2026-10-31 00:00:00）
  他は stock_snapshot.py と共通
"""
from __future__ import annotations
import os
from collections import defaultdict, Counter
from datetime import datetime
from ne_client import ne_post
from stock_snapshot import open_spreadsheet, read_jan_sku_map, fetch_all, to_int, JST

SINCE = os.environ.get("SHIP_SINCE", "2026-10-27 18:00:00")
UNTIL = os.environ.get("SHIP_UNTIL", "").strip()
SHEET_OUT = "出荷集計(臨時)"
HIST_FIELDS = ",".join([
    "stock_io_history_id", "stock_io_history_date", "stock_io_history_goods_id",
    "stock_io_history_quantity", "stock_io_history_io_flag", "stock_io_history_io_type_name",
])


def main():
    params = {
        "fields": HIST_FIELDS,
        "stock_io_history_date-gte": SINCE,
        "stock_io_history_deleted_flag-eq": "0",
    }
    if UNTIL:
        params["stock_io_history_date-lt"] = UNTIL
    rows = fetch_all("/api_v1_master_stockiohistory/search", params, page_size=1000)
    print(f"履歴 {len(rows)}件 (since={SINCE} until={UNTIL or '指定なし'})")

    types = Counter((r.get("stock_io_history_io_type_name"), str(r.get("stock_io_history_io_flag"))) for r in rows)
    for (t, f), c in types.most_common():
        print(f"  種別 {t} / io_flag={f}: {c}件")

    agg = defaultdict(lambda: defaultdict(int))
    for r in rows:
        sku = r.get("stock_io_history_goods_id")
        q = abs(to_int(r.get("stock_io_history_quantity")))
        is_out = str(r.get("stock_io_history_io_flag")) == "1"
        is_ship = "出荷" in (r.get("stock_io_history_io_type_name") or "")
        key = ("ship_out" if is_out else "ship_in") if is_ship else ("other_out" if is_out else "other_in")
        agg[sku][key] += q

    ss = open_spreadsheet()
    mapping, _ = read_jan_sku_map(ss)
    sku_jans = defaultdict(list)
    for jan, sku in mapping.items():
        sku_jans[sku].append(jan)

    out = [["SKU", "JAN", "出荷数(出庫)", "出荷系の入庫(取消等)", "差引出荷数", "その他の出庫", "その他の入庫"]]
    for sku in sorted(agg):
        a = agg[sku]
        for jan in (sku_jans.get(sku) or [""]):
            out.append([sku, jan, a["ship_out"], a["ship_in"], a["ship_out"] - a["ship_in"], a["other_out"], a["other_in"]])
    out.append([f"取得 {datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S')} / since {SINCE} / until {UNTIL or '指定なし'}"])

    try:
        ws = ss.worksheet(SHEET_OUT)
        ws.clear()
    except Exception:
        ws = ss.add_worksheet(title=SHEET_OUT, rows=max(len(out) + 50, 200), cols=len(out[0]))
    if ws.row_count < len(out):
        ws.add_rows(len(out) - ws.row_count)
    ws.update(range_name="A1", values=out, value_input_option="RAW")
    print(f"「{SHEET_OUT}」に{len(out) - 2}行を書き込みました (SKU {len(agg)}件)")


if __name__ == "__main__":
    main()
