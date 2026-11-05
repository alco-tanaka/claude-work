"""調査用: 卸フリー在庫がマイナスの商品について、NEの出荷数(since以降)と照合する。
NEは読み取りのみ。結果はシート「卸フリーマイナス検証(臨時)」に出力。

環境変数: CHECK_SINCE (既定 2026-10-27 18:00:00), 他は stock_snapshot.py と共通
"""
from __future__ import annotations
import os
from datetime import datetime
from collections import defaultdict
from ne_client import ne_post
from stock_snapshot import (open_spreadsheet, read_jan_sku_map, norm_jan, fetch_all,
                            fetch_stock, fetch_names, to_int, SHEET_MASTER, MASTER_HEADER_ROW, JST)

SINCE = os.environ.get("CHECK_SINCE", "2026-10-27 18:00:00")
SHEET_OUT = "卸フリーマイナス検証(臨時)"
HIST_FIELDS = ",".join([
    "stock_io_history_id", "stock_io_history_date", "stock_io_history_goods_id",
    "stock_io_history_quantity", "stock_io_history_io_flag",
    "stock_io_history_io_type_name", "stock_io_history_reason",
])


def main():
    ss = open_spreadsheet()
    vals = ss.worksheet(SHEET_MASTER).get_all_values()
    mapping, _ = read_jan_sku_map(ss)

    targets = []  # (jan, K, L, P)
    for row in vals[MASTER_HEADER_ROW:]:
        if len(row) < 16 or not row[9].strip():
            continue
        if to_int(row[11]) < 0:
            targets.append((norm_jan(row[9]), to_int(row[10]), to_int(row[11]), to_int(row[15])))
    print(f"卸フリーマイナス: {len(targets)}件")

    skus = sorted({mapping[j] for j, *_ in targets if j in mapping})
    stock = fetch_stock(skus)
    names = fetch_names(skus)

    out = [["SKU", "JAN", "商品名", "アラジン現在庫", "アラジン卸フリー", "アラジンWEB在庫",
            "NE総在庫", "NE引当", "出荷数(since以降)", "出荷以外の出庫", "入庫", "卸フリー+出荷数",
            "WEB在庫-NE総", "備考"]]
    summary = defaultdict(int)
    for jan, k, l, p in targets:
        sku = mapping.get(jan)
        if not sku:
            out.append(["", jan, "", k, l, p, "", "", "", "", "", "", "", "対照表未登録"])
            continue
        hist = fetch_all("/api_v1_master_stockiohistory/search", {
            "fields": HIST_FIELDS,
            "stock_io_history_goods_id-eq": sku,
            "stock_io_history_date-gte": SINCE,
            "stock_io_history_deleted_flag-eq": "0",
        })
        ship = other_out = inn = 0
        for h in hist:
            q = abs(to_int(h.get("stock_io_history_quantity")))
            if "出荷" in (h.get("stock_io_history_io_type_name") or ""):
                ship += q
            elif str(h.get("stock_io_history_io_flag")) == "1":
                other_out += q
            else:
                inn += q
        s = stock.get(sku, {})
        ne_total = to_int(s.get("stock_quantity")) if s else ""
        diff = (p - ne_total) if s else ""
        out.append([sku, jan, names.get(sku, ""), k, l, p, ne_total,
                    to_int(s.get("stock_allocation_quantity")) if s else "",
                    ship, other_out, inn, l + ship, diff, "" if s else "NE未登録"])
        summary["total"] += 1
        if -l == ship:
            summary["一致(マイナス=出荷数)"] += 1
        if l + ship == 0 and other_out == 0 and inn == 0:
            summary["出荷のみで完全一致"] += 1
    print(dict(summary))

    now = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
    out.append([f"取得日時 {now} / since {SINCE}"])
    try:
        ws = ss.worksheet(SHEET_OUT)
        ws.clear()
    except Exception:
        ws = ss.add_worksheet(title=SHEET_OUT, rows=max(len(out) + 50, 200), cols=len(out[0]))
    ws.update(range_name="A1", values=out, value_input_option="RAW")
    print(f"「{SHEET_OUT}」に{len(out) - 2}行を書き込みました")


if __name__ == "__main__":
    main()
