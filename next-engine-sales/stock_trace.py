"""調査用: 指定SKUのNE入出庫履歴(master_stockiohistory)を表示する。読み取りのみ。
Usage:
  TRACE_SKUS=8901360990032,5089997101 TRACE_SINCE="2026-10-27 18:00:00" python stock_trace.py

最初はフィールドを絞らずに取得し、項目名と中身を確認する。
"""
from __future__ import annotations
import os, json
from ne_client import ne_post

SKUS = [s.strip() for s in os.environ.get("TRACE_SKUS", "").split(",") if s.strip()]
SINCE = os.environ.get("TRACE_SINCE", "2026-10-27 18:00:00")
LIMIT = int(os.environ.get("TRACE_LIMIT", "200"))
HIST_FIELDS = ",".join([
    "stock_io_history_id", "stock_io_history_date", "stock_io_history_goods_id",
    "stock_io_history_before_stock_quantity", "stock_io_history_after_stock_quantity",
    "stock_io_history_quantity", "stock_io_history_io_flag",
    "stock_io_history_io_type_id", "stock_io_history_io_type_name",
    "stock_io_history_reason", "stock_io_history_cut_form_id",
    "stock_io_history_pic_name", "stock_io_history_creator_name",
])


def main():
    if not SKUS:
        raise SystemExit("TRACE_SKUS が空です")
    for sku in SKUS:
        print(f"===== {sku} (since {SINCE}) =====")
        # 在庫マスタ（現在値）
        st = ne_post("/api_v1_master_stock/search", {
            "fields": "stock_goods_id,stock_quantity,stock_allocation_quantity,stock_free_quantity",
            "stock_goods_id-eq": sku,
        }).get("data", [])
        for r in st:
            print("[master_stock]", json.dumps(r, ensure_ascii=False))
        # 入出庫履歴（NE APIは fields 指定が必須）
        d = ne_post("/api_v1_master_stockiohistory/search", {
            "fields": HIST_FIELDS,
            "stock_io_history_goods_id-eq": sku,
            "stock_io_history_date-gte": SINCE,
            "stock_io_history_deleted_flag-eq": "0",
            "limit": LIMIT,
        })
        rows = sorted(d.get("data", []), key=lambda r: r.get("stock_io_history_date", ""))
        print(f"[stockiohistory] {len(rows)}件 (入出庫日の昇順)")
        for r in rows:
            print(json.dumps(r, ensure_ascii=False))


if __name__ == "__main__":
    main()
