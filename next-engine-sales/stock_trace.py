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
LIMIT = int(os.environ.get("TRACE_LIMIT", "50"))


def main():
    if not SKUS:
        raise SystemExit("TRACE_SKUS が空です")
    for sku in SKUS:
        print(f"===== {sku} (since {SINCE}) =====")
        # 在庫マスタ（現在値）
        st = ne_post("/api_v1_master_stock/search", {"stock_goods_id-eq": sku}).get("data", [])
        for r in st:
            print("[master_stock]", json.dumps(r, ensure_ascii=False))
        # 入出庫履歴（フィールド指定なし = 全項目）
        d = ne_post("/api_v1_master_stockiohistory/search", {
            "stock_io_history_goods_id-eq": sku,
            "stock_io_history_date-gte": SINCE,
            "limit": LIMIT,
        })
        rows = d.get("data", [])
        print(f"[stockiohistory] {len(rows)}件")
        for r in rows:
            print(json.dumps(r, ensure_ascii=False))


if __name__ == "__main__":
    main()
