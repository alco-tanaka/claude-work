"""前日売上集計 → Slack 通知
Usage:
  python daily_sales.py          # 通常実行（Slack投稿あり）
  python daily_sales.py --preview # プレビューのみ（Slack投稿なし）
"""
from __future__ import annotations
import sys, os, requests
from datetime import datetime, timedelta, timezone
from collections import defaultdict
from ne_client import ne_post, load_env

JST = timezone(timedelta(hours=9))

def yesterday_range():
    today = datetime.now(JST).replace(hour=0, minute=0, second=0, microsecond=0)
    y = today - timedelta(days=1)
    f = "%Y-%m-%d %H:%M:%S"
    return y.strftime(f), today.strftime(f)

def fetch_all(endpoint, params, page_size=500):
    rows, offset = [], 0
    while True:
        d = ne_post(endpoint, {**params, "limit": page_size, "offset": offset})
        chunk = d.get("data", [])
        rows.extend(chunk)
        if len(chunk) < page_size: break
        offset += page_size
    return rows

def to_int(v):
    try: return int(float(v))
    except: return 0

def jpy(n): return f"¥{n:,}"

def common_name(names: list[str]) -> str:
    """商品名リストの共通部分を抽出し、末尾の区切り文字を除去する"""
    if not names:
        return ""
    prefix = os.path.commonprefix(sorted(names))
    clean = prefix.rstrip("/ 　・-_（()）0123456789").strip()
    return clean if clean else names[0]

def fetch_goods_info(goods_ids: list[str]) -> dict[str, dict]:
    """goods_id → {rep_id, name} のマッピングを返す（200件バッチ）"""
    mapping = {}
    for i in range(0, len(goods_ids), 200):
        chunk = goods_ids[i:i + 200]
        data = fetch_all("/api_v1_master_goods/search", {
            "fields": "goods_id,goods_name,goods_representation_id",
            "goods_id-in": ",".join(chunk),
        })
        for row in data:
            gid = row.get("goods_id", "")
            mapping[gid] = {
                "rep_id": row.get("goods_representation_id") or gid,
                "name":   row.get("goods_name", ""),
            }
    return mapping

def build_report():
    start, end = yesterday_range()
    date_label = start[:10]
    shop_names = {r["shop_id"]: r["shop_name"] for r in fetch_all("/api_v1_master_shop/search", {"fields": "shop_id,shop_name"})}
    orders = fetch_all("/api_v1_receiveorder_base/search", {
        "fields": "receive_order_id,receive_order_shop_id,receive_order_total_amount,receive_order_date,receive_order_cancel_type_id",
        "receive_order_date-gte": start, "receive_order_date-lt": end})
    active = [o for o in orders if to_int(o.get("receive_order_cancel_type_id")) == 0]
    active_ids = {o["receive_order_id"] for o in active}

    # 明細は日付フィルター不可のため受注IDで絞り込む（100件ずつバッチ）
    rows = []
    id_list = list(active_ids)
    for i in range(0, max(len(id_list), 1), 100):
        chunk = id_list[i:i + 100]
        if not chunk: break
        rows.extend(fetch_all("/api_v1_receiveorder_row/search", {
            "fields": "receive_order_row_goods_id,receive_order_row_goods_name,receive_order_row_quantity,receive_order_row_receive_order_id",
            "receive_order_row_receive_order_id-in": ",".join(chunk)}))

    # 店舗別集計
    by_shop = defaultdict(lambda: {"count": 0, "amount": 0})
    for o in active:
        s = by_shop[o.get("receive_order_shop_id", "")]
        s["count"] += 1
        s["amount"] += to_int(o.get("receive_order_total_amount"))
    total_count = sum(s["count"] for s in by_shop.values())
    total_amount = sum(s["amount"] for s in by_shop.values())
    avg = (total_amount // total_count) if total_count else 0

    # 受注明細のgoods_idから goods_representation_id と goods_name を取得
    unique_ids = list({r.get("receive_order_row_goods_id", "") for r in rows if r.get("receive_order_row_goods_id")})
    goods_info = fetch_goods_info(unique_ids)

    # 代表商品コード単位で数量・商品名を集計
    by_rep: dict[str, dict] = defaultdict(lambda: {"names": [], "qty": 0})
    for r in rows:
        gid = r.get("receive_order_row_goods_id", "")
        if not gid: continue
        info = goods_info.get(gid, {})
        rep_id = info.get("rep_id", gid)
        # 商品名はマスターから取得した名称を優先し、なければ受注明細の名称を使う
        name = info.get("name") or r.get("receive_order_row_goods_name", "")
        g = by_rep[rep_id]
        if name and name not in g["names"]:
            g["names"].append(name)
        g["qty"] += to_int(r.get("receive_order_row_quantity"))

    top10 = sorted(
        [{"rep_id": rep, "name": common_name(g["names"]) or rep, "qty": g["qty"]}
         for rep, g in by_rep.items()],
        key=lambda x: x["qty"], reverse=True
    )[:10]

    L = [f"*ネクストエンジン売上速報 — {date_label}*", "",
         f"合計: {jpy(total_amount)} / {total_count}件 / 客単価 {jpy(avg)}", "",
         "*店舗(モール)別*"]
    if not by_shop:
        L.append("  (受注なし)")
    else:
        for sid, s in sorted(by_shop.items(), key=lambda kv: kv[1]["amount"], reverse=True):
            name = shop_names.get(sid, f"店舗ID:{sid}")
            a = (s["amount"] // s["count"]) if s["count"] else 0
            L.append(f"  • {name}: {s['count']}件 / {jpy(s['amount'])} / 客単価 {jpy(a)}")
    L += ["", "*商品別数量 TOP10（代表商品コード単位）*"]
    if not top10:
        L.append("  (受注なし)")
    else:
        for i, g in enumerate(top10, 1):
            L.append(f"  {i}. {g['name']} — {g['qty']}個")
    return "\n".join(L)

def post_to_slack(text):
    env = load_env()
    urls = [v for k, v in env.items() if k.startswith("SLACK_WEBHOOK_URL") and v]
    if not urls: raise RuntimeError(".env に SLACK_WEBHOOK_URL がありません")
    for url in urls:
        r = requests.post(url, json={"text": text, "mrkdwn": True}, timeout=30)
        if r.status_code != 200: raise RuntimeError(f"Slack投稿失敗: {r.status_code} {r.text}")

def main():
    preview = "--preview" in sys.argv
    try:
        report = build_report()
        print(report)
        if preview:
            print("\n（プレビューモード: Slack投稿スキップ）")
        else:
            post_to_slack(report)
            print("\n✅ Slack投稿成功")
    except Exception as e:
        msg = f"⚠️ ネクストエンジン売上通知エラー: {e}"
        print(msg, file=sys.stderr)
        if not preview:
            try: post_to_slack(msg)
            except: pass
        sys.exit(1)

if __name__ == "__main__":
    main()
