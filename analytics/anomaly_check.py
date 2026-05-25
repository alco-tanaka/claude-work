"""GA4 異常検知 — 毎日実行してセッション急増/急減をSlackに通知
Usage:
  python anomaly_check.py            # 本番Slack (#01_webdiv)
  python anomaly_check.py --test     # テストSlack (#03_ai活用)
  python anomaly_check.py --preview  # 通知なし（チェック結果のみ表示）
"""
from __future__ import annotations
import sys, yaml, requests
from datetime import date, timedelta
from pathlib import Path
from ga_client import load_env, run_report

CONFIG_PATH = Path(__file__).parent / "config.yaml"


def load_config() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return yaml.safe_load(f)


def get_28d_avg(property_id: str, base_date: date) -> float:
    """base_dateの28日前〜1日前の1日平均セッション数を返す"""
    start = (base_date - timedelta(days=28)).strftime("%Y-%m-%d")
    end   = (base_date - timedelta(days=1)).strftime("%Y-%m-%d")
    rows  = run_report(property_id, start, end, ["sessions"])
    if not rows:
        return 0.0
    return sum(float(r.get("sessions", 0)) for r in rows) / 28


def check_site(property_id: str, spike: float, drop: float) -> dict | None:
    yesterday = date.today() - timedelta(days=1)
    ystr = yesterday.strftime("%Y-%m-%d")

    rows = run_report(property_id, ystr, ystr, ["sessions"])
    yesterday_sess = float(rows[0].get("sessions", 0)) if rows else 0.0
    avg = get_28d_avg(property_id, yesterday)

    if avg < 10:
        return None  # 流量が少なすぎるサイトはスキップ

    ratio = yesterday_sess / avg if avg > 0 else 0.0
    if ratio >= spike:
        return {"type": "spike", "sessions": yesterday_sess, "avg": avg, "ratio": ratio}
    if ratio <= drop:
        return {"type": "drop",  "sessions": yesterday_sess, "avg": avg, "ratio": ratio}
    return None


def main():
    preview   = "--preview" in sys.argv
    test_mode = "--test" in sys.argv

    config    = load_config()
    env       = load_env()
    sites     = config.get("sites", [])
    anomaly   = config.get("anomaly", {})
    spike_thr = float(anomaly.get("spike_threshold", 2.5))
    drop_thr  = float(anomaly.get("drop_threshold",  0.4))
    slack_url = env.get("SLACK_WEBHOOK_TEST") if test_mode else env.get("SLACK_WEBHOOK_PROD")

    yesterday = date.today() - timedelta(days=1)
    alerts = []

    for site in sites:
        name    = site["name"]
        prop_id = str(site["property_id"])
        print(f"  チェック中: {name}")
        try:
            result = check_site(prop_id, spike_thr, drop_thr)
            if result:
                alerts.append({"name": name, **result})
                print(f"    ⚠️ {result['type'].upper()} (ratio: {result['ratio']:.2f}x)")
        except Exception as e:
            print(f"    [エラー] {name}: {e}")

    if not alerts:
        print("✅ 異常なし")
        return

    lines = [f"⚠️ *GA4 異常検知アラート — {yesterday.strftime('%Y-%m-%d')}*", ""]
    for a in alerts:
        icon  = "🔴" if a["type"] == "spike" else "🟡"
        label = "セッション急増" if a["type"] == "spike" else "セッション急減"
        lines.append(f"{icon} *{a['name']}*")
        lines.append(
            f"  {label}: {int(a['sessions']):,}件"
            f"（過去28日平均 {int(a['avg']):,}件の {a['ratio']*100:.0f}%）"
        )
    message = "\n".join(lines)
    print(message)

    if preview:
        print("\n（プレビューモード: Slack投稿スキップ）")
        return

    if slack_url:
        r = requests.post(slack_url, json={"text": message, "mrkdwn": True}, timeout=30)
        if r.status_code != 200:
            raise RuntimeError(f"Slack投稿失敗: {r.status_code} {r.text}")
        print(f"✅ Slack投稿完了（{len(alerts)}件のアラート）")


if __name__ == "__main__":
    main()
