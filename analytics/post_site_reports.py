"""GA4 サイト別詳細レポートを生成し、Slackへ1メッセージにまとめて投稿する
Usage:
  python post_site_reports.py                 # 先月分・本番Slack (#01_webdiv)
  python post_site_reports.py --test           # テストSlack (#03_ai活用) に投稿
  python post_site_reports.py --month 2026-06  # 指定月
  python post_site_reports.py --no-generate    # 既存ファイルのみ使用（再生成しない）
"""
from __future__ import annotations
import sys, os
from datetime import date, timedelta
from pathlib import Path
import yaml

sys.path.insert(0, str(Path(__file__).parent))

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

from ga_client import load_env
from detail_report import build_detail_report, fetch_benchmark_metrics
from slack_utils import post_slack_files_multi

CONFIG_PATH = Path(__file__).parent / "config.yaml"

SLACK_CHANNEL_PROD = "G014LJM32FM"  # #01_webdiv
SLACK_CHANNEL_TEST = "C0B2Z3QRUBB"  # #03_ai活用_alcoどじょうunite


def target_month(args: list[str]) -> tuple[int, int]:
    for i, a in enumerate(args):
        if a == "--month" and i + 1 < len(args):
            y, m = args[i + 1].split("-")
            return int(y), int(m)
    first = date.today().replace(day=1)
    prev = first - timedelta(days=1)
    return prev.year, prev.month


def main():
    test_mode = "--test" in sys.argv
    skip_generate = "--no-generate" in sys.argv
    year, month = target_month(sys.argv)
    print(f"対象月: {year}年{month}月")

    with open(CONFIG_PATH, encoding="utf-8") as f:
        config = yaml.safe_load(f)
    sites = config["sites"]
    base_dir = "/tmp/ga4-report" if os.environ.get("GITHUB_ACTIONS") else config.get("output_dir", ".")
    output_dir = os.path.join(base_dir, f"{year}{month:02d}")
    os.makedirs(output_dir, exist_ok=True)

    file_paths = []
    if skip_generate:
        for site in sites:
            fpath = os.path.join(output_dir, f"{site['name']}_詳細レポート_{year}{month:02d}.pptx")
            if os.path.exists(fpath):
                file_paths.append(fpath)
            else:
                print(f"[警告] ファイルが見つかりません（スキップ）: {fpath}")
    else:
        print("ベンチマークデータ取得中（他サイト比較用）...")
        benchmark = fetch_benchmark_metrics(sites, year, month)
        for site in sites:
            try:
                out_path = build_detail_report(
                    site["name"], site["property_id"], year, month, output_dir,
                    search_console_url=site.get("search_console_url"),
                    has_ecommerce=site.get("has_ecommerce", False),
                    benchmark=benchmark,
                )
                file_paths.append(out_path)
            except Exception as e:
                print(f"[エラー] {site['name']}: {e}")

    if not file_paths:
        print("投稿対象ファイルがありません。終了します。")
        return

    env = load_env()
    bot_token = env.get("SLACK_BOT_TOKEN")
    if not bot_token:
        print("[警告] SLACK_BOT_TOKENが未設定。Slack投稿をスキップ。")
        return
    channel_id = SLACK_CHANNEL_TEST if test_mode else SLACK_CHANNEL_PROD

    text = (
        f"*GA4 月次レポート（サイト別） — {year}年{month}月*\n"
        f"全{len(file_paths)}サイトの詳細レポートです。"
    )
    post_slack_files_multi(bot_token, channel_id, file_paths, initial_comment=text)
    print("✅ Slackファイル投稿完了")


if __name__ == "__main__":
    main()
