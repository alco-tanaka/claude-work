"""GA4 月次レポート生成（全サイトまとめ） → pptx（data/outputs/ga4-report/YYYYMM/）+ Googleスプレッドシート
社内記録用。Slackへの投稿は行わない（サイト別詳細レポートの投稿は post_site_reports.py を使用）。
Usage:
  python monthly_report.py               # 先月分
  python monthly_report.py --preview     # ファイル生成のみ・Sheets書き込みなし
  python monthly_report.py --month 2026-04  # 指定月
"""
from __future__ import annotations
import sys, os, yaml
from datetime import date, timedelta
from pathlib import Path
from ga_client import load_env, fetch_site_data, get_gspread_client
from pptx_builder import build_report

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

CONFIG_PATH = Path(__file__).parent / "config.yaml"


def load_config() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return yaml.safe_load(f)


def target_month(args: list[str]) -> tuple[int, int]:
    for i, a in enumerate(args):
        if a == "--month" and i + 1 < len(args):
            y, m = args[i + 1].split("-")
            return int(y), int(m)
    first = date.today().replace(day=1)
    prev = first - timedelta(days=1)
    return prev.year, prev.month


def write_to_sheets(sites_data: list[dict], year: int, month: int) -> str:
    env = load_env()
    spreadsheet_id = env.get("SPREADSHEET_ID")
    if not spreadsheet_id:
        print("[警告] SPREADSHEET_IDが未設定。スプレッドシート書き込みをスキップ。")
        return ""

    gc = get_gspread_client()
    sh = gc.open_by_key(spreadsheet_id)
    sheet_name = f"{year}{month:02d}"

    try:
        ws = sh.worksheet(sheet_name)
        ws.clear()
    except Exception:
        ws = sh.add_worksheet(title=sheet_name, rows=50, cols=20)

    headers = ["サイト名", "セッション", "ユーザー", "新規ユーザー", "PV",
               "平均滞在時間(秒)", "エンゲージメント率(%)",
               "前月セッション", "前月ユーザー", "前月PV"]
    ws.update("A1", [headers])

    rows = []
    for sd in sites_data:
        if sd.get("error"):
            rows.append([sd["name"], "エラー"] + [""] * (len(headers) - 2))
            continue
        t = sd["data"]["totals"]
        p = sd["data"]["prev_totals"]
        rows.append([
            sd["name"],
            int(float(t["sessions"])),
            int(float(t["totalUsers"])),
            int(float(t["newUsers"])),
            int(float(t["screenPageViews"])),
            round(float(t["averageSessionDuration"]), 1),
            round(float(t["engagementRate"]) * 100, 2),
            int(float(p["sessions"])),
            int(float(p["totalUsers"])),
            int(float(p["screenPageViews"])),
        ])
    ws.update("A2", rows)
    print(f"✅ スプレッドシート更新: シート '{sheet_name}'")
    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"


def main():
    preview  = "--preview" in sys.argv
    year, month = target_month(sys.argv)
    print(f"対象月: {year}年{month}月")

    config = load_config()
    sites = config.get("sites", [])

    # GA4データ取得
    sites_data = []
    for site in sites:
        name = site["name"]
        prop_id = str(site["property_id"])
        print(f"  取得中: {name}")
        try:
            sites_data.append({"name": name, "data": fetch_site_data(prop_id, year, month)})
        except Exception as e:
            print(f"  [エラー] {name}: {e}")
            sites_data.append({"name": name, "error": str(e)})

    # pptx生成（月ごとにサブフォルダへ保存）
    base_dir = "/tmp" if os.environ.get("GITHUB_ACTIONS") else config.get("output_dir", ".")
    output_dir = os.path.join(base_dir, f"{year}{month:02d}")
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"GA4_月次レポート_{year}{month:02d}.pptx")
    build_report(sites_data, year, month, output_path)

    if preview:
        print("\n（プレビューモード: Sheets書き込みスキップ）")
        return

    write_to_sheets(sites_data, year, month)


if __name__ == "__main__":
    main()
