"""GA4 月次レポート生成 → pptx + Googleスプレッドシート + Google Drive + Slack投稿
Usage:
  python monthly_report.py               # 先月分・本番Slack (#01_webdiv)
  python monthly_report.py --preview     # ファイル生成のみ・Slack投稿なし
  python monthly_report.py --test        # テストSlack (#03_ai活用) に投稿
  python monthly_report.py --month 2026-04  # 指定月
"""
from __future__ import annotations
import sys, os, yaml, requests
from datetime import date, timedelta
from pathlib import Path
from ga_client import load_env, fetch_site_data, get_gspread_client, get_drive_service
from pptx_builder import build_report

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


def upload_to_drive(file_path: str) -> str:
    env = load_env()
    folder_id = env.get("DRIVE_FOLDER_ID")
    if not folder_id:
        print("[警告] DRIVE_FOLDER_IDが未設定。Driveアップロードをスキップ。")
        return ""

    from googleapiclient.http import MediaFileUpload
    service = get_drive_service()
    file_name = Path(file_path).name
    mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

    # 同名ファイルがあれば上書き
    existing = service.files().list(
        q=f"name='{file_name}' and '{folder_id}' in parents and trashed=false",
        fields="files(id)"
    ).execute().get("files", [])

    media = MediaFileUpload(file_path, mimetype=mime, resumable=True)
    if existing:
        file_id = existing[0]["id"]
        service.files().update(fileId=file_id, media_body=media).execute()
    else:
        meta = {"name": file_name, "parents": [folder_id], "mimeType": mime}
        created = service.files().create(body=meta, media_body=media, fields="id").execute()
        file_id = created["id"]

    service.permissions().create(
        fileId=file_id,
        body={"type": "anyone", "role": "reader"}
    ).execute()
    drive_url = f"https://drive.google.com/file/d/{file_id}/view"
    print(f"✅ Driveアップロード完了: {drive_url}")
    return drive_url


def post_slack(webhook_url: str, text: str, drive_url: str = "", sheets_url: str = ""):
    attachments = []
    if drive_url:
        attachments.append({"text": f"📊 <{drive_url}|PowerPointを開く>", "color": "#1A376C"})
    if sheets_url:
        attachments.append({"text": f"📋 <{sheets_url}|スプレッドシートを開く>", "color": "#27AE60"})
    r = requests.post(webhook_url, json={"text": text, "mrkdwn": True,
                                          "attachments": attachments}, timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"Slack投稿失敗: {r.status_code} {r.text}")


def main():
    preview  = "--preview" in sys.argv
    test_mode = "--test" in sys.argv
    year, month = target_month(sys.argv)
    print(f"対象月: {year}年{month}月")

    config = load_config()
    env = load_env()
    sites = config.get("sites", [])
    slack_url = env.get("SLACK_WEBHOOK_TEST") if test_mode else env.get("SLACK_WEBHOOK_PROD")

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

    # pptx生成
    output_dir = "/tmp" if os.environ.get("GITHUB_ACTIONS") else config.get("output_dir", ".")
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"GA4_月次レポート_{year}{month:02d}.pptx")
    build_report(sites_data, year, month, output_path)

    if preview:
        print("\n（プレビューモード: Slack投稿・Drive/Sheets書き込みスキップ）")
        return

    sheets_url = write_to_sheets(sites_data, year, month)
    drive_url  = upload_to_drive(output_path)

    if slack_url:
        ok = sum(1 for sd in sites_data if not sd.get("error"))
        ng = len(sites_data) - ok
        text = (
            f"*GA4 月次レポート — {year}年{month}月*\n"
            f"全{len(sites_data)}サイトのアクセスデータをまとめました。\n"
            f"✅ 取得成功: {ok}サイト"
            + (f"  ⚠️ エラー: {ng}サイト" if ng else "")
        )
        post_slack(slack_url, text, drive_url, sheets_url)
        print("✅ Slack投稿完了")


if __name__ == "__main__":
    main()
