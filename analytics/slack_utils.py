"""Slack投稿共通処理（Bot Token経由のファイルアップロード）"""
from __future__ import annotations
import os
import requests


def post_slack(webhook_url: str, text: str, sheets_url: str = ""):
    attachments = []
    if sheets_url:
        attachments.append({"text": f"📋 <{sheets_url}|スプレッドシートを開く>", "color": "#27AE60"})
    r = requests.post(webhook_url, json={"text": text, "mrkdwn": True,
                                          "attachments": attachments}, timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"Slack投稿失敗: {r.status_code} {r.text}")


def _upload_one(bot_token: str, file_path: str) -> dict:
    filename = os.path.basename(file_path)
    filesize = os.path.getsize(file_path)
    headers = {"Authorization": f"Bearer {bot_token}"}

    r1 = requests.post(
        "https://slack.com/api/files.getUploadURLExternal",
        headers=headers,
        data={"filename": filename, "length": filesize},
        timeout=30,
    )
    resp1 = r1.json()
    if not resp1.get("ok"):
        raise RuntimeError(f"files.getUploadURLExternal失敗 ({filename}): {resp1}")
    upload_url, file_id = resp1["upload_url"], resp1["file_id"]

    with open(file_path, "rb") as f:
        r2 = requests.post(upload_url, files={"file": f}, timeout=120)
    if r2.status_code != 200:
        raise RuntimeError(f"ファイルアップロード失敗 ({filename}): {r2.status_code} {r2.text}")

    return {"id": file_id, "title": filename}


def post_slack_file(bot_token: str, channel_id: str, file_path: str, initial_comment: str = ""):
    """Bot Token経由でファイル1件をSlackチャンネルにアップロード投稿する"""
    return post_slack_files_multi(bot_token, channel_id, [file_path], initial_comment)


def post_slack_files_multi(bot_token: str, channel_id: str, file_paths: list[str], initial_comment: str = ""):
    """複数ファイルを1つのSlackメッセージにまとめてアップロード投稿する"""
    headers = {"Authorization": f"Bearer {bot_token}"}
    files_payload = [_upload_one(bot_token, fp) for fp in file_paths]

    payload = {"files": files_payload, "channel_id": channel_id}
    if initial_comment:
        payload["initial_comment"] = initial_comment
    r3 = requests.post(
        "https://slack.com/api/files.completeUploadExternal",
        headers={**headers, "Content-Type": "application/json"},
        json=payload,
        timeout=60,
    )
    resp3 = r3.json()
    if not resp3.get("ok"):
        raise RuntimeError(f"files.completeUploadExternal失敗: {resp3}")
    return resp3
