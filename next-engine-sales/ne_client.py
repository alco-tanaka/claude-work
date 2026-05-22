"""ネクストエンジンAPIクライアント (リトライ付き)"""
import os, time, requests, json, base64
from pathlib import Path
from typing import Any, Dict

API_BASE = "https://api.next-engine.org"
ENV_PATH = Path(__file__).parent / ".env"

def load_env() -> Dict[str, str]:
    env = {}
    if ENV_PATH.exists():
        for line in ENV_PATH.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line: continue
            k, _, v = line.partition("=")
            env[k.strip()] = v.strip()
    for key in ["NE_ACCESS_TOKEN", "NE_REFRESH_TOKEN", "NE_CLIENT_ID",
                "NE_CLIENT_SECRET", "NE_REDIRECT_URI", "SLACK_WEBHOOK_URL",
                "SLACK_WEBHOOK_URL_WEBDIV"]:
        val = os.environ.get(key)
        if val:
            env[key] = val
    return env

def _update_github_secret(pat: str, repo: str, name: str, value: str) -> None:
    """GitHub Secretsを更新する（PyNaClで暗号化）"""
    try:
        from nacl import encoding, public
        r = requests.get(
            f"https://api.github.com/repos/{repo}/actions/secrets/public-key",
            headers={"Authorization": f"Bearer {pat}", "Accept": "application/vnd.github+json"},
            timeout=10
        )
        key_data = r.json()
        pub_key = public.PublicKey(key_data["key"].encode(), encoding.Base64Encoder())
        encrypted = base64.b64encode(public.SealedBox(pub_key).encrypt(value.encode())).decode()
        requests.put(
            f"https://api.github.com/repos/{repo}/actions/secrets/{name}",
            headers={"Authorization": f"Bearer {pat}", "Accept": "application/vnd.github+json"},
            json={"encrypted_value": encrypted, "key_id": key_data["key_id"]},
            timeout=10
        )
    except Exception as e:
        print(f"[警告] GitHub Secrets更新失敗 ({name}): {e}")

def save_tokens(access_token: str, refresh_token: str) -> None:
    # 同プロセス内の後続APIコールが更新済みトークンを使えるよう即時反映
    os.environ["NE_ACCESS_TOKEN"] = access_token
    os.environ["NE_REFRESH_TOKEN"] = refresh_token

    # GitHub Actions環境: 一時ファイルに書き出し（後続ステップでSecretsを更新）
    token_output = os.environ.get("NE_TOKEN_OUTPUT")
    if token_output:
        Path(token_output).write_text(
            json.dumps({"access_token": access_token, "refresh_token": refresh_token}),
            encoding="utf-8"
        )

    # ローカル環境: GitHub Secretsも同時更新（GH_PAT/.envに設定済みの場合）
    if not os.environ.get("GITHUB_ACTIONS"):
        env = load_env()
        gh_pat = env.get("GH_PAT")
        gh_repo = env.get("GH_REPO")
        if gh_pat and gh_repo:
            _update_github_secret(gh_pat, gh_repo, "NE_ACCESS_TOKEN", access_token)
            _update_github_secret(gh_pat, gh_repo, "NE_REFRESH_TOKEN", refresh_token)
            print("[情報] GitHub Secretsにトークンを同期しました")

    # ローカル環境: .envへの書き戻し
    if not ENV_PATH.exists():
        return
    last_err = None
    for attempt in range(5):
        try:
            lines = ENV_PATH.read_text(encoding="utf-8").splitlines()
            out, seen = [], set()
            for line in lines:
                s = line.strip()
                if "=" in s and not s.startswith("#"):
                    k = s.split("=", 1)[0].strip()
                    if k == "NE_ACCESS_TOKEN":
                        out.append(f"NE_ACCESS_TOKEN={access_token}"); seen.add(k); continue
                    if k == "NE_REFRESH_TOKEN":
                        out.append(f"NE_REFRESH_TOKEN={refresh_token}"); seen.add(k); continue
                out.append(line)
            if "NE_ACCESS_TOKEN" not in seen: out.append(f"NE_ACCESS_TOKEN={access_token}")
            if "NE_REFRESH_TOKEN" not in seen: out.append(f"NE_REFRESH_TOKEN={refresh_token}")
            ENV_PATH.write_text("\n".join(out) + "\n", encoding="utf-8")
            return
        except (PermissionError, OSError) as e:
            last_err = e
            time.sleep(2 ** attempt)
    raise RuntimeError(f"トークン保存失敗(5回リトライ後): {last_err}")

def ne_post(endpoint: str, params: Dict[str, Any] = None) -> Dict[str, Any]:
    env = load_env()
    body = {"access_token": env.get("NE_ACCESS_TOKEN", ""), "refresh_token": env.get("NE_REFRESH_TOKEN", "")}
    if params: body.update(params)
    resp = requests.post(f"{API_BASE}{endpoint}", data=body, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    if data.get("result") == "error" and data.get("code") == "002002":
        body["access_token"] = ""
        resp = requests.post(f"{API_BASE}{endpoint}", data=body, timeout=30)
        resp.raise_for_status()
        data = resp.json()

    na, nr = data.get("access_token"), data.get("refresh_token")
    if na and nr:
        save_tokens(na, nr)
    if data.get("result") == "error":
        raise RuntimeError(f"NE API error: {data.get('code')} {data.get('message')}")
    return data
