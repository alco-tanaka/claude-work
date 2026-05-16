  """ネクストエンジンAPIクライアント (リトライ付き)"""
  import os, time, requests
  from pathlib import Path
  from typing import Any, Dict

  API_BASE = "https://api.next-engine.org"
  ENV_PATH = Path(__file__).parent / ".env"

  def load_env() -> Dict[str, str]:
      env = {}
      if ENV_PATH.exists():
          for line in ENV_PATH.read_text(encoding="utf-8").splitlines():
              line = line.strip()
              if not line or line.startswith("#") or "=" not in line: continue
              k, _, v = line.partition("=")
              env[k.strip()] = v.strip()
      for key in ["NE_ACCESS_TOKEN", "NE_REFRESH_TOKEN", "NE_CLIENT_ID",
                  "NE_CLIENT_SECRET", "NE_REDIRECT_URI", "SLACK_WEBHOOK_URL"]:
          val = os.environ.get(key)
          if val:
              env[key] = val
      return env

  def save_tokens(access_token: str, refresh_token: str) -> None:
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
      if na and nr and (na != env.get("NE_ACCESS_TOKEN") or nr != env.get("NE_REFRESH_TOKEN")):
          save_tokens(na, nr)
      if data.get("result") == "error":
          raise RuntimeError(f"NE API error: {data.get('code')} {data.get('message')}")
      return data
