"""Renew the OneDrive refresh token and persist it to Render.

The GitHub Actions workflow provides RENDER_API_KEY and RENDER_SERVICE_ID.
The OneDrive client ID, tenant, scope, and current refresh token are read from
the Render service environment; the new refresh token is written back to Render
with the single-variable API. No Microsoft token is stored in GitHub.
"""

from __future__ import annotations

import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import date

RENDER_API = "https://api.render.com/v1"
TOKEN_HOST = "https://login.microsoftonline.com"
DEFAULT_SCOPE = "Files.ReadWrite offline_access"


class RenewError(RuntimeError):
    """A safe error message; never include tokens or secret values."""


def _open_json(
    method: str,
    url: str,
    *,
    headers: dict | None = None,
    form: dict | None = None,
    body: dict | None = None,
    opener=None,
    timeout: float = 30.0,
):
    data = None
    request_headers = {
        "Accept": "application/json",
        "User-Agent": "WorshipPPT/1.0",
    }
    if form is not None:
        data = urllib.parse.urlencode(form).encode("utf-8")
        request_headers["Content-Type"] = "application/x-www-form-urlencoded"
    elif body is not None:
        data = json.dumps(body).encode("utf-8")
        request_headers["Content-Type"] = "application/json"
    request_headers.update(headers or {})

    request = urllib.request.Request(
        url, data=data, headers=request_headers, method=method
    )
    client = opener or urllib.request.build_opener()
    try:
        with client.open(request, timeout=timeout) as response:
            raw = response.read()
    except urllib.error.HTTPError as exc:
        detail = ""
        try:
            payload = json.loads(exc.read().decode("utf-8", errors="replace"))
            if isinstance(payload, dict):
                detail = str(payload.get("error") or payload.get("message") or "")
        except ValueError:
            pass
        raise RenewError(
            f"HTTP {exc.code}" + (f" ({detail[:120]})" if detail else "")
        ) from None
    except (urllib.error.URLError, OSError) as exc:
        raise RenewError(f"网络请求失败：{exc}") from None
    if not raw:
        return {}
    try:
        return json.loads(raw.decode("utf-8"))
    except ValueError:
        raise RenewError("服务返回了无法解析的 JSON。") from None


def render_get_env(
    api_key: str,
    service_id: str,
    key: str,
    *,
    opener=None,
    allow_missing: bool = False,
) -> str | None:
    url = (
        f"{RENDER_API}/services/{urllib.parse.quote(service_id, safe='')}"
        f"/env-vars/{urllib.parse.quote(key, safe='')}"
    )
    try:
        payload = _open_json(
            "GET",
            url,
            headers={"Authorization": f"Bearer {api_key}"},
            opener=opener,
        )
    except RenewError as exc:
        if allow_missing and "HTTP 404" in str(exc):
            return None
        raise
    if not isinstance(payload, dict) or "value" not in payload:
        raise RenewError(f"Render 没有返回 {key} 的值。")
    value = payload.get("value")
    return None if value is None else str(value)


def render_set_env(
    api_key: str,
    service_id: str,
    key: str,
    value: str,
    *,
    opener=None,
) -> None:
    url = (
        f"{RENDER_API}/services/{urllib.parse.quote(service_id, safe='')}"
        f"/env-vars/{urllib.parse.quote(key, safe='')}"
    )
    _open_json(
        "PUT",
        url,
        headers={"Authorization": f"Bearer {api_key}"},
        body={"value": value},
        opener=opener,
    )


def refresh_refresh_token(
    client_id: str,
    tenant: str,
    scope: str,
    refresh_token: str,
    *,
    opener=None,
) -> str:
    authority = f"{TOKEN_HOST}/{urllib.parse.quote(tenant, safe='')}"
    payload = _open_json(
        "POST",
        f"{authority}/oauth2/v2.0/token",
        form={
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": scope,
        },
        opener=opener,
    )
    if not isinstance(payload, dict):
        raise RenewError("Microsoft 返回了异常响应。")
    error = payload.get("error")
    if error:
        if error == "invalid_grant":
            raise RenewError(
                "刷新令牌已过期或被撤销。请重新运行 onedrive_login.py，"
                "并把新的刷新令牌更新到 Render。"
            )
        raise RenewError(f"Microsoft 登录失败（{error}）。")
    new_token = payload.get("refresh_token")
    if not new_token:
        raise RenewError(
            "Microsoft 没有返回新的刷新令牌；请确认 scope 含 offline_access。"
        )
    return str(new_token)


def main() -> int:
    api_key = (os.environ.get("RENDER_API_KEY") or "").strip()
    service_id = (os.environ.get("RENDER_SERVICE_ID") or "").strip()
    if not api_key or not service_id:
        print("缺少 RENDER_API_KEY 或 RENDER_SERVICE_ID。", file=sys.stderr)
        return 2

    try:
        client_id = render_get_env(api_key, service_id, "ONEDRIVE_CLIENT_ID")
        tenant = render_get_env(api_key, service_id, "ONEDRIVE_TENANT_ID") or "consumers"
        current_token = render_get_env(
            api_key, service_id, "ONEDRIVE_REFRESH_TOKEN"
        )
        scope = (
            render_get_env(
                api_key,
                service_id,
                "ONEDRIVE_SCOPE",
                allow_missing=True,
            )
            or DEFAULT_SCOPE
        )
        if not client_id or not current_token:
            print(
                "Render 中缺少 ONEDRIVE_CLIENT_ID 或 ONEDRIVE_REFRESH_TOKEN。",
                file=sys.stderr,
            )
            return 2

        new_token = refresh_refresh_token(
            client_id, tenant, scope, current_token
        )
        render_set_env(
            api_key, service_id, "ONEDRIVE_REFRESH_TOKEN", new_token
        )
        render_set_env(
            api_key,
            service_id,
            "ONEDRIVE_TOKEN_ISSUED_AT",
            date.today().isoformat(),
        )
        print(
            "OneDrive 刷新令牌已更新到 Render，签发日期："
            f"{date.today().isoformat()}。"
        )
        return 0
    except RenewError as exc:
        print(f"续期失败：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
