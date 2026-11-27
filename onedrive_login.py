"""One-time Microsoft device-code login for the worship OneDrive archive.

Run this locally, sign in with the dedicated church Microsoft account, and put
the printed refresh token into the Render environment.  The password is never
requested or stored.

Example:
    python3 onedrive_login.py --client-id <APPLICATION_CLIENT_ID> --check
"""

from __future__ import annotations

import argparse
import json
import os
import stat
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path

from src.backend.onedrive import (
    DEFAULT_ARCHIVE_SHARE_URL,
    DEFAULT_SCOPE,
    DEFAULT_TENANT,
    GRAPH_BASE,
    OneDriveArchive,
    OneDriveError,
    OneDriveSettings,
    parse_dates_from_filename,
)


TOKEN_HOST = "https://login.microsoftonline.com"
DEVICE_GRANT = "urn:ietf:params:oauth:grant-type:device_code"


def _post_form(url: str, data: dict, timeout: float = 30.0) -> dict:
    body = urllib.parse.urlencode(data).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=body,
        method="POST",
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
            "User-Agent": "WorshipPPT/1.0",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="replace")
        try:
            payload = json.loads(raw)
        except ValueError:
            payload = {"error": f"http_{exc.code}", "error_description": raw[:300]}
        payload.setdefault("_status", exc.code)
    return payload if isinstance(payload, dict) else {}


def device_login(
    client_id: str, tenant: str, scope: str
) -> dict:
    authority = f"{TOKEN_HOST}/{urllib.parse.quote(tenant, safe='')}"
    device = _post_form(
        f"{authority}/oauth2/v2.0/devicecode",
        {"client_id": client_id, "scope": scope},
    )
    if "device_code" not in device:
        raise SystemExit(
            "无法开始设备登录："
            + str(device.get("error_description") or device.get("error") or device)
        )

    print()
    print(device.get("message") or (
        f"请打开 {device.get('verification_uri')} 并输入代码 "
        f"{device.get('user_code')}"
    ))
    print()

    interval = max(1, int(device.get("interval") or 5))
    expires_in = max(60, int(device.get("expires_in") or 900))
    deadline = time.time() + expires_in

    while time.time() < deadline:
        time.sleep(interval)
        token = _post_form(
            f"{authority}/oauth2/v2.0/token",
            {
                "client_id": client_id,
                "grant_type": DEVICE_GRANT,
                "device_code": device["device_code"],
            },
        )
        error = token.get("error")
        if not error:
            if not token.get("refresh_token"):
                raise SystemExit(
                    "登录成功，但没有返回刷新令牌。请确认应用权限包含 offline_access。"
                )
            return token
        if error == "authorization_pending":
            continue
        if error == "slow_down":
            interval += 5
            continue
        if error == "expired_token":
            raise SystemExit("设备代码已过期，请重新运行登录脚本。")
        if error == "authorization_declined":
            raise SystemExit("登录已被拒绝。")
        raise SystemExit(
            "登录失败："
            + str(token.get("error_description") or error or "unknown")
        )
    raise SystemExit("等待登录超时，请重新运行登录脚本。")


def _delete_item(archive: OneDriveArchive, drive_id: str, item_id: str) -> None:
    token = archive.tokens.get_access_token()
    url = (
        f"{GRAPH_BASE}/drives/{urllib.parse.quote(drive_id, safe='')}"
        f"/items/{urllib.parse.quote(item_id, safe='')}"
    )
    request = urllib.request.Request(
        url,
        method="DELETE",
        headers={
            "Authorization": f"Bearer {token}",
            "User-Agent": "WorshipPPT/1.0",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            response.read()
    except urllib.error.HTTPError as exc:
        raise OneDriveError(f"删除测试文件失败（HTTP {exc.code}）。") from None
    except (urllib.error.URLError, OSError) as exc:
        raise OneDriveError(f"删除测试文件失败：{exc}") from None


def _write_env(path: Path, values: dict[str, str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "# Local OneDrive credentials for testing. Do not commit this file.",
        *[f"{key}={value}" for key, value in values.items()],
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")
    path.chmod(stat.S_IRUSR | stat.S_IWUSR)


def _print_check(
    archive: OneDriveArchive, check_date: date | None, check_write: bool
) -> None:
    root = archive.root()
    print(f"归档模式：{archive.settings.mode}")
    print(f"归档文件夹：{root.name or root.path}")
    listing = archive.list_files(force=True)
    print(f"找到 {len(listing.files)} 个 .pptx 文件。")
    for warning in listing.warnings:
        print(f"警告：{warning}")
    for item in listing.files[:20]:
        dates = parse_dates_from_filename(item.name)
        print(
            f"  - {item.name}  日期={dates[0].isoformat() if dates else '未知'}"
            f"  大小={item.size or 0}  修改={item.modified or '未知'}"
        )

    if check_date is not None:
        matches = [
            item
            for item in listing.files
            if check_date in parse_dates_from_filename(item.name)
        ]
        if matches:
            print(f"包含 {check_date.isoformat()} 的文件：")
            for item in matches:
                print(f"  * {item.name}")
        else:
            print(f"没有找到包含 {check_date.isoformat()} 的文件。")

    if check_write:
        test_name = f"_worshipppt_write_test_{int(time.time())}.pptx"
        # A minimal valid PPTX package.  It is removed immediately after upload.
        from io import BytesIO

        from pptx import Presentation

        buffer = BytesIO()
        Presentation().save(buffer)
        test_bytes = buffer.getvalue()
        try:
            uploaded = archive.graph.upload_file(root, test_name, test_bytes)
            print(f"写回测试成功：{uploaded.get('name') or test_name}")
            drive_id = str(
                (uploaded.get("parentReference") or {}).get("driveId")
                or root.drive_id
            )
            item_id = str(uploaded.get("id") or "")
            if item_id:
                _delete_item(archive, drive_id, item_id)
                print("写回测试文件已删除。")
            else:
                print("写回测试成功，但未返回 item id，请手动删除测试文件。")
        except OneDriveError as exc:
            print(f"写回测试失败：{exc}")
            print("如果文件夹是公开可编辑的，请检查应用权限和分享链接。")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--client-id",
        default=os.environ.get("ONEDRIVE_CLIENT_ID", ""),
        help="Application (client) ID from Entra App registrations.",
    )
    parser.add_argument(
        "--tenant",
        default=os.environ.get("ONEDRIVE_TENANT_ID", DEFAULT_TENANT),
        help="Tenant ID, common, organizations, or consumers (default: consumers).",
    )
    parser.add_argument(
        "--scope",
        default=os.environ.get("ONEDRIVE_SCOPE", DEFAULT_SCOPE),
        help="Delegated scopes (default: Files.ReadWrite offline_access).",
    )
    parser.add_argument(
        "--archive-share-url",
        default=os.environ.get("ONEDRIVE_ARCHIVE_URL", DEFAULT_ARCHIVE_SHARE_URL),
    )
    parser.add_argument(
        "--archive-path",
        default=os.environ.get("ONEDRIVE_ARCHIVE_PATH", ""),
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Resolve the archive and list files after login.",
    )
    parser.add_argument(
        "--check-date",
        default="",
        help="Optional YYYY-MM-DD to look for in filenames during --check.",
    )
    parser.add_argument(
        "--check-write",
        action="store_true",
        help="Upload and delete a tiny test PPTX (requires --check).",
    )
    parser.add_argument(
        "--write-env",
        default="",
        help="Write the credentials to a local env file (0600).",
    )
    parser.add_argument(
        "--no-print-token",
        action="store_true",
        help="Do not print the refresh token (still writes it with --write-env).",
    )
    args = parser.parse_args()

    client_id = args.client_id.strip()
    if not client_id:
        print("缺少 --client-id 或 ONEDRIVE_CLIENT_ID。", file=sys.stderr)
        return 2
    tenant = args.tenant.strip() or DEFAULT_TENANT
    scope = args.scope.strip() or DEFAULT_SCOPE
    check_date = None
    if args.check_date.strip():
        try:
            check_date = date.fromisoformat(args.check_date.strip())
        except ValueError:
            print("--check-date 必须是 YYYY-MM-DD。", file=sys.stderr)
            return 2

    token = device_login(client_id, tenant, scope)
    refresh_token = str(token["refresh_token"])
    issued_at = date.today().isoformat()
    values = {
        "ONEDRIVE_CLIENT_ID": client_id,
        "ONEDRIVE_TENANT_ID": tenant,
        "ONEDRIVE_REFRESH_TOKEN": refresh_token,
        "ONEDRIVE_TOKEN_ISSUED_AT": issued_at,
        "ONEDRIVE_SCOPE": scope,
        "ONEDRIVE_ARCHIVE_URL": args.archive_share_url.strip(),
    }
    if args.archive_path.strip():
        values["ONEDRIVE_ARCHIVE_PATH"] = args.archive_path.strip()

    print("登录成功。请把以下变量设置到 Render（不要把令牌发到聊天或提交到 Git）：")
    for key, value in values.items():
        if key == "ONEDRIVE_REFRESH_TOKEN" and args.no_print_token:
            print(f"{key}=<hidden; use --write-env>")
        else:
            print(f"{key}={value}")

    if args.write_env:
        path = Path(args.write_env).expanduser().resolve()
        _write_env(path, values)
        print(f"已写入本地环境文件：{path}（权限 0600，已被 .gitignore 忽略）")

    if args.check or args.check_write:
        settings = OneDriveSettings(
            client_id=client_id,
            refresh_token=refresh_token,
            tenant=tenant,
            scope=scope,
            archive_share_url=args.archive_share_url.strip() or DEFAULT_ARCHIVE_SHARE_URL,
            archive_path=args.archive_path.strip(),
        )
        archive = OneDriveArchive(settings)
        try:
            _print_check(archive, check_date, args.check_write)
        except OneDriveError as exc:
            print(f"归档检查失败：{exc}", file=sys.stderr)
            return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
