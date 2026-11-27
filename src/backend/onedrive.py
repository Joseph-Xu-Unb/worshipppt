"""Microsoft Graph access for the church OneDrive worship archive.

The archive is a public "anyone with the link can edit" folder owned by another
Microsoft account.  A dedicated personal Microsoft account signs in once with
the device-code flow; the server keeps only the refresh token and uses delegated
Graph permissions to read last week's deck, write this week's deck back, and
create a view link for the new file.

No password is ever stored.  The retrieval side only issues GET requests.  The
write side only uploads the generated deck into the configured archive folder
and creates a share link for that file; it never deletes or renames files.
"""

from __future__ import annotations

import base64
import json
import os
import re
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import deque
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

from src.backend.pptx_inputs import (
    MAX_LAST_WEEK_PPTX_BYTES,
    is_allowed_onedrive_url,
    validate_pptx_bytes,
)


GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_HOST = "https://login.microsoftonline.com"
DEFAULT_ARCHIVE_SHARE_URL = (
    "https://1drv.ms/f/c/1b40ff3184add23e/"
    "IgBvMYWWe1v8RpbhkkjUr142Aar3qCdtzJEPobKBbub6zMU?e=lFGb49"
)
DEFAULT_SCOPE = "Files.ReadWrite offline_access"
DEFAULT_TENANT = "consumers"
TOKEN_RENEWAL_WARN_DAYS = 75
TOKEN_RENEWAL_EXPIRED_DAYS = 90
MAX_ARCHIVE_DEPTH = 4
MAX_ARCHIVE_CHILDREN = 5000
MAX_CANDIDATES = 10
SIMPLE_UPLOAD_THRESHOLD = 4 * 1024 * 1024
UPLOAD_CHUNK_SIZE = 16 * 327680  # 5 MiB, a multiple of the required 320 KiB
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
REDIRECT_CODES = {301, 302, 303, 307, 308}
UPLOAD_HOST_SUFFIXES = (
    "microsoftpersonalcontent.com",
    "files.1drv.com",
    "onedrive.live.com",
    "sharepoint.com",
    "graph.microsoft.com",
)

_DATE_PATTERNS = (
    re.compile(r"(?<!\d)(\d{4})-(\d{1,2})-(\d{1,2})(?!\d)"),
    re.compile(r"(?<!\d)(\d{4})[._](\d{1,2})[._](\d{1,2})(?!\d)"),
    re.compile(r"(?<!\d)(\d{4})年(\d{1,2})月(\d{1,2})日"),
    re.compile(r"(?<!\d)(\d{4})(\d{2})(\d{2})(?!\d)"),
)


class OneDriveError(RuntimeError):
    """Base error with a user-safe message (never includes tokens or URLs)."""


class OneDriveAuthError(OneDriveError):
    pass


class OneDriveLookupError(OneDriveError):
    pass


class OneDriveUploadError(OneDriveError):
    pass


class GraphHTTPError(OneDriveError):
    def __init__(self, status: int, code: str, message: str):
        self.status = status
        self.code = code
        self.message = message
        super().__init__(f"Microsoft Graph 请求失败（{status}）：{code} {message}".strip())


@dataclass(frozen=True)
class OneDriveSettings:
    client_id: str
    refresh_token: str = field(repr=False)
    tenant: str = DEFAULT_TENANT
    scope: str = DEFAULT_SCOPE
    archive_share_url: str = DEFAULT_ARCHIVE_SHARE_URL
    archive_path: str = ""
    token_issued_at: date | None = None
    auto_enabled: bool = True
    upload_enabled: bool = True
    timeout: float = 30.0

    @property
    def mode(self) -> str:
        return "path" if self.archive_path.strip() else "share"

    @property
    def can_write(self) -> bool:
        lowered = self.scope.lower()
        return "files.readwrite" in lowered

    @classmethod
    def from_env(cls, env: dict | None = None) -> "OneDriveSettings | None":
        values = os.environ if env is None else env
        client_id = (values.get("ONEDRIVE_CLIENT_ID") or "").strip()
        refresh_token = (values.get("ONEDRIVE_REFRESH_TOKEN") or "").strip()
        if not client_id or not refresh_token:
            return None

        tenant = (values.get("ONEDRIVE_TENANT_ID") or DEFAULT_TENANT).strip()
        scope = (values.get("ONEDRIVE_SCOPE") or DEFAULT_SCOPE).strip()
        share_url = (
            values.get("ONEDRIVE_ARCHIVE_URL") or DEFAULT_ARCHIVE_SHARE_URL
        ).strip()
        archive_path = (values.get("ONEDRIVE_ARCHIVE_PATH") or "").strip()
        issued_raw = (values.get("ONEDRIVE_TOKEN_ISSUED_AT") or "").strip()
        issued_at = None
        if issued_raw:
            try:
                issued_at = date.fromisoformat(issued_raw)
            except ValueError:
                issued_at = None

        return cls(
            client_id=client_id,
            refresh_token=refresh_token,
            tenant=tenant or DEFAULT_TENANT,
            scope=scope or DEFAULT_SCOPE,
            archive_share_url=share_url or DEFAULT_ARCHIVE_SHARE_URL,
            archive_path=archive_path,
            token_issued_at=issued_at,
            auto_enabled=_env_bool(values, "ONEDRIVE_AUTO_ENABLED", True),
            upload_enabled=_env_bool(values, "ONEDRIVE_UPLOAD_ENABLED", True),
            timeout=_env_float(values, "ONEDRIVE_TIMEOUT", 30.0),
        )

    def token_age_days(self, today: date | None = None) -> int | None:
        if self.token_issued_at is None:
            return None
        return ((today or date.today()) - self.token_issued_at).days

    def token_status(self, today: date | None = None) -> str:
        age = self.token_age_days(today)
        if age is None:
            return "unknown"
        if age >= TOKEN_RENEWAL_EXPIRED_DAYS:
            return "expired"
        if age >= TOKEN_RENEWAL_WARN_DAYS:
            return "renew_soon"
        return "ok"


def _env_bool(values, key: str, default: bool) -> bool:
    raw = values.get(key)
    if raw is None or not str(raw).strip():
        return default
    return str(raw).strip().lower() not in {"0", "false", "no", "off"}


def _env_float(values, key: str, default: float) -> float:
    raw = values.get(key)
    if raw is None or not str(raw).strip():
        return default
    try:
        return max(5.0, float(raw))
    except ValueError:
        return default


@dataclass(frozen=True)
class DriveItemRef:
    drive_id: str
    item_id: str
    name: str
    is_folder: bool
    path: str = ""
    web_url: str = ""
    modified: str = ""
    size: int | None = None
    share_id: str = ""


@dataclass(frozen=True)
class OneDriveFile:
    drive_id: str
    item_id: str
    name: str
    modified: str = ""
    size: int | None = None
    web_url: str = ""
    path: str = ""
    share_id: str = ""

    def as_candidate(self) -> dict:
        dates = parse_dates_from_filename(self.name)
        return {
            "name": self.name,
            "modified": self.modified,
            "size": self.size,
            "date": dates[0].isoformat() if dates else "",
            "path": self.path,
        }


@dataclass
class TreeListing:
    files: list[OneDriveFile] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    truncated: bool = False


@dataclass
class ArchiveLookup:
    file: OneDriveFile | None = None
    target: date | None = None
    candidates: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    error: str = ""


@dataclass
class AutoLastWeekResult:
    content: bytes | None = None
    filename: str = ""
    attempted: bool = False
    error: str = ""
    candidates: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def previous_sunday(service_date: date) -> date:
    """Return the Sunday strictly before ``service_date``."""
    days = (service_date.weekday() + 1) % 7
    return service_date - timedelta(days=days or 7)


def parse_dates_from_filename(name: str) -> list[date]:
    """Extract valid calendar dates embedded in a filename."""
    found: list[date] = []
    for pattern in _DATE_PATTERNS:
        for match in pattern.finditer(name):
            try:
                parsed = date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
            except ValueError:
                continue
            if parsed not in found:
                found.append(parsed)
    return found


def is_pptx_filename(name: str) -> bool:
    base = name.rsplit("/", 1)[-1].rsplit("\\", 1)[-1].strip()
    return base.lower().endswith(".pptx") and not base.startswith("~$")


def is_safe_upload_name(name: str) -> bool:
    if not name or name != name.strip():
        return False
    if name.startswith("~$"):
        return False
    if any(char in name for char in ("/", "\\", "\x00", "?", "#")):
        return False
    if ".." in name or not name.lower().endswith(".pptx"):
        return False
    return True


def encode_share_url(url: str) -> str:
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii")
    return "u!" + encoded.rstrip("=")


class _NoRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def canonical_share_url(url: str, *, timeout: float = 20.0) -> str:
    """Resolve a 1drv.ms short link to the canonical OneDrive sharing URL."""
    if not is_allowed_onedrive_url(url):
        raise OneDriveError("归档链接必须是 https OneDrive/SharePoint 链接。")
    host = (urllib.parse.urlparse(url).hostname or "").lower()
    if not (host == "1drv.ms" or host.endswith(".1drv.ms")):
        return url

    request = urllib.request.Request(url, headers={"User-Agent": "WorshipPPT/1.0"})
    opener = urllib.request.build_opener(_NoRedirectHandler())
    try:
        with opener.open(request, timeout=timeout):
            return url
    except urllib.error.HTTPError as exc:
        if exc.code in REDIRECT_CODES:
            location = exc.headers.get("Location")
            if location:
                candidate = urllib.parse.urljoin(url, location)
                if is_allowed_onedrive_url(candidate):
                    # "redeem" is a single-use hand-off parameter; Graph only
                    # needs the stable sharing token (the "e" query value).
                    parts = urllib.parse.urlsplit(candidate)
                    query = [
                        (key, value)
                        for key, value in urllib.parse.parse_qsl(
                            parts.query, keep_blank_values=True
                        )
                        if key.lower() != "redeem"
                    ]
                    return urllib.parse.urlunsplit(
                        (
                            parts.scheme,
                            parts.netloc,
                            parts.path,
                            urllib.parse.urlencode(query),
                            parts.fragment,
                        )
                    )
                raise OneDriveError(
                    "归档短链接重定向到了非 OneDrive 主机，已拒绝。"
                ) from None
        return url
    except (urllib.error.URLError, OSError):
        return url


def _sanitize_message(message: str, *, limit: int = 300) -> str:
    text = re.sub(r"https?://\S+", "<url>", message or "")
    text = re.sub(r"\s+", " ", text).strip()
    return text[:limit]


def _graph_error(status: int, body: bytes) -> GraphHTTPError:
    code = ""
    message = ""
    try:
        payload = json.loads(body.decode("utf-8", errors="replace"))
        error = payload.get("error") or {}
        if isinstance(error, dict):
            code = str(error.get("code") or "")
            message = str(error.get("message") or "")
    except (ValueError, AttributeError):
        message = body.decode("utf-8", errors="replace")
    return GraphHTTPError(status, code or "unknown", _sanitize_message(message))


class TokenProvider:
    def __init__(self, settings: OneDriveSettings, opener=None):
        self._settings = settings
        self._opener = opener or urllib.request.build_opener()
        self._lock = threading.Lock()
        self._access_token: str | None = None
        self._expires_at = 0.0
        self._refresh_token = settings.refresh_token

    @property
    def token_url(self) -> str:
        tenant = urllib.parse.quote(self._settings.tenant.strip() or DEFAULT_TENANT, safe="")
        return f"{TOKEN_HOST}/{tenant}/oauth2/v2.0/token"

    def get_access_token(self, force: bool = False) -> str:
        with self._lock:
            if (
                not force
                and self._access_token
                and time.time() < self._expires_at
            ):
                return self._access_token
            return self._refresh_locked()

    def invalidate(self) -> None:
        with self._lock:
            self._access_token = None
            self._expires_at = 0.0

    def _refresh_locked(self) -> str:
        body = urllib.parse.urlencode(
            {
                "client_id": self._settings.client_id,
                "grant_type": "refresh_token",
                "refresh_token": self._refresh_token,
                "scope": self._settings.scope,
            }
        ).encode("utf-8")
        request = urllib.request.Request(
            self.token_url,
            data=body,
            method="POST",
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "Accept": "application/json",
                "User-Agent": "WorshipPPT/1.0",
            },
        )
        try:
            with self._opener.open(request, timeout=self._settings.timeout) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            raw = exc.read()
            error_code = ""
            try:
                error_payload = json.loads(raw.decode("utf-8", errors="replace"))
                error_code = str(error_payload.get("error") or "")
            except ValueError:
                pass
            if error_code == "invalid_grant":
                raise OneDriveAuthError(
                    "OneDrive 登录已过期或被撤销，请重新运行 onedrive_login.py，"
                    "并更新 ONEDRIVE_REFRESH_TOKEN。"
                ) from None
            raise OneDriveAuthError(
                f"OneDrive 登录失败（{error_code or exc.code}），"
                "请重新运行 onedrive_login.py。"
            ) from None
        except (urllib.error.URLError, OSError) as exc:
            raise OneDriveAuthError(
                f"无法连接 Microsoft 登录服务：{_sanitize_message(str(exc))}"
            ) from None

        access_token = payload.get("access_token")
        if not access_token:
            raise OneDriveAuthError("Microsoft 登录服务没有返回访问令牌。")
        new_refresh = payload.get("refresh_token")
        if new_refresh:
            self._refresh_token = new_refresh
        try:
            expires_in = int(payload.get("expires_in") or 3600)
        except (TypeError, ValueError):
            expires_in = 3600
        self._access_token = access_token
        self._expires_at = time.time() + max(expires_in - 60, 30)
        return access_token


class GraphClient:
    def __init__(
        self,
        settings: OneDriveSettings,
        tokens: TokenProvider | None = None,
        opener=None,
    ):
        self._settings = settings
        self._tokens = tokens or TokenProvider(settings)
        self._opener = opener or urllib.request.build_opener()

    @property
    def tokens(self) -> TokenProvider:
        return self._tokens

    def _request(
        self,
        method: str,
        url: str,
        *,
        data: bytes | None = None,
        headers: dict | None = None,
        timeout: float | None = None,
        retry_auth: bool = True,
    ) -> bytes:
        token = self._tokens.get_access_token()
        request_headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "User-Agent": "WorshipPPT/1.0",
        }
        request_headers.update(headers or {})
        request = urllib.request.Request(
            url, data=data, headers=request_headers, method=method
        )
        try:
            with self._opener.open(
                request, timeout=timeout or self._settings.timeout
            ) as response:
                return response.read()
        except urllib.error.HTTPError as exc:
            if exc.code == 401 and retry_auth:
                self._tokens.invalidate()
                return self._request(
                    method,
                    url,
                    data=data,
                    headers=headers,
                    timeout=timeout,
                    retry_auth=False,
                )
            raise _graph_error(exc.code, exc.read()) from None
        except (urllib.error.URLError, OSError) as exc:
            raise OneDriveError(
                f"无法连接 Microsoft Graph：{_sanitize_message(str(exc))}"
            ) from None

    def _request_json(
        self,
        method: str,
        url: str,
        *,
        data: bytes | None = None,
        headers: dict | None = None,
    ) -> dict:
        raw = self._request(method, url, data=data, headers=headers)
        if not raw:
            return {}
        try:
            payload = json.loads(raw.decode("utf-8"))
        except ValueError:
            raise OneDriveError("Microsoft Graph 返回了无法解析的响应。") from None
        return payload if isinstance(payload, dict) else {}

    def get_json(self, path: str, *, headers: dict | None = None) -> dict:
        return self._request_json("GET", GRAPH_BASE + path, headers=headers)

    def get_url_json(self, url: str) -> dict:
        return self._request_json("GET", url)

    def resolve_share(self, share_url: str) -> DriveItemRef:
        canonical = canonical_share_url(share_url, timeout=self._settings.timeout)
        candidates = [canonical]
        if share_url not in candidates:
            candidates.append(share_url)

        last_error: OneDriveError | None = None
        for candidate in candidates:
            encoded = encode_share_url(candidate)
            try:
                payload = self.get_json(
                    f"/shares/{encoded}/driveItem",
                    headers={"Prefer": "redeemSharingLinkIfNecessary"},
                )
            except OneDriveError as exc:
                last_error = exc
                continue
            item = payload.get("remoteItem") or payload
            parent = item.get("parentReference") or {}
            drive_id = str(parent.get("driveId") or "")
            if not drive_id:
                remote = payload.get("remoteItem") or {}
                drive_id = str(
                    (remote.get("parentReference") or {}).get("driveId") or ""
                )
            if not item.get("id"):
                last_error = OneDriveLookupError("无法解析 OneDrive 归档文件夹。")
                continue
            return DriveItemRef(
                drive_id=drive_id,
                item_id=str(item["id"]),
                name=str(item.get("name") or "OneDrive archive"),
                is_folder="folder" in item,
                path=str(item.get("name") or ""),
                web_url=str(item.get("webUrl") or ""),
                modified=str(item.get("lastModifiedDateTime") or ""),
                size=item.get("size"),
                share_id=encoded,
            )
        raise last_error or OneDriveLookupError("无法解析 OneDrive 归档链接。")

    def get_path_item(self, path: str) -> DriveItemRef:
        cleaned = (path or "").strip().strip("/")
        if not cleaned:
            raise OneDriveError("ONEDRIVE_ARCHIVE_PATH 不能为空。")
        if any(char in cleaned for char in ("\x00", "\\", "?", "#")) or ".." in cleaned.split("/"):
            raise OneDriveError("ONEDRIVE_ARCHIVE_PATH 含有不安全字符。")
        quoted = "/".join(urllib.parse.quote(part, safe="") for part in cleaned.split("/"))
        payload = self.get_json(f"/me/drive/root:/{quoted}")
        parent = payload.get("parentReference") or {}
        drive_id = str(parent.get("driveId") or "")
        if not drive_id:
            try:
                drive_id = str(self.get_json("/me/drive").get("id") or "")
            except OneDriveError:
                drive_id = ""
        if not payload.get("id"):
            raise OneDriveLookupError("无法解析 OneDrive 文件夹路径。")
        return DriveItemRef(
            drive_id=drive_id,
            item_id=str(payload["id"]),
            name=str(payload.get("name") or cleaned.rsplit("/", 1)[-1]),
            is_folder="folder" in payload,
            path=cleaned,
            web_url=str(payload.get("webUrl") or ""),
            modified=str(payload.get("lastModifiedDateTime") or ""),
            size=payload.get("size"),
        )

    def _children_url(self, ref: DriveItemRef) -> str:
        if ref.drive_id:
            return (
                f"{GRAPH_BASE}/drives/{urllib.parse.quote(ref.drive_id, safe='')}"
                f"/items/{urllib.parse.quote(ref.item_id, safe='')}/children"
                "?$top=200"
            )
        if ref.share_id:
            return (
                f"{GRAPH_BASE}/shares/{ref.share_id}/driveItem/children"
                "?$top=200"
            )
        raise OneDriveError("无法确定归档文件夹的 driveId。")

    def list_children(self, ref: DriveItemRef) -> list[DriveItemRef]:
        url = self._children_url(ref)
        results: list[DriveItemRef] = []
        while url:
            payload = self.get_url_json(url)
            for raw in payload.get("value") or []:
                if not isinstance(raw, dict):
                    continue
                item = raw.get("remoteItem") or raw
                name = str(item.get("name") or "")
                item_id = str(item.get("id") or "")
                if not name or not item_id:
                    continue
                parent = item.get("parentReference") or {}
                child_drive = str(parent.get("driveId") or ref.drive_id)
                results.append(
                    DriveItemRef(
                        drive_id=child_drive,
                        item_id=item_id,
                        name=name,
                        is_folder="folder" in item,
                        path=f"{ref.path}/{name}".strip("/"),
                        web_url=str(item.get("webUrl") or ""),
                        modified=str(item.get("lastModifiedDateTime") or ""),
                        size=item.get("size"),
                    )
                )
            next_link = payload.get("@odata.nextLink")
            url = str(next_link) if next_link else ""
            if url and not url.startswith(GRAPH_BASE + "/"):
                raise OneDriveError("Microsoft Graph 返回了意外的分页地址。")
        return results

    def list_tree(self, root: DriveItemRef) -> TreeListing:
        listing = TreeListing()
        if not root.is_folder:
            if is_pptx_filename(root.name):
                listing.files.append(
                    OneDriveFile(
                        drive_id=root.drive_id,
                        item_id=root.item_id,
                        name=root.name,
                        modified=root.modified,
                        size=root.size,
                        web_url=root.web_url,
                        path=root.path,
                        share_id=root.share_id,
                    )
                )
            return listing

        queue: deque[tuple[DriveItemRef, int]] = deque([(root, 0)])
        scanned = 0
        while queue:
            folder, depth = queue.popleft()
            if scanned >= MAX_ARCHIVE_CHILDREN:
                listing.truncated = True
                listing.warnings.append(
                    "归档文件较多，超过 5000 个条目后停止扫描。"
                )
                break
            try:
                children = self.list_children(folder)
            except OneDriveError as exc:
                listing.warnings.append(f"无法读取文件夹 {folder.path}: {exc}")
                continue
            for child in children:
                scanned += 1
                if scanned > MAX_ARCHIVE_CHILDREN:
                    listing.truncated = True
                    break
                if child.is_folder:
                    if depth + 1 <= MAX_ARCHIVE_DEPTH:
                        queue.append((child, depth + 1))
                    else:
                        listing.warnings.append(
                            f"已达到文件夹深度上限，跳过 {child.path}。"
                        )
                    continue
                if is_pptx_filename(child.name):
                    listing.files.append(
                        OneDriveFile(
                            drive_id=child.drive_id,
                            item_id=child.item_id,
                            name=child.name,
                            modified=child.modified,
                            size=child.size,
                            web_url=child.web_url,
                            path=child.path,
                            share_id=child.share_id,
                        )
                    )
        listing.files.sort(key=lambda item: item.name.lower())
        return listing

    def _content_url(self, file: OneDriveFile) -> str:
        if file.drive_id and file.item_id:
            return (
                f"{GRAPH_BASE}/drives/{urllib.parse.quote(file.drive_id, safe='')}"
                f"/items/{urllib.parse.quote(file.item_id, safe='')}/content"
            )
        if file.share_id:
            return f"{GRAPH_BASE}/shares/{file.share_id}/driveItem/content"
        raise OneDriveError("无法确定要下载的 OneDrive 文件。")

    def download_file(self, file: OneDriveFile, *, retry_auth: bool = True) -> bytes:
        token = self._tokens.get_access_token()
        url = self._content_url(file)
        request = urllib.request.Request(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "User-Agent": "WorshipPPT/1.0",
            },
        )
        opener = urllib.request.build_opener(_NoRedirectHandler())
        location = ""
        try:
            with opener.open(request, timeout=self._settings.timeout) as response:
                data = response.read(MAX_LAST_WEEK_PPTX_BYTES + 1)
                return _validate_download(data)
        except urllib.error.HTTPError as exc:
            if exc.code == 401 and retry_auth:
                self._tokens.invalidate()
                return self.download_file(file, retry_auth=False)
            if exc.code in REDIRECT_CODES:
                location = exc.headers.get("Location") or ""
                if not location:
                    raise OneDriveUploadError("OneDrive 下载重定向缺少地址。") from None
            else:
                raise _graph_error(exc.code, exc.read()) from None
        except (urllib.error.URLError, OSError) as exc:
            raise OneDriveError(
                f"无法下载 OneDrive 文件：{_sanitize_message(str(exc))}"
            ) from None

        if not is_allowed_download_url(location):
            raise OneDriveError("OneDrive 下载地址不是允许的主机，已拒绝。")
        download_request = urllib.request.Request(
            location, headers={"User-Agent": "WorshipPPT/1.0"}
        )
        try:
            with urllib.request.urlopen(
                download_request, timeout=max(self._settings.timeout, 60.0)
            ) as response:
                data = response.read(MAX_LAST_WEEK_PPTX_BYTES + 1)
        except (urllib.error.URLError, OSError) as exc:
            raise OneDriveError(
                f"无法下载 OneDrive 文件：{_sanitize_message(str(exc))}"
            ) from None
        return _validate_download(data)

    def _simple_upload(
        self, folder: DriveItemRef, filename: str, content: bytes
    ) -> dict:
        if not folder.drive_id:
            raise OneDriveUploadError("无法确定归档文件夹的 driveId，不能上传。")
        encoded_name = urllib.parse.quote(filename, safe="")
        url = (
            f"{GRAPH_BASE}/drives/{urllib.parse.quote(folder.drive_id, safe='')}"
            f"/items/{urllib.parse.quote(folder.item_id, safe='')}:/{encoded_name}:/content"
            "?%40microsoft.graph.conflictBehavior=replace"
        )
        return self._request_json(
            "PUT",
            url,
            data=content,
            headers={"Content-Type": PPTX_MIME},
        )

    def _session_upload(
        self, folder: DriveItemRef, filename: str, content: bytes
    ) -> dict:
        if not folder.drive_id:
            raise OneDriveUploadError("无法确定归档文件夹的 driveId，不能上传。")
        encoded_name = urllib.parse.quote(filename, safe="")
        url = (
            f"{GRAPH_BASE}/drives/{urllib.parse.quote(folder.drive_id, safe='')}"
            f"/items/{urllib.parse.quote(folder.item_id, safe='')}:/{encoded_name}:/"
            "createUploadSession"
        )
        body = json.dumps(
            {"item": {"@microsoft.graph.conflictBehavior": "replace"}}
        ).encode("utf-8")
        session = self._request_json(
            "POST", url, data=body, headers={"Content-Type": "application/json"}
        )
        upload_url = str(session.get("uploadUrl") or "")
        parsed = urllib.parse.urlparse(upload_url)
        host = (parsed.hostname or "").lower()
        if (
            parsed.scheme != "https"
            or not host
            or not any(
                host == suffix or host.endswith("." + suffix)
                for suffix in UPLOAD_HOST_SUFFIXES
            )
        ):
            raise OneDriveUploadError("OneDrive 返回了不安全的上传地址，已拒绝。")

        total = len(content)
        extra_headers = {"User-Agent": "WorshipPPT/1.0"}
        offset = 0
        last_payload: dict = {}
        while offset < total:
            end = min(offset + UPLOAD_CHUNK_SIZE, total)
            chunk = content[offset:end]
            request = urllib.request.Request(
                upload_url,
                data=chunk,
                method="PUT",
                headers={
                    **extra_headers,
                    "Content-Range": f"bytes {offset}-{end - 1}/{total}",
                    "Content-Type": PPTX_MIME,
                },
            )
            try:
                with urllib.request.urlopen(
                    request, timeout=max(self._settings.timeout, 120.0)
                ) as response:
                    raw = response.read()
            except urllib.error.HTTPError as exc:
                if exc.code == 401:
                    raise OneDriveUploadError(
                        "OneDrive 分块上传失败（401）。请重试；若持续失败，"
                        "可尝试重新生成或减小文件。"
                    ) from None
                raise OneDriveUploadError(
                    f"OneDrive 分块上传失败（{exc.code}）。"
                ) from None
            except (urllib.error.URLError, OSError) as exc:
                raise OneDriveUploadError(
                    f"OneDrive 分块上传失败：{_sanitize_message(str(exc))}"
                ) from None
            if raw:
                try:
                    parsed_payload = json.loads(raw.decode("utf-8"))
                    if isinstance(parsed_payload, dict):
                        last_payload = parsed_payload
                except ValueError:
                    pass
            offset = end
        return last_payload

    def upload_file(
        self, folder: DriveItemRef, filename: str, content: bytes
    ) -> dict:
        if not is_safe_upload_name(filename):
            raise OneDriveUploadError("上传文件名不安全。")
        if len(content) > MAX_LAST_WEEK_PPTX_BYTES:
            raise OneDriveUploadError("生成的 PPT 超过 50 MB 上限。")
        try:
            result = self._simple_upload(folder, filename, content)
            if result.get("id"):
                return result
        except GraphHTTPError as exc:
            if exc.status not in (400, 413) or len(content) <= SIMPLE_UPLOAD_THRESHOLD:
                raise
        result = self._session_upload(folder, filename, content)
        if not result.get("id"):
            # Some personal-OneDrive upload sessions return an empty final
            # response.  Resolve the uploaded child by name instead.
            for child in self.list_children(folder):
                if (
                    not child.is_folder
                    and child.name.lower() == filename.lower()
                ):
                    return {
                        "id": child.item_id,
                        "name": child.name,
                        "webUrl": child.web_url,
                        "parentReference": {"driveId": child.drive_id},
                    }
            raise OneDriveUploadError("OneDrive 上传完成，但没有返回文件信息。")
        return result

    def create_share_link(self, drive_id: str, item_id: str) -> str:
        if not drive_id or not item_id:
            raise OneDriveUploadError("无法为上传文件创建分享链接。")
        url = (
            f"{GRAPH_BASE}/drives/{urllib.parse.quote(drive_id, safe='')}"
            f"/items/{urllib.parse.quote(item_id, safe='')}/createLink"
        )
        body = json.dumps({"type": "view", "scope": "anonymous"}).encode("utf-8")
        payload = self._request_json(
            "POST", url, data=body, headers={"Content-Type": "application/json"}
        )
        link = payload.get("link") or {}
        web_url = str(link.get("webUrl") or "")
        if not web_url:
            raise OneDriveUploadError("OneDrive 没有返回分享链接。")
        return web_url


def is_allowed_download_url(url: str) -> bool:
    """Allow OneDrive/SharePoint and Microsoft consumer download hosts."""
    if is_allowed_onedrive_url(url):
        return True
    parsed = urllib.parse.urlparse(url)
    host = (parsed.hostname or "").lower()
    return (
        parsed.scheme == "https"
        and bool(host)
        and (
            host == "microsoftpersonalcontent.com"
            or host.endswith(".microsoftpersonalcontent.com")
        )
    )


def _validate_download(data: bytes) -> bytes:
    if len(data) > MAX_LAST_WEEK_PPTX_BYTES:
        raise OneDriveError("下载的 PPT 超过 50 MB 上限。")
    try:
        return validate_pptx_bytes(data)
    except ValueError as exc:
        raise OneDriveError(str(exc)) from None


class OneDriveArchive:
    def __init__(self, settings: OneDriveSettings):
        self.settings = settings
        self.tokens = TokenProvider(settings)
        self.graph = GraphClient(settings, self.tokens)
        # list_files() holds this lock and calls root(), which re-acquires it.
        self._lock = threading.RLock()
        self._root: DriveItemRef | None = None
        self._listing: TreeListing | None = None
        self._listing_at = 0.0

    def root(self, *, force: bool = False) -> DriveItemRef:
        with self._lock:
            if self._root is not None and not force:
                return self._root
            if self.settings.mode == "path":
                self._root = self.graph.get_path_item(self.settings.archive_path)
            else:
                self._root = self.graph.resolve_share(self.settings.archive_share_url)
            return self._root

    def list_files(self, *, force: bool = False) -> TreeListing:
        with self._lock:
            if (
                self._listing is not None
                and not force
                and time.time() - self._listing_at < 60
            ):
                return self._listing
            root = self.root()
            self._listing = self.graph.list_tree(root)
            self._listing_at = time.time()
            return self._listing

    def lookup(self, service_date: date, filename: str = "") -> ArchiveLookup:
        listing = self.list_files()
        files = listing.files
        target = previous_sunday(service_date)

        if filename.strip():
            wanted = filename.strip().lower()
            matches = [item for item in files if item.name.lower() == wanted]
            if len(matches) == 1:
                return ArchiveLookup(
                    file=matches[0], target=target, warnings=list(listing.warnings)
                )
            return ArchiveLookup(
                target=target,
                candidates=[item.as_candidate() for item in files[:MAX_CANDIDATES]],
                warnings=list(listing.warnings),
                error=f"没有找到文件名为 {filename} 的 PPT。",
            )

        exact = [
            item
            for item in files
            if target in parse_dates_from_filename(item.name)
        ]
        if len(exact) == 1:
            return ArchiveLookup(
                file=exact[0], target=target, warnings=list(listing.warnings)
            )

        candidates = exact or files
        candidates = sorted(
            candidates,
            key=lambda item: (
                parse_dates_from_filename(item.name)[0]
                if parse_dates_from_filename(item.name)
                else date.min
            ),
            reverse=True,
        )[:MAX_CANDIDATES]
        if exact:
            error = (
                f"找到多个包含 {target.isoformat()} 的 PPT，"
                "请用 last_week_filename 选择正确文件。"
            )
        else:
            error = f"归档中没有找到 {target.isoformat()} 的 PPT。"
        return ArchiveLookup(
            target=target,
            candidates=[item.as_candidate() for item in candidates],
            warnings=list(listing.warnings),
            error=error,
        )

    def upload_generated_deck(self, service_date: date, content: bytes) -> dict:
        if not self.settings.upload_enabled:
            return {
                "uploaded": False,
                "filename": "",
                "web_url": "",
                "share_url": "",
                "error": "OneDrive 写回未启用。",
                "warnings": [],
            }
        if not self.settings.can_write:
            return {
                "uploaded": False,
                "filename": "",
                "web_url": "",
                "share_url": "",
                "error": "当前 OneDrive 权限不含 Files.ReadWrite，不能写回。",
                "warnings": [],
            }

        filename = f"{service_date.isoformat()}.pptx"
        warnings: list[str] = []
        try:
            root = self.root()
            if not root.is_folder:
                raise OneDriveUploadError("OneDrive 归档链接不是文件夹。")
            try:
                if any(
                    item.name.lower() == filename.lower()
                    for item in self.list_files().files
                ):
                    warnings.append(
                        f"归档中已有同日期文件 {filename}，本次已替换。"
                    )
            except OneDriveError:
                pass
            uploaded = self.graph.upload_file(root, filename, content)
            drive_id = str(
                (uploaded.get("parentReference") or {}).get("driveId")
                or root.drive_id
            )
            item_id = str(uploaded.get("id") or "")
            web_url = str(uploaded.get("webUrl") or "")
            share_url = ""
            try:
                share_url = self.graph.create_share_link(drive_id, item_id)
            except OneDriveError as exc:
                warnings.append(
                    f"文件已写回，但创建分享链接失败：{exc}"
                )
            return {
                "uploaded": True,
                "filename": str(uploaded.get("name") or filename),
                "web_url": web_url,
                "share_url": share_url,
                "error": "",
                "warnings": warnings,
            }
        except OneDriveError as exc:
            return {
                "uploaded": False,
                "filename": filename,
                "web_url": "",
                "share_url": "",
                "error": str(exc),
                "warnings": warnings,
            }
        except Exception:
            return {
                "uploaded": False,
                "filename": filename,
                "web_url": "",
                "share_url": "",
                "error": "OneDrive 写回时发生未知错误。",
                "warnings": warnings,
            }


_ARCHIVE: OneDriveArchive | None = None
_ARCHIVE_KEY: tuple | None = None
_ARCHIVE_LOCK = threading.Lock()


def _settings_key(settings: OneDriveSettings) -> tuple:
    return (
        settings.client_id,
        settings.refresh_token,
        settings.tenant,
        settings.scope,
        settings.archive_share_url,
        settings.archive_path,
        settings.auto_enabled,
        settings.upload_enabled,
    )


def get_onedrive_archive() -> OneDriveArchive | None:
    global _ARCHIVE, _ARCHIVE_KEY
    settings = OneDriveSettings.from_env()
    if settings is None:
        with _ARCHIVE_LOCK:
            _ARCHIVE = None
            _ARCHIVE_KEY = None
        return None
    key = _settings_key(settings)
    with _ARCHIVE_LOCK:
        if _ARCHIVE is None or _ARCHIVE_KEY != key:
            _ARCHIVE = OneDriveArchive(settings)
            _ARCHIVE_KEY = key
        return _ARCHIVE


def reset_onedrive_cache_for_tests() -> None:
    global _ARCHIVE, _ARCHIVE_KEY
    with _ARCHIVE_LOCK:
        _ARCHIVE = None
        _ARCHIVE_KEY = None


def auto_last_week_pptx(
    service_date: date, filename: str = ""
) -> AutoLastWeekResult:
    settings = OneDriveSettings.from_env()
    if settings is None:
        return AutoLastWeekResult()
    if not settings.auto_enabled:
        return AutoLastWeekResult(
            attempted=False, error="OneDrive 自动读取已关闭。"
        )
    archive = get_onedrive_archive()
    if archive is None:
        return AutoLastWeekResult()
    try:
        lookup = archive.lookup(service_date, filename=filename)
        if lookup.file is None:
            return AutoLastWeekResult(
                attempted=True,
                error=lookup.error or "没有唯一确定上周文件。",
                candidates=lookup.candidates,
                warnings=lookup.warnings,
            )
        content = archive.graph.download_file(lookup.file)
        return AutoLastWeekResult(
            content=content,
            filename=lookup.file.name,
            attempted=True,
            warnings=lookup.warnings,
        )
    except OneDriveError as exc:
        return AutoLastWeekResult(attempted=True, error=str(exc))
    except Exception:
        return AutoLastWeekResult(
            attempted=True, error="自动读取 OneDrive 时发生未知错误。"
        )


def get_status(today: date | None = None) -> dict:
    settings = OneDriveSettings.from_env()
    if settings is None:
        return {
            "configured": False,
            "mode": "none",
            "auto_enabled": False,
            "upload_enabled": False,
            "token_status": "unknown",
            "token_age_days": None,
            "archive": "",
            "error": "未配置 ONEDRIVE_CLIENT_ID 或 ONEDRIVE_REFRESH_TOKEN。",
        }
    age = settings.token_age_days(today)
    return {
        "configured": True,
        "mode": settings.mode,
        "auto_enabled": settings.auto_enabled,
        "upload_enabled": settings.upload_enabled and settings.can_write,
        "token_status": settings.token_status(today),
        "token_age_days": age,
        "archive": settings.archive_path if settings.mode == "path" else "OneDrive 归档",
        "error": "" if settings.can_write else "当前权限为只读，写回/分享链接不可用。",
    }
