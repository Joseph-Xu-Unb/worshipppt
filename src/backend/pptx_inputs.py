"""Validation and public-download helpers for last-week PPTX inputs."""

from __future__ import annotations

import base64
import binascii
import urllib.error
import urllib.request
import zipfile
from io import BytesIO
from urllib.parse import urljoin, urlparse

from pptx import Presentation


MAX_LAST_WEEK_PPTX_BYTES = 50 * 1024 * 1024
MAX_PPTX_UNCOMPRESSED_BYTES = 250 * 1024 * 1024


def is_allowed_onedrive_host(host: str) -> bool:
    """Return True for OneDrive/SharePoint download hosts."""
    host = host.lower().rstrip(".")
    return (
        host in {"onedrive.live.com", "1drv.ms", "onedrive.com", "sharepoint.com"}
        or host.endswith(".onedrive.live.com")
        or host.endswith(".1drv.ms")
        or host.endswith(".onedrive.com")
        or host.endswith(".sharepoint.com")
        or host in {"files.1drv.com"}
        or host.endswith(".files.1drv.com")
    )


def is_allowed_onedrive_url(url: str) -> bool:
    parsed = urlparse(url)
    return (
        parsed.scheme == "https"
        and bool(parsed.hostname)
        and is_allowed_onedrive_host(parsed.hostname)
    )


class OneDriveRedirectHandler(urllib.request.HTTPRedirectHandler):
    """Follow redirects only while they stay on OneDrive/SharePoint hosts."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        newurl = urljoin(req.full_url, newurl)
        if not is_allowed_onedrive_url(newurl):
            raise urllib.error.HTTPError(
                req.full_url,
                code,
                "Blocked redirect to a non-OneDrive host.",
                headers,
                fp,
            )
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def validate_pptx_bytes(
    data: bytes, *, max_bytes: int = MAX_LAST_WEEK_PPTX_BYTES
) -> bytes:
    """Validate that ``data`` is an unencrypted PowerPoint package."""
    if len(data) > max_bytes:
        raise ValueError("Last week's PPT exceeds the 50 MB limit.")
    if not data.startswith(b"PK"):
        raise ValueError("The supplied last-week file is not a .pptx file.")
    try:
        with zipfile.ZipFile(BytesIO(data)) as archive:
            uncompressed = sum(info.file_size for info in archive.infolist())
            if uncompressed > MAX_PPTX_UNCOMPRESSED_BYTES:
                raise ValueError(
                    "The supplied last-week file expands beyond the 250 MB safety limit."
                )
        Presentation(BytesIO(data))
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(
            "The supplied last-week file is not a readable .pptx file."
        ) from exc
    return data


def decode_base64_pptx(
    value: str, *, max_bytes: int = MAX_LAST_WEEK_PPTX_BYTES
) -> bytes:
    compact = "".join(value.split())
    if len(compact) > max_bytes * 2:
        raise ValueError("last_week_pptx_base64 exceeds the 50 MB limit.")
    try:
        data = base64.b64decode(compact, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ValueError("last_week_pptx_base64 is not valid base64.") from exc
    return validate_pptx_bytes(data, max_bytes=max_bytes)


def download_public_pptx(
    url: str, *, timeout: float = 30.0
) -> bytes:
    """Download a public OneDrive/SharePoint PPTX with host and size limits."""
    if not is_allowed_onedrive_url(url):
        raise ValueError(
            "last_week_pptx_url must be an https OneDrive or SharePoint link."
        )

    opener = urllib.request.build_opener(OneDriveRedirectHandler())
    request = urllib.request.Request(url, headers={"User-Agent": "WorshipPPT/1.0"})
    try:
        with opener.open(request, timeout=timeout) as response:
            data = response.read(MAX_LAST_WEEK_PPTX_BYTES + 1)
    except (urllib.error.URLError, OSError) as exc:
        raise ValueError(f"Could not download last week's PPT: {exc}") from exc

    if len(data) > MAX_LAST_WEEK_PPTX_BYTES:
        raise ValueError("Last week's PPT exceeds the 50 MB limit.")
    return validate_pptx_bytes(data)
