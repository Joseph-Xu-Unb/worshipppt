"""FastAPI backend for generating worship PowerPoint files."""

from __future__ import annotations

import base64
import binascii
import json
import os
import tempfile
import urllib.error
import urllib.request
from contextlib import asynccontextmanager
from datetime import date
from io import BytesIO
from pathlib import Path
from urllib.parse import urljoin, urlparse

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse, StreamingResponse
from pptx import Presentation

from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings

from src.backend.content import is_marker, section_pages, validate_section
from src.backend.matters import copy_matters_slides
from src.backend.worship import generate_worship_ppt


BACKEND_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BACKEND_DIR.parent.parent
DIST_DIR = PROJECT_ROOT / "dist"
DATA_DIR = PROJECT_ROOT / "data"
DEFAULT_TEMPLATE_PATH = DATA_DIR / "template.pptx"
SAMPLE_JSON_PATH = DATA_DIR / "sample_worship_data.json"
SAMPLE_PPT_PATH = DATA_DIR / "2026-05-17.pptx"
REQUIRED_SECTIONS = ("call_to_worship", "hymns", "theme_scripture", "response_hymn")

WORSHIP_ARCHIVE_URL = (
    "https://1drv.ms/f/c/1b40ff3184add23e/IgBvMYWWe1v8RpbhkkjUr142Aar3qCdtzJEPobKBbub6zMU?e=lFGb49"
)

MATTERS_COPY_REMINDER = (
    "重要后续步骤（事项汇报页复制）：模板生成的事项汇报页（matters）默认为空白。请前往往期崇拜 PPT 归档库（"
    f"{WORSHIP_ARCHIVE_URL}"
    "），找到上周（即距离本次崇拜最近的前一个主日）的 worship ppt 文件，将其中的“事项报告 / 事项汇报”（matters）页内容完整复制到当前生成的 PPT 对应页面中，确保事项汇报不留空。"
)

LAST_WEEK_DOWNLOAD_LINK_GUIDE = (
    "上周 PPT 下载链接获取（用于服务端复制事项汇报页）：\n"
    "目标文件：上周主日的 worship ppt（即距离本次崇拜最近的前一个主日，文件名通常为该主日日期，"
    "例如 2026-11-22.pptx，来源应为 onedrive.live.com）。不要自行假定为其他文件；先核对文件名和来源。\n"
    "如果该文件已在你的浏览器中下载过，请从浏览器下载记录中取得其直接下载链接，不要重新浏览 OneDrive 页面：\n"
    "1. 打开浏览器下载记录：点击浏览器右上角菜单 → 下载内容（Downloads），或按 Ctrl+J，"
    "或在地址栏输入 chrome://downloads 后回车。\n"
    "2. 在下载列表中找到上述目标文件，核对文件名和来源无误。\n"
    "3. 点击该文件条目旁的链接（🔗）图标，复制下载链接。链接格式通常为：\n"
    "   https://onedrive.live.com/personal/1b40ff3184add23e/_layouts/15/download.aspx?UniqueId=<文件GUID>\n"
    "4. 将该链接作为 last_week_pptx_url 传给 generate_worship_powerpoint，由服务端复制上周的事项汇报页。\n"
    "如果下载记录中没有该文件或无法取得链接，请改用往期崇拜 PPT 归档库链接打开并下载："
    f"{WORSHIP_ARCHIVE_URL}"
)

POST_GENERATION_REMINDER = f"{MATTERS_COPY_REMINDER}\n\n{LAST_WEEK_DOWNLOAD_LINK_GUIDE}"

MATTERS_COPIED_REMINDER = (
    "已从上周 PPT 复制 {count} 页事项汇报/公告内容到本次生成的 PPT。"
    "请打开生成的文件核对文字、图片、表格、图表和分组内容是否完整。"
)

MATTERS_COPY_GUIDE_PREFIX = (
    "Preferred: pass last_week_pptx_url or last_week_pptx_base64 to "
    "generate_worship_powerpoint. The server copies the matters slides, including "
    "pictures, tables, grouped shapes, and charts.\n"
    "If that is not possible, follow these manual instructions:\n\n"
)

MAX_LAST_WEEK_PPTX_BYTES = 50 * 1024 * 1024

MCP_INSTRUCTIONS = (
    "WorshipPPT Presentation Server.\n"
    "Generates Sunday worship service presentations (PPTX) from structured JSON worship data.\n"
    "CRITICAL AGENT INSTRUCTION:\n"
    "The generated presentation's announcements/matters slide ('matters') is intentionally left blank by the template. "
    "To have the server copy last week's matters slide for you, pass last_week_pptx_url (a public https "
    "OneDrive/SharePoint download link) or last_week_pptx_base64 (the .pptx bytes) to generate_worship_powerpoint. "
    "The server copies text, pictures, tables, grouped shapes, charts, and embedded chart data.\n"
    "If you cannot supply the file, the response returns instructions for copying it manually: "
    f"go to the worship PPT archive ({WORSHIP_ARCHIVE_URL}), locate last week's worship PPT, and copy its "
    "matters/announcements content into the new presentation. If last week's PPT was already downloaded in your "
    "browser, get its direct download link from the browser's download history (Ctrl+J or chrome://downloads): "
    "click the link icon next to the entry and use the resulting download.aspx?UniqueId=... URL."
)

mcp_server = FastMCP(
    "WorshipPPT",
    instructions=MCP_INSTRUCTIONS,
    stateless_http=True,
    json_response=True,
    streamable_http_path="/",
    transport_security=TransportSecuritySettings(
        enable_dns_rebinding_protection=True,
        allowed_hosts=os.getenv(
            "MCP_ALLOWED_HOSTS",
            "localhost,localhost:*,127.0.0.1,127.0.0.1:*,worshipppt.onrender.com",
        ).split(","),
        allowed_origins=os.getenv(
            "MCP_ALLOWED_ORIGINS",
            "http://localhost:*,http://127.0.0.1:*,https://worshipppt.onrender.com",
        ).split(","),
    ),
)
mcp_app = mcp_server.streamable_http_app()


@asynccontextmanager
async def lifespan(app: FastAPI):
    async with mcp_server.session_manager.run():
        yield


app = FastAPI(lifespan=lifespan)
app.mount("/mcp", mcp_app)

allowed_origins = [
    origin.strip()
    for origin in os.getenv("ALLOWED_ORIGINS", "*").split(",")
    if origin.strip()
]
app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins or ["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=[
        "Content-Disposition",
        "X-Worship-Warnings",
        "X-Worship-Matters-Copied",
    ],
)


def _validate_data(data):
    """Return a validation error string, or None when input data is usable."""
    if not isinstance(data, dict):
        return "JSON root must be an object."

    missing = [key for key in REQUIRED_SECTIONS if key not in data]
    if missing:
        return f"Missing required JSON section(s): {', '.join(missing)}."

    for section in ("call_to_worship", "theme_scripture", "response_hymn"):
        error = validate_section(data[section], section)
        if error:
            return error

    if not isinstance(data["hymns"], list) or not data["hymns"]:
        return "hymns must be a non-empty list."

    for index, hymn in enumerate(data["hymns"]):
        error = validate_section(hymn, f"hymns[{index}]")
        if error:
            return error

    return None


def _is_allowed_onedrive_host(host: str) -> bool:
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


def _is_allowed_onedrive_url(url: str) -> bool:
    parsed = urlparse(url)
    return (
        parsed.scheme == "https"
        and bool(parsed.hostname)
        and _is_allowed_onedrive_host(parsed.hostname)
    )


class _OneDriveRedirectHandler(urllib.request.HTTPRedirectHandler):
    """Follow redirects only while they stay on OneDrive/SharePoint hosts."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        newurl = urljoin(req.full_url, newurl)
        if not _is_allowed_onedrive_url(newurl):
            raise urllib.error.HTTPError(
                req.full_url,
                code,
                "Blocked redirect to a non-OneDrive host.",
                headers,
                fp,
            )
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _validate_pptx_bytes(data: bytes) -> bytes:
    """Validate that ``data`` is an unencrypted PowerPoint package."""
    if not data.startswith(b"PK"):
        raise ValueError("The supplied last-week file is not a .pptx file.")
    try:
        Presentation(BytesIO(data))
    except Exception as exc:
        raise ValueError(
            "The supplied last-week file is not a readable .pptx file."
        ) from exc
    return data


def _download_pptx(url: str) -> bytes:
    """Download a public OneDrive/SharePoint PPTX with host and size limits."""
    if not _is_allowed_onedrive_url(url):
        raise ValueError(
            "last_week_pptx_url must be an https OneDrive or SharePoint link."
        )

    opener = urllib.request.build_opener(_OneDriveRedirectHandler())
    request = urllib.request.Request(url, headers={"User-Agent": "WorshipPPT/1.0"})
    try:
        with opener.open(request, timeout=30) as response:
            data = response.read(MAX_LAST_WEEK_PPTX_BYTES + 1)
    except (urllib.error.URLError, OSError) as exc:
        raise ValueError(f"Could not download last week's PPT: {exc}") from exc

    if len(data) > MAX_LAST_WEEK_PPTX_BYTES:
        raise ValueError("Last week's PPT exceeds the 50 MB limit.")
    return _validate_pptx_bytes(data)


def _decode_pptx_base64(value: str) -> bytes:
    compact = "".join(value.split())
    if len(compact) > MAX_LAST_WEEK_PPTX_BYTES * 2:
        raise ValueError("last_week_pptx_base64 exceeds the 50 MB limit.")
    try:
        data = base64.b64decode(compact, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ValueError("last_week_pptx_base64 is not valid base64.") from exc
    if len(data) > MAX_LAST_WEEK_PPTX_BYTES:
        raise ValueError("Last week's PPT exceeds the 50 MB limit.")
    return _validate_pptx_bytes(data)


def _last_week_pptx_from_inputs(url: str, encoded: str) -> bytes | None:
    """Resolve the optional last-week deck from a URL or base64 payload."""
    if encoded:
        return _decode_pptx_base64(encoded)
    if url:
        return _download_pptx(url)
    return None


def _copy_matters_into(
    output_path: str | Path, last_week_pptx: bytes | None
) -> tuple[int, list[str]]:
    if not last_week_pptx:
        return 0, []
    return copy_matters_slides(output_path, last_week_pptx)


@mcp_server.tool()
def get_sample_worship_data() -> dict:
    """Get example worship JSON. Sections accept legacy lines or explicit pages.

    For songs prefer pages: [{"lines": ["lyric 1", "lyric 2"], "footer": "间奏"}].
    Each entry is exactly one slide; footer is optional (间奏/前奏/尾声/结束).
    Do not also supply section-level lines when using pages. Blank lines do not
    define slide boundaries in the legacy lines format.

    CRITICAL POST-PROCESSING NOTE:
    The template leaves the matters/announcements slide blank. To let the server copy
    last week's matters slide for you, pass last_week_pptx_url or last_week_pptx_base64
    to generate_worship_powerpoint. Otherwise, visit the OneDrive archive:
    https://1drv.ms/f/c/1b40ff3184add23e/IgBvMYWWe1v8RpbhkkjUr142Aar3qCdtzJEPobKBbub6zMU?e=lFGb49
    find last week's worship PPT, and copy its matters/announcements slide into this presentation.
    If last week's PPT was already downloaded in your browser, get its direct download link from the
    browser's download history (Ctrl+J or chrome://downloads) instead of re-opening OneDrive.
    """
    data = json.loads(SAMPLE_JSON_PATH.read_text(encoding="utf-8"))
    for song in [*data["hymns"], data["response_hymn"]]:
        pages = []
        for lines in section_pages(song):
            if is_marker(lines[-1]):
                pages.append({"lines": lines[:-1], "footer": lines[-1]})
            else:
                pages.append({"lines": lines})
        song.pop("lines", None)
        song["pages"] = pages
    return data


@mcp_server.tool()
def validate_worship_data(data: dict) -> dict:
    """Validate worship JSON before generating a presentation."""
    error = _validate_data(data)
    return {"ok": error is None, "error": error}


@mcp_server.tool()
def generate_worship_powerpoint(
    data: dict,
    service_date: str = "",
    last_week_pptx_url: str = "",
    last_week_pptx_base64: str = "",
) -> dict:
    """Generate a PPTX with the bundled template. Decode pptx_base64 to save the file.

    service_date must be YYYY-MM-DD; omitted dates use the server's current date.
    Review returned warnings for content that needs attention.

    To have the server copy last week's announcement/matters slide into this deck, pass
    last_week_pptx_url (a public https OneDrive/SharePoint download link) or
    last_week_pptx_base64 (the .pptx bytes). Pictures, tables, grouped shapes, charts,
    and embedded chart data are copied; the generated deck's matters layout is kept.
    If neither is supplied, matters stays blank and the response includes manual-copy
    instructions.

    CRITICAL POST-PROCESSING REQUIREMENT:
    The template leaves the matters/announcements slide blank. If you did not pass a
    last-week file, you must visit the OneDrive worship archive:
    https://1drv.ms/f/c/1b40ff3184add23e/IgBvMYWWe1v8RpbhkkjUr142Aar3qCdtzJEPobKBbub6zMU?e=lFGb49
    find last week's worship PPT, and copy its matters/announcements slide into this PPT.
    The response's post_generation_prompt explains how to obtain last week's PPT directly
    from the browser download history when it was already downloaded.
    """
    error = _validate_data(data)
    if error:
        raise ValueError(error)
    selected_date = date.fromisoformat(service_date) if service_date else date.today()
    last_week_pptx = _last_week_pptx_from_inputs(
        last_week_pptx_url.strip(), last_week_pptx_base64.strip()
    )
    with tempfile.TemporaryDirectory(prefix="worship-mcp-") as directory:
        template = Path(directory) / "template.pptx"
        output = Path(directory) / f"Worship_{selected_date.isoformat()}.pptx"
        template.write_bytes(DEFAULT_TEMPLATE_PATH.read_bytes())
        warnings = generate_worship_ppt(
            str(template), data, str(output), selected_date=selected_date.isoformat()
        )
        try:
            matters_copied, matters_warnings = _copy_matters_into(
                output, last_week_pptx
            )
        except Exception as exc:
            raise ValueError(
                f"Could not copy last week's matters slide: {exc}"
            ) from exc

        all_warnings = (warnings or []) + matters_warnings
        if matters_copied:
            reminder = MATTERS_COPIED_REMINDER.format(count=matters_copied)
            post_generation_prompt = ""
        else:
            all_warnings.append(POST_GENERATION_REMINDER)
            reminder = POST_GENERATION_REMINDER
            post_generation_prompt = LAST_WEEK_DOWNLOAD_LINK_GUIDE

        return {
            "filename": output.name,
            "mime_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "pptx_base64": base64.b64encode(output.read_bytes()).decode("ascii"),
            "warnings": all_warnings,
            "reminder": reminder,
            "post_generation_prompt": post_generation_prompt,
            "matters_copied": matters_copied > 0,
            "matters_slides_copied": matters_copied,
        }


@mcp_server.prompt("worship_matters_copy_guide")
def worship_matters_copy_guide() -> str:
    """Prompt guide for copying last week's announcements into this week's PPT."""
    return MATTERS_COPY_GUIDE_PREFIX + POST_GENERATION_REMINDER


@app.get("/api/health")
def health():
    """Report whether the backend is running."""
    return {"ok": True}


@app.post("/api/validate")
async def validate(data: UploadFile = File(...)):
    """Validate uploaded worship JSON before preview/generation."""
    if not data.filename or not data.filename.lower().endswith(".json"):
        raise HTTPException(status_code=400, detail="Worship data must be a .json file.")

    try:
        payload = json.loads((await data.read()).decode("utf-8"))
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid JSON: {exc.msg}.") from exc

    validation_error = _validate_data(payload)
    if validation_error:
        raise HTTPException(status_code=400, detail=validation_error)

    return {"ok": True}


@app.get("/downloads/template")
def download_template():
    """Download the bundled PowerPoint template."""
    if not DEFAULT_TEMPLATE_PATH.exists():
        raise HTTPException(status_code=404, detail="template.pptx not found.")
    return FileResponse(
        DEFAULT_TEMPLATE_PATH,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="template.pptx",
    )


@app.get("/downloads/sample-json")
def download_sample_json():
    """Download the bundled sample worship JSON data."""
    if not SAMPLE_JSON_PATH.exists():
        raise HTTPException(status_code=404, detail="Sample JSON file not found.")
    return FileResponse(
        SAMPLE_JSON_PATH,
        media_type="application/json",
        filename="sample_worship_data.json",
    )


@app.get("/downloads/sample-ppt")
def download_sample_ppt():
    """Download the bundled sample PowerPoint deck."""
    if not SAMPLE_PPT_PATH.exists():
        raise HTTPException(status_code=404, detail="Sample PPT file not found.")
    return FileResponse(
        SAMPLE_PPT_PATH,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="2026-05-17.pptx",
    )


@app.post("/api/generate")
async def generate(
    data: UploadFile = File(...),
    selected_date: str = Form(default="", alias="date"),
    slides: str = Form(default=""),
    last_week_pptx: UploadFile | None = File(default=None),
    last_week_pptx_url: str = Form(default=""),
):
    """Generate a worship PPT from the bundled template and uploaded JSON file."""
    selected_date = selected_date.strip() or date.today().isoformat()

    if not data.filename or not data.filename.lower().endswith(".json"):
        raise HTTPException(status_code=400, detail="Worship data must be a .json file.")
    if not DEFAULT_TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail="Server template.pptx is missing.",
        )

    try:
        payload = json.loads((await data.read()).decode("utf-8"))
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid JSON: {exc.msg}.") from exc

    validation_error = _validate_data(payload)
    if validation_error:
        raise HTTPException(status_code=400, detail=validation_error)

    last_week_bytes = None
    if last_week_pptx is not None and last_week_pptx.filename:
        if not last_week_pptx.filename.lower().endswith(".pptx"):
            raise HTTPException(
                status_code=400,
                detail="Last week's file must be a .pptx file.",
            )
        raw = await last_week_pptx.read()
        if len(raw) > MAX_LAST_WEEK_PPTX_BYTES:
            raise HTTPException(
                status_code=400,
                detail="Last week's PPT exceeds the 50 MB limit.",
            )
        try:
            last_week_bytes = _validate_pptx_bytes(raw)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
    elif last_week_pptx_url.strip():
        try:
            last_week_bytes = _download_pptx(last_week_pptx_url.strip())
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    if slides:
        try:
            preview_slides = json.loads(slides)
        except json.JSONDecodeError as exc:
            raise HTTPException(status_code=400, detail="Invalid edited slide data.") from exc

        allowed_kinds = {
            "call_to_worship",
            "hymn",
            "theme_scripture",
            "response_hymn",
        }
        if not isinstance(preview_slides, list):
            raise HTTPException(status_code=400, detail="Edited slides must be a list.")
        for slide in preview_slides:
            if (
                not isinstance(slide, dict)
                or slide.get("kind") not in allowed_kinds
                or not isinstance(slide.get("title"), str)
                or not slide["title"].strip()
                or not isinstance(slide.get("lines"), list)
                or not all(isinstance(line, str) for line in slide["lines"])
            ):
                raise HTTPException(status_code=400, detail="An edited slide is invalid.")
        payload["preview_slides"] = preview_slides

    with tempfile.TemporaryDirectory(prefix="worship-ppt-") as tmp_dir_name:
        tmp_dir = Path(tmp_dir_name)
        template_path = tmp_dir / DEFAULT_TEMPLATE_PATH.name
        output_path = tmp_dir / f"Worship_{selected_date}.pptx"

        template_path.write_bytes(DEFAULT_TEMPLATE_PATH.read_bytes())
        warnings = generate_worship_ppt(
            str(template_path),
            payload,
            str(output_path),
            selected_date=selected_date,
        )
        try:
            matters_copied, matters_warnings = _copy_matters_into(
                output_path, last_week_bytes
            )
        except Exception as exc:
            raise HTTPException(
                status_code=400,
                detail=f"Could not copy last week's matters slide: {exc}",
            ) from exc

        warnings = (warnings or []) + matters_warnings
        ppt_bytes = BytesIO(output_path.read_bytes())
        response = StreamingResponse(
            ppt_bytes,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        response.headers["Content-Disposition"] = (
            f'attachment; filename="{output_path.name}"'
        )
        response.headers["X-Worship-Matters-Copied"] = str(matters_copied)
        if warnings:
            response.headers["X-Worship-Warnings"] = json.dumps(warnings)
        return response


@app.exception_handler(HTTPException)
async def http_exception_handler(_, exc: HTTPException):
    return JSONResponse(status_code=exc.status_code, content={"error": exc.detail})


@app.get("/")
def index():
    """Serve the built React app when available."""
    index_path = DIST_DIR / "index.html"
    if not index_path.exists():
        return PlainTextResponse(
            "React app is not built yet. Run 'npm install' and 'npm run build', "
            "or use 'npm run dev' during development.",
            status_code=404,
        )
    return FileResponse(index_path)


@app.get("/{path:path}")
def static_assets(path: str):
    """Serve React build assets with SPA fallback."""
    asset_path = DIST_DIR / path
    if asset_path.exists() and asset_path.is_file():
        return FileResponse(asset_path)
    return index()
