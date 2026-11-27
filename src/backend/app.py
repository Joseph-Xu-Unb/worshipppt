"""FastAPI backend for generating worship PowerPoint files."""

from __future__ import annotations

import base64
import json
import os
import tempfile
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from datetime import date
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse, StreamingResponse
from pptx import Presentation
from starlette.concurrency import run_in_threadpool

from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings

from src.backend import onedrive
from src.backend.content import is_marker, section_pages, validate_section
from src.backend.matters import copy_matters_slides, list_matters_candidates
from src.backend.pptx_inputs import (
    MAX_LAST_WEEK_PPTX_BYTES,
    decode_base64_pptx as _decode_pptx_base64,
    download_public_pptx as _download_pptx,
    is_allowed_onedrive_url as _is_allowed_onedrive_url,
    validate_pptx_bytes as _validate_pptx_bytes,
)
from src.backend.worship import generate_worship_ppt


BACKEND_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BACKEND_DIR.parent.parent
DIST_DIR = PROJECT_ROOT / "dist"
DATA_DIR = PROJECT_ROOT / "data"
DEFAULT_TEMPLATE_PATH = DATA_DIR / "template.pptx"
SAMPLE_JSON_PATH = DATA_DIR / "sample_worship_data.json"
SAMPLE_PPT_PATH = DATA_DIR / "2026-05-17.pptx"
REQUIRED_SECTIONS = ("call_to_worship", "hymns", "theme_scripture", "response_hymn")
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

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
    "Preferred: let the server retrieve last week's deck from the church OneDrive "
    "automatically. Call generate_worship_powerpoint with service_date; the server "
    "copies the matters slides, writes the new deck back to the archive, and returns "
    "uploaded_share_url. Explicit last_week_pptx_url or last_week_pptx_base64 "
    "overrides automatic retrieval.\n"
    "If automatic mode is unavailable, follow these manual instructions:\n\n"
)

MATTERS_CANDIDATES_PROMPT = (
    "Python 未能在上周 PPT 中自动识别事项汇报页（旧模板可能没有名为 matters 的版式）。"
    "请查看返回的 matters_candidates：每项包含 slide_number（从 1 开始的页码）、"
    "layout、text_preview，以及是否有图片/表格/图表/分组。"
    "根据这些内容判断哪些是上周的事项汇报/公告页，然后用同样的 last_week_pptx_url "
    "或 last_week_pptx_base64，并传入 last_week_slide_numbers=[页码,...] 再次调用 "
    "generate_worship_powerpoint。若候选都不正确，请直接读取上周 PPT 的内容，"
    "按事项汇报的实际内容选择页码；不要编造事项内容。"
)

MATTERS_NO_CANDIDATES_PROMPT = (
    "Python 未能在上周 PPT 中找到事项汇报页，也无法列出候选页。"
    "请直接读取上周 PPT，按内容找到上周的事项汇报/公告页，并用 "
    "last_week_slide_numbers=[页码,...] 再次调用 generate_worship_powerpoint；"
    "若无法读取文件，请说明具体原因，不要编造事项内容。"
)

AUTO_LAST_WEEK_REMINDER = (
    "未能自动从 OneDrive 读取上周 PPT：{error}。本次生成的事项汇报页留空。"
    "请根据 last_week_archive_candidates 选择文件，或手动上传上周 PPT/提供下载链接。"
)

AUTO_LAST_WEEK_PROMPT = (
    "自动读取 OneDrive 未唯一确定上周文件。请查看 last_week_archive_candidates"
    "（name、date、modified、size），选择正确文件名后用 last_week_filename=\"...\""
    "并保留相同 service_date 重试；如果候选都不正确，请用户手动上传或提供链接，"
    "不要编造事项内容。"
)

AUTO_MATTERS_CANDIDATES_PROMPT = (
    "服务器已从 OneDrive 获取上周 PPT，但未能自动识别事项汇报页"
    "（旧模板可能没有 matters 版式）。请查看 matters_candidates 的 "
    "slide_number、layout、text_preview 和内容标记，选择正确页码；保留相同 "
    "service_date（以及可选的 last_week_filename），传入 "
    "last_week_slide_numbers=[...] 重新调用 generate_worship_powerpoint。"
    "不要编造事项内容。"
)

AUTO_MATTERS_NO_CANDIDATES_PROMPT = (
    "服务器已从 OneDrive 获取上周 PPT，但未能找到事项汇报页或候选页。"
    "请用户手动提供上周文件或页码；不要编造事项内容。"
)

UPLOAD_SUCCESS_REMINDER = (
    "已把本次 PPT 写回 OneDrive 归档：{filename}。分享链接：{share_url}"
)

UPLOAD_FAILED_REMINDER = "OneDrive 写回失败：{error}"

MCP_INSTRUCTIONS = (
    "WorshipPPT Presentation Server.\n"
    "Generates Sunday worship service presentations (PPTX) from structured JSON worship data.\n"
    "AUTOMATIC ONEDRIVE MODE (preferred): call generate_worship_powerpoint with service_date only. "
    "The server looks in the church archive for the strictly previous Sunday's <YYYY-MM-DD>.pptx, "
    "copies its matters/announcements slides, uploads the new deck back to the archive as "
    "<service_date>.pptx, creates an anonymous view link, and returns uploaded_share_url. "
    "After a successful generation, present uploaded_share_url to the user; it is a no-login "
    "download=1 link that downloads the PPTX directly. uploaded_view_url may require a Microsoft sign-in. "
    "Do not ask the user to visit OneDrive or use browser download history when automatic mode succeeds.\n"
    "If the response has last_week_archive_candidates or auto_last_week_error, choose a filename and retry "
    "with last_week_filename=\"...\", or provide an explicit last_week_pptx_url/last_week_pptx_base64. "
    "If matters_candidates is returned for an older template, retry with last_week_slide_numbers=[...]. "
    "Do not invent announcement content.\n"
    "If automatic mode is configured but returns last_week_archive_candidates or auto_last_week_error, choose a "
    "filename and retry, or ask the user for a manual upload or download link. "
    "If automatic mode is not configured, the response includes manual instructions; only then use the archive "
    f"link ({WORSHIP_ARCHIVE_URL}) or browser download history (Ctrl+J / chrome://downloads). "
    "The server never stores the OneDrive password; a dedicated account authorizes it once through device code."
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
        "X-Worship-Last-Week-Metadata",
        "X-Worship-Upload-Metadata",
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


def _last_week_pptx_from_inputs(url: str, encoded: str) -> bytes | None:
    """Resolve the optional last-week deck from a URL or base64 payload."""
    if encoded:
        return _decode_pptx_base64(encoded)
    if url:
        return _download_pptx(url)
    return None


def _copy_matters_into(
    output_path: str | Path,
    last_week_pptx: bytes | None,
    slide_numbers: list[int] | None = None,
) -> tuple[int, list[str]]:
    if not last_week_pptx:
        return 0, []
    return copy_matters_slides(output_path, last_week_pptx, slide_numbers)


def _matters_candidates(last_week_pptx: bytes | None) -> list[dict]:
    if not last_week_pptx:
        return []
    try:
        return list_matters_candidates(Presentation(BytesIO(last_week_pptx)))
    except Exception:
        return []


@dataclass
class LastWeekSelection:
    content: bytes | None = None
    source: str = "none"
    filename: str = ""
    explicit: bool = False
    auto_attempted: bool = False
    auto_error: str = ""
    archive_candidates: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class DeckResult:
    filename: str
    content: bytes
    warnings: list[str]
    matters_copied: int
    matters_candidates: list[dict]
    selection: LastWeekSelection


def _parse_service_date(value: str) -> date:
    raw = (value or "").strip()
    if not raw:
        return date.today()
    try:
        return date.fromisoformat(raw)
    except ValueError:
        raise ValueError("service_date must be YYYY-MM-DD.") from None


def _resolve_last_week_selection(
    service_date: date,
    *,
    uploaded_bytes: bytes | None = None,
    uploaded_name: str = "",
    url: str = "",
    encoded: str = "",
    auto_last_week: bool = True,
    archive_filename: str = "",
) -> LastWeekSelection:
    """Resolve an explicit last-week deck first, then try OneDrive."""
    if uploaded_bytes is not None:
        return LastWeekSelection(
            content=_validate_pptx_bytes(uploaded_bytes),
            source="upload",
            filename=uploaded_name,
            explicit=True,
        )
    if encoded.strip():
        return LastWeekSelection(
            content=_decode_pptx_base64(encoded.strip()),
            source="base64",
            explicit=True,
        )
    if url.strip():
        return LastWeekSelection(
            content=_download_pptx(url.strip()),
            source="url",
            explicit=True,
        )
    if not auto_last_week:
        return LastWeekSelection()

    result = onedrive.auto_last_week_pptx(
        service_date, filename=archive_filename
    )
    if result.content:
        return LastWeekSelection(
            content=result.content,
            source="onedrive",
            filename=result.filename,
            auto_attempted=result.attempted,
            warnings=list(result.warnings),
        )
    return LastWeekSelection(
        auto_attempted=result.attempted,
        auto_error=result.error,
        archive_candidates=list(result.candidates),
        warnings=list(result.warnings),
    )


def _generate_deck(
    payload: dict,
    selected_date: date,
    selection: LastWeekSelection,
    slide_numbers: list[int] | None,
) -> DeckResult:
    with tempfile.TemporaryDirectory(prefix="worship-ppt-") as directory:
        template = Path(directory) / "template.pptx"
        output = Path(directory) / f"Worship_{selected_date.isoformat()}.pptx"
        template.write_bytes(DEFAULT_TEMPLATE_PATH.read_bytes())
        warnings = generate_worship_ppt(
            str(template),
            payload,
            str(output),
            selected_date=selected_date.isoformat(),
        )
        all_warnings = list(warnings or []) + list(selection.warnings)
        matters_copied = 0
        matters_candidates: list[dict] = []
        if selection.content:
            try:
                matters_copied, matters_warnings = _copy_matters_into(
                    output, selection.content, slide_numbers
                )
                all_warnings.extend(matters_warnings)
            except Exception as exc:
                if selection.source == "onedrive":
                    all_warnings.append(
                        f"无法复制 OneDrive 上周文件的事项页：{exc}"
                    )
                else:
                    raise ValueError(
                        f"Could not copy last week's matters slide: {exc}"
                    ) from exc
            if matters_copied == 0:
                matters_candidates = _matters_candidates(selection.content)
        return DeckResult(
            filename=output.name,
            content=output.read_bytes(),
            warnings=all_warnings,
            matters_copied=matters_copied,
            matters_candidates=matters_candidates,
            selection=selection,
        )


def _upload_deck(
    selected_date: date, content: bytes, *, enabled: bool = True
) -> dict:
    empty = {
        "uploaded": False,
        "filename": "",
        "web_url": "",
        "view_url": "",
        "share_url": "",
        "error": "",
        "warnings": [],
    }
    if not enabled:
        return empty
    archive = onedrive.get_onedrive_archive()
    if archive is None:
        return empty
    return archive.upload_generated_deck(selected_date, content)


def _build_result_text(
    deck: DeckResult, upload: dict
) -> tuple[list[str], str, str]:
    selection = deck.selection
    all_warnings = list(deck.warnings)
    all_warnings.extend(upload.get("warnings") or [])
    reminder = ""
    post_generation_prompt = ""

    if deck.matters_copied:
        reminder = MATTERS_COPIED_REMINDER.format(count=deck.matters_copied)
    elif selection.content:
        if deck.matters_candidates:
            reminder = (
                "未能自动识别上周 PPT 的事项汇报页。请查看 "
                "matters_candidates 并选择正确页码后重试。"
            )
            post_generation_prompt = (
                AUTO_MATTERS_CANDIDATES_PROMPT
                if selection.source == "onedrive"
                else MATTERS_CANDIDATES_PROMPT
            )
        else:
            reminder = "未能找到上周 PPT 的事项汇报页，请按内容手动选择页码。"
            post_generation_prompt = (
                AUTO_MATTERS_NO_CANDIDATES_PROMPT
                if selection.source == "onedrive"
                else MATTERS_NO_CANDIDATES_PROMPT
            )
    elif selection.auto_attempted:
        message = AUTO_LAST_WEEK_REMINDER.format(
            error=selection.auto_error or "没有找到唯一的上周文件"
        )
        all_warnings.append(message)
        reminder = message
        post_generation_prompt = AUTO_LAST_WEEK_PROMPT
    else:
        all_warnings.append(POST_GENERATION_REMINDER)
        reminder = POST_GENERATION_REMINDER
        post_generation_prompt = LAST_WEEK_DOWNLOAD_LINK_GUIDE

    if upload.get("uploaded"):
        suffix = UPLOAD_SUCCESS_REMINDER.format(
            filename=upload.get("filename") or "",
            share_url=upload.get("share_url") or upload.get("web_url") or "",
        )
        reminder = f"{reminder}\n\n{suffix}" if reminder else suffix
    elif upload.get("error"):
        suffix = UPLOAD_FAILED_REMINDER.format(error=upload["error"])
        all_warnings.append(suffix)
        reminder = f"{reminder}\n\n{suffix}" if reminder else suffix

    return all_warnings, reminder, post_generation_prompt


def _last_week_metadata(selection: LastWeekSelection) -> str:
    return json.dumps(
        {
            "source": selection.source,
            "filename": selection.filename,
            "auto_attempted": selection.auto_attempted,
            "auto_error": selection.auto_error,
            "candidates": selection.archive_candidates,
        },
        ensure_ascii=True,
        separators=(",", ":"),
    )


def _upload_metadata(upload: dict) -> str:
    return json.dumps(
        {
            "uploaded": bool(upload.get("uploaded")),
            "filename": upload.get("filename") or "",
            "web_url": upload.get("web_url") or "",
            "view_url": upload.get("view_url") or "",
            "share_url": upload.get("share_url") or "",
            "error": upload.get("error") or "",
            "warnings": upload.get("warnings") or [],
        },
        ensure_ascii=True,
        separators=(",", ":"),
    )


@mcp_server.tool()
def get_sample_worship_data() -> dict:
    """Get example worship JSON. Sections accept legacy lines or explicit pages.

    For songs prefer pages: [{"lines": ["lyric 1", "lyric 2"], "footer": "间奏"}].
    Each entry is exactly one slide; footer is optional (间奏/前奏/尾声/结束).
    Do not also supply section-level lines when using pages. Blank lines do not
    define slide boundaries in the legacy lines format.

    MATTERS/AUTOMATION NOTE:
    When OneDrive is configured, generate_worship_powerpoint automatically reads the
    strictly previous Sunday's <YYYY-MM-DD>.pptx from the church archive, copies its
    matters/announcements slides, writes the new deck back, and returns
    uploaded_share_url. Explicit last_week_pptx_url/base64 inputs override automatic
    retrieval. If automatic mode is unavailable, the response includes the archive
    link and manual fallback instructions.
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
    last_week_slide_numbers: list[int] | None = None,
    use_onedrive: bool = True,
    last_week_filename: str = "",
    upload_to_onedrive: bool = True,
) -> dict:
    """Generate a PPTX with the bundled template. Decode pptx_base64 to save the file.

    Automatic mode (preferred): pass service_date and let the server find the
    strictly previous Sunday's <YYYY-MM-DD>.pptx in the church OneDrive archive,
    copy its matters/announcements slides, upload the final deck back to the
    archive, and create an anonymous view link. The response includes
    last_week_source, last_week_filename, uploaded_share_url, uploaded_view_url,
    and related fields. uploaded_share_url is a no-login download=1 link; present it
    to the user when it is present. uploaded_view_url is the OneDrive view link and
    may require a Microsoft sign-in.

    Explicit last_week_pptx_url or last_week_pptx_base64 overrides automatic
    retrieval. If the archive has multiple matching files, pass last_week_filename
    to choose one. Older templates may return matters_candidates; retry with
    last_week_slide_numbers=[...]. Do not invent announcement content.

    If automatic mode is unavailable, the response includes manual-copy
    instructions and the browser-download-history fallback.
    """
    error = _validate_data(data)
    if error:
        raise ValueError(error)
    selected_date = _parse_service_date(service_date)
    selection = _resolve_last_week_selection(
        selected_date,
        url=last_week_pptx_url,
        encoded=last_week_pptx_base64,
        auto_last_week=use_onedrive,
        archive_filename=last_week_filename,
    )
    if last_week_slide_numbers and not selection.content:
        raise ValueError(
            "last_week_slide_numbers requires a last-week deck. Automatic OneDrive "
            "retrieval did not produce one; provide last_week_pptx_url/base64 or "
            "choose last_week_filename first."
        )

    deck = _generate_deck(data, selected_date, selection, last_week_slide_numbers)
    upload = _upload_deck(
        selected_date, deck.content, enabled=upload_to_onedrive
    )
    all_warnings, reminder, post_generation_prompt = _build_result_text(
        deck, upload
    )

    return {
        "filename": deck.filename,
        "mime_type": PPTX_MIME,
        "pptx_base64": base64.b64encode(deck.content).decode("ascii"),
        "warnings": all_warnings,
        "reminder": reminder,
        "post_generation_prompt": post_generation_prompt,
        "matters_copied": deck.matters_copied > 0,
        "matters_slides_copied": deck.matters_copied,
        "matters_candidates": deck.matters_candidates,
        "last_week_source": selection.source,
        "last_week_filename": selection.filename,
        "auto_last_week_attempted": selection.auto_attempted,
        "auto_last_week_error": selection.auto_error,
        "last_week_archive_candidates": selection.archive_candidates,
        "uploaded_to_onedrive": bool(upload.get("uploaded")),
        "uploaded_filename": upload.get("filename") or "",
        "uploaded_web_url": upload.get("web_url") or "",
        "uploaded_view_url": upload.get("view_url") or "",
        "uploaded_share_url": upload.get("share_url") or "",
        "upload_error": upload.get("error") or "",
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
    last_week_slide_numbers: str = Form(default=""),
    use_onedrive: bool = Form(default=True),
    last_week_filename: str = Form(default=""),
    upload_to_onedrive: bool = Form(default=True),
):
    """Generate a worship PPT from the bundled template and uploaded JSON file."""
    try:
        parsed_date = _parse_service_date(selected_date)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

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

    uploaded_bytes = None
    uploaded_name = ""
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
            uploaded_bytes = _validate_pptx_bytes(raw)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        uploaded_name = last_week_pptx.filename

    try:
        selection = await run_in_threadpool(
            _resolve_last_week_selection,
            parsed_date,
            uploaded_bytes=uploaded_bytes,
            uploaded_name=uploaded_name,
            url=last_week_pptx_url,
            auto_last_week=use_onedrive,
            archive_filename=last_week_filename,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    slide_numbers = None
    if last_week_slide_numbers.strip():
        try:
            slide_numbers = [
                int(part)
                for part in last_week_slide_numbers.split(",")
                if part.strip()
            ]
        except ValueError as exc:
            raise HTTPException(
                status_code=400,
                detail="last_week_slide_numbers must be comma-separated integers.",
            ) from exc
        if not slide_numbers:
            raise HTTPException(
                status_code=400,
                detail="last_week_slide_numbers must contain at least one number.",
            )
        if not selection.content:
            raise HTTPException(
                status_code=400,
                detail="last_week_slide_numbers requires a last-week deck. "
                "Automatic OneDrive retrieval did not produce one; upload a file, "
                "provide a URL, or choose last_week_filename first.",
            )

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

    try:
        deck = await run_in_threadpool(
            _generate_deck,
            payload,
            parsed_date,
            selection,
            slide_numbers,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    upload = await run_in_threadpool(
        _upload_deck,
        parsed_date,
        deck.content,
        enabled=upload_to_onedrive,
    )
    all_warnings, _reminder, _post_prompt = _build_result_text(deck, upload)

    response = StreamingResponse(BytesIO(deck.content), media_type=PPTX_MIME)
    response.headers["Content-Disposition"] = (
        f'attachment; filename="{deck.filename}"'
    )
    response.headers["X-Worship-Matters-Copied"] = str(deck.matters_copied)
    if all_warnings:
        response.headers["X-Worship-Warnings"] = json.dumps(all_warnings)
    response.headers["X-Worship-Last-Week-Metadata"] = _last_week_metadata(selection)
    response.headers["X-Worship-Upload-Metadata"] = _upload_metadata(upload)
    return response


@app.get("/api/onedrive/status")
def onedrive_status():
    """Report OneDrive configuration without exposing tokens or URLs."""
    return onedrive.get_status()


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
