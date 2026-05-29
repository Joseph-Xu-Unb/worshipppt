"""FastAPI backend for generating worship PowerPoint files."""

from __future__ import annotations

import json
import os
import tempfile
from datetime import date
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse, StreamingResponse

from src.backend.worship import generate_worship_ppt


BACKEND_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BACKEND_DIR.parent.parent
DIST_DIR = PROJECT_ROOT / "dist"
DATA_DIR = PROJECT_ROOT / "data"
DEFAULT_TEMPLATE_PATH = DATA_DIR / "template.pptx"
SAMPLE_JSON_PATH = DATA_DIR / "sample_worship_data.json"
SAMPLE_PPT_PATH = DATA_DIR / "2026-05-17.pptx"
REQUIRED_SECTIONS = ("call_to_worship", "hymns", "theme_scripture", "response_hymn")

app = FastAPI()

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
    expose_headers=["Content-Disposition", "X-Worship-Warnings"],
)


def _validate_data(data):
    """Return a validation error string, or None when input data is usable."""
    if not isinstance(data, dict):
        return "JSON root must be an object."

    missing = [key for key in REQUIRED_SECTIONS if key not in data]
    if missing:
        return f"Missing required JSON section(s): {', '.join(missing)}."

    for section in ("call_to_worship", "theme_scripture", "response_hymn"):
        item = data[section]
        if not isinstance(item, dict):
            return f"{section} must be an object."
        if not isinstance(item.get("title"), str):
            return f"{section}.title must be a string."
        if not isinstance(item.get("lines"), list):
            return f"{section}.lines must be a list."

    if not isinstance(data["hymns"], list) or not data["hymns"]:
        return "hymns must be a non-empty list."

    for index, hymn in enumerate(data["hymns"], start=1):
        if not isinstance(hymn, dict):
            return f"hymns[{index}] must be an object."
        if not isinstance(hymn.get("title"), str):
            return f"hymns[{index}].title must be a string."
        if not isinstance(hymn.get("lines"), list):
            return f"hymns[{index}].lines must be a list."

    return None


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
        ppt_bytes = BytesIO(output_path.read_bytes())
        response = StreamingResponse(
            ppt_bytes,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        response.headers["Content-Disposition"] = (
            f'attachment; filename="{output_path.name}"'
        )
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
