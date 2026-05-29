"""FastAPI backend for generating worship PowerPoint files."""

from __future__ import annotations

import json
import tempfile
from datetime import date
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse, StreamingResponse

from src.backend.worship import generate_worship_ppt


BACKEND_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BACKEND_DIR.parent.parent
DIST_DIR = PROJECT_ROOT / "dist"
REQUIRED_SECTIONS = ("call_to_worship", "hymns", "theme_scripture", "response_hymn")

app = FastAPI()


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


@app.post("/api/generate")
async def generate(
    template: UploadFile = File(...),
    data: UploadFile = File(...),
    selected_date: str = Form(default="", alias="date"),
):
    """Generate a worship PPT from uploaded template and JSON files."""
    selected_date = selected_date.strip() or date.today().isoformat()

    if not template.filename or not template.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Template must be a .pptx file.")
    if not data.filename or not data.filename.lower().endswith(".json"):
        raise HTTPException(status_code=400, detail="Worship data must be a .json file.")

    try:
        payload = json.loads((await data.read()).decode("utf-8"))
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid JSON: {exc.msg}.") from exc

    validation_error = _validate_data(payload)
    if validation_error:
        raise HTTPException(status_code=400, detail=validation_error)

    with tempfile.TemporaryDirectory(prefix="worship-ppt-") as tmp_dir_name:
        tmp_dir = Path(tmp_dir_name)
        template_path = tmp_dir / Path(template.filename).name
        output_path = tmp_dir / f"Worship_{selected_date}.pptx"

        template_path.write_bytes(await template.read())
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
