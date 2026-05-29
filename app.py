"""Web backend for generating worship PowerPoint files."""

from __future__ import annotations

import json
import tempfile
from io import BytesIO
from datetime import date
from pathlib import Path

from flask import Flask, jsonify, request, send_file, send_from_directory
from werkzeug.utils import secure_filename

from worship import generate_worship_ppt


BASE_DIR = Path(__file__).resolve().parent
DIST_DIR = BASE_DIR / "dist"
REQUIRED_SECTIONS = ("call_to_worship", "hymns", "theme_scripture", "response_hymn")

app = Flask(__name__, static_folder=None)


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
    return jsonify({"ok": True})


@app.post("/api/generate")
def generate():
    """Generate a worship PPT from uploaded template and JSON files."""
    template_file = request.files.get("template")
    json_file = request.files.get("data")
    selected_date = request.form.get("date", "").strip() or date.today().isoformat()

    if not template_file:
        return jsonify({"error": "PPT template file is required."}), 400
    if not json_file:
        return jsonify({"error": "JSON worship data file is required."}), 400
    if not template_file.filename.lower().endswith(".pptx"):
        return jsonify({"error": "Template must be a .pptx file."}), 400
    if not json_file.filename.lower().endswith(".json"):
        return jsonify({"error": "Worship data must be a .json file."}), 400

    try:
        data = json.load(json_file.stream)
    except json.JSONDecodeError as exc:
        return jsonify({"error": f"Invalid JSON: {exc.msg}."}), 400

    validation_error = _validate_data(data)
    if validation_error:
        return jsonify({"error": validation_error}), 400

    with tempfile.TemporaryDirectory(prefix="worship-ppt-") as tmp_dir_name:
        tmp_dir = Path(tmp_dir_name)
        template_path = tmp_dir / secure_filename(template_file.filename)
        output_path = tmp_dir / f"Worship_{selected_date}.pptx"

        template_file.save(template_path)
        warnings = generate_worship_ppt(
            str(template_path),
            data,
            str(output_path),
            selected_date=selected_date,
        )
        ppt_bytes = BytesIO(output_path.read_bytes())
        ppt_bytes.seek(0)

        response = send_file(
            ppt_bytes,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=output_path.name,
        )
        if warnings:
            response.headers["X-Worship-Warnings"] = json.dumps(warnings)
        return response


@app.get("/")
def index():
    """Serve the built React app when available."""
    index_path = DIST_DIR / "index.html"
    if not index_path.exists():
        return (
            "React app is not built yet. Run 'npm install' and 'npm run build', "
            "or use 'npm run dev' during development.",
            404,
        )
    return send_from_directory(DIST_DIR, "index.html")


@app.get("/<path:path>")
def static_assets(path):
    """Serve React build assets with SPA fallback."""
    asset_path = DIST_DIR / path
    if asset_path.exists() and asset_path.is_file():
        return send_from_directory(DIST_DIR, path)
    return index()


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
