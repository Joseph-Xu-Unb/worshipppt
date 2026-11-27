"""Canonical Muse prompt and prompt-version handshake.

The prompt text lives in ``src/shared/muse_prompt.txt`` so the web guide and the
MCP server always display the same instructions. The version is a short hash of
the template plus the shared church resource URLs; any prompt or URL change
produces a new version.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
SHARED_DIR = PROJECT_ROOT / "src" / "shared"
PROMPT_PATH = SHARED_DIR / "muse_prompt.txt"
RESOURCES_PATH = SHARED_DIR / "church_resources.json"

RESOURCES = json.loads(RESOURCES_PATH.read_text(encoding="utf-8"))
LYRICS_LIBRARY_URL = str(RESOURCES["lyricsLibraryUrl"])
WORSHIP_ARCHIVE_URL = str(RESOURCES["worshipArchiveUrl"])

_TEMPLATE = PROMPT_PATH.read_text(encoding="utf-8")
MUSE_PROMPT_VERSION = hashlib.sha256(
    f"{_TEMPLATE}\n{LYRICS_LIBRARY_URL}\n{WORSHIP_ARCHIVE_URL}".encode("utf-8")
).hexdigest()[:12]

MUSE_PROMPT = (
    _TEMPLATE.replace("{{LYRICS_LIBRARY_URL}}", LYRICS_LIBRARY_URL)
    .replace("{{WORSHIP_ARCHIVE_URL}}", WORSHIP_ARCHIVE_URL)
    .strip()
)

MUSE_PROMPT_SYNC_GUIDE = (
    "Call check_muse_prompt or get_muse_prompt before generating. If your stored "
    "prompt version differs, replace your stored instructions with the returned "
    "prompt. Then pass muse_prompt_version to generate_worship_powerpoint; missing "
    "or stale versions are rejected."
)
