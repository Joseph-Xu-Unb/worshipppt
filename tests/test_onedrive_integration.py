import base64
import json
import os
from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from fastapi.testclient import TestClient
from pptx import Presentation

from src.backend import app as app_module
from src.backend.app import generate_worship_powerpoint, get_sample_worship_data
from src.backend.onedrive import AutoLastWeekResult, reset_onedrive_cache_for_tests


ROOT = Path(__file__).resolve().parents[1]


def setUpModule():
    for key in list(os.environ):
        if key.startswith("ONEDRIVE_"):
            os.environ.pop(key, None)
    reset_onedrive_cache_for_tests()


def build_last_week_deck() -> bytes:
    with tempfile.TemporaryDirectory() as directory:
        prs = Presentation(str(ROOT / "data/template.pptx"))
        layout = next(
            layout for layout in prs.slide_layouts if layout.name == "matters"
        )
        slide = prs.slides.add_slide(layout)
        body = next(
            placeholder
            for placeholder in slide.placeholders
            if placeholder.placeholder_format.idx == 10
        )
        body.text = "LAST_WEEK_MATTERS"
        path = Path(directory) / "2026-11-22.pptx"
        prs.save(str(path))
        return path.read_bytes()


TEST_DECK = build_last_week_deck()


class StubArchive:
    def __init__(self, uploaded=True, share_url="https://1drv.ms/u/share"):
        self.uploaded = uploaded
        self.share_url = share_url

    def upload_generated_deck(self, service_date, content):
        if not self.uploaded:
            return {
                "uploaded": False,
                "filename": f"{service_date.isoformat()}.pptx",
                "web_url": "",
                "share_url": "",
                "error": "write failed",
                "warnings": [],
            }
        return {
            "uploaded": True,
            "filename": f"{service_date.isoformat()}.pptx",
            "web_url": "https://onedrive.live.com/item",
            "share_url": self.share_url,
            "error": "",
            "warnings": [],
        }


class AutoGenerationTests(unittest.TestCase):
    def test_auto_success_copies_matters(self):
        with patch(
            "src.backend.onedrive.auto_last_week_pptx",
            return_value=AutoLastWeekResult(
                content=TEST_DECK,
                filename="2026-11-22.pptx",
                attempted=True,
            ),
        ):
            result = generate_worship_powerpoint(
                get_sample_worship_data(),
                "2026-11-29",
                upload_to_onedrive=False,
            )

        self.assertTrue(result["matters_copied"])
        self.assertEqual(result["last_week_source"], "onedrive")
        self.assertEqual(result["last_week_filename"], "2026-11-22.pptx")
        self.assertTrue(result["auto_last_week_attempted"])
        self.assertFalse(result["uploaded_to_onedrive"])
        output = BytesIO(base64.b64decode(result["pptx_base64"]))
        texts = [
            shape.text
            for slide in Presentation(output).slides
            for shape in slide.shapes
            if shape.has_text_frame
        ]
        self.assertIn("LAST_WEEK_MATTERS", texts)

    def test_declining_copy_leaves_matters_blank_without_manual_reminder(self):
        result = generate_worship_powerpoint(
            get_sample_worship_data(),
            "2026-11-29",
            use_onedrive=False,
            upload_to_onedrive=False,
        )
        self.assertFalse(result["matters_copied"])
        self.assertEqual(result["last_week_source"], "none")
        self.assertTrue(result["auto_last_week_declined"])
        self.assertIn("不复制", result["reminder"])
        self.assertFalse(
            any("1drv.ms" in warning for warning in result["warnings"])
        )

    def test_auto_failure_still_returns_valid_deck(self):
        with patch(
            "src.backend.onedrive.auto_last_week_pptx",
            return_value=AutoLastWeekResult(
                attempted=True,
                error="archive unavailable",
                candidates=[{"name": "2026-11-22.pptx"}],
            ),
        ):
            result = generate_worship_powerpoint(
                get_sample_worship_data(),
                "2026-11-29",
                upload_to_onedrive=False,
            )

        self.assertFalse(result["matters_copied"])
        self.assertEqual(result["last_week_source"], "none")
        self.assertEqual(result["auto_last_week_error"], "archive unavailable")
        self.assertEqual(
            result["last_week_archive_candidates"][0]["name"], "2026-11-22.pptx"
        )
        self.assertTrue(any("archive unavailable" in w for w in result["warnings"]))
        Presentation(BytesIO(base64.b64decode(result["pptx_base64"])))

    def test_explicit_base64_takes_precedence(self):
        with patch("src.backend.onedrive.auto_last_week_pptx") as auto:
            auto.side_effect = AssertionError("automatic retrieval should not run")
            result = generate_worship_powerpoint(
                get_sample_worship_data(),
                "2026-11-29",
                last_week_pptx_base64=base64.b64encode(TEST_DECK).decode("ascii"),
                upload_to_onedrive=False,
            )
        self.assertTrue(result["matters_copied"])
        self.assertEqual(result["last_week_source"], "base64")

    def test_upload_success_is_returned(self):
        with patch(
            "src.backend.onedrive.auto_last_week_pptx",
            return_value=AutoLastWeekResult(
                content=TEST_DECK,
                filename="2026-11-22.pptx",
                attempted=True,
            ),
        ), patch(
            "src.backend.onedrive.get_onedrive_archive",
            return_value=StubArchive(),
        ):
            result = generate_worship_powerpoint(
                get_sample_worship_data(),
                "2026-11-29",
                upload_to_onedrive=True,
            )
        self.assertTrue(result["uploaded_to_onedrive"])
        self.assertEqual(result["uploaded_filename"], "2026-11-29.pptx")
        self.assertEqual(result["uploaded_share_url"], "https://1drv.ms/u/share")
        self.assertIn("https://1drv.ms/u/share", result["reminder"])

    def test_upload_failure_is_soft(self):
        with patch(
            "src.backend.onedrive.auto_last_week_pptx",
            return_value=AutoLastWeekResult(
                content=TEST_DECK,
                filename="2026-11-22.pptx",
                attempted=True,
            ),
        ), patch(
            "src.backend.onedrive.get_onedrive_archive",
            return_value=StubArchive(uploaded=False),
        ):
            result = generate_worship_powerpoint(
                get_sample_worship_data(),
                "2026-11-29",
                upload_to_onedrive=True,
            )
        self.assertFalse(result["uploaded_to_onedrive"])
        self.assertEqual(result["upload_error"], "write failed")
        self.assertTrue(result["matters_copied"])


class RestTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # FastMCP's session manager may only be started once per app instance.
        cls.client = TestClient(app_module.app)
        cls.client.__enter__()

    @classmethod
    def tearDownClass(cls):
        cls.client.__exit__(None, None, None)

    def test_status_endpoint(self):
        status = {
            "configured": True,
            "mode": "share",
            "auto_enabled": True,
            "upload_enabled": True,
            "token_status": "ok",
            "token_age_days": 1,
            "archive": "OneDrive 归档",
            "error": "",
        }
        with patch("src.backend.onedrive.get_status", return_value=status):
            response = self.client.get("/api/onedrive/status")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json(), status)

    def test_generate_returns_metadata_headers(self):
        sample = json.dumps(get_sample_worship_data()).encode("utf-8")
        with patch(
            "src.backend.onedrive.auto_last_week_pptx",
            return_value=AutoLastWeekResult(
                content=TEST_DECK,
                filename="2026-11-22.pptx",
                attempted=True,
            ),
        ):
            response = self.client.post(
                "/api/generate",
                data={
                    "date": "2026-11-29",
                    "use_onedrive": "true",
                    "upload_to_onedrive": "false",
                },
                files={
                    "data": (
                        "worship.json",
                        sample,
                        "application/json",
                    )
                },
            )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["X-Worship-Matters-Copied"], "1")
        metadata = json.loads(
            response.headers["X-Worship-Last-Week-Metadata"]
        )
        self.assertEqual(metadata["source"], "onedrive")
        self.assertEqual(metadata["filename"], "2026-11-22.pptx")
        upload = json.loads(response.headers["X-Worship-Upload-Metadata"])
        self.assertFalse(upload["uploaded"])
        self.assertTrue(response.content.startswith(b"PK"))

    def test_generate_rejects_invalid_date(self):
        sample = json.dumps(get_sample_worship_data()).encode("utf-8")
        response = self.client.post(
            "/api/generate",
            data={"date": "11/29/2026", "upload_to_onedrive": "false"},
            files={"data": ("worship.json", sample, "application/json")},
        )
        self.assertEqual(response.status_code, 400)


if __name__ == "__main__":
    unittest.main()
