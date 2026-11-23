import base64
import copy
import json
from pathlib import Path
import tempfile
import unittest

from pptx import Presentation

from src.backend.app import (
    _validate_data,
    generate_worship_powerpoint,
    get_sample_worship_data,
)
from src.backend.content import section_pages, split_lines, validate_section


ROOT = Path(__file__).resolve().parents[1]


class PaginationTests(unittest.TestCase):
    def test_markers_stay_after_full_and_partial_pages(self):
        for marker in ("间奏", "（间奏）", "( 间奏 )", "结束", "（尾声）", "前奏"):
            with self.subTest(marker=marker):
                self.assertEqual(split_lines(["A", "B", "C", "D", marker, "E"]),
                                 [["A", "B", "C", "D", marker], ["E"]])
                self.assertEqual(split_lines(["A", "", "B", marker, "C", "D"]),
                                 [["A", "B", marker], ["C", "D"]])

    def test_markers_do_not_consume_lyric_slots(self):
        self.assertEqual(split_lines(["A", "间奏", "B", "C", "D", "E", "F", "结束"]),
                         [["A", "间奏"], ["B", "C", "D", "E"], ["F", "结束"]])
        self.assertEqual(split_lines(["A", "间奏", "尾声", "B"]),
                         [["A", "间奏", "尾声"], ["B"]])

    def test_orphan_marker_is_rejected(self):
        self.assertIsNotNone(validate_section({"title": "Test", "lines": ["（间奏）", "A"]}, "song"))

    def test_explicit_pages_keep_boundaries(self):
        song = {"title": "Test", "pages": [
            {"lines": ["A", "B"], "footer": "（间奏）"}, {"lines": ["C"]},
            {"lines": ["D", "E"], "footer": "结束"},
        ]}
        self.assertIsNone(validate_section(song, "song"))
        self.assertEqual(section_pages(song), [["A", "B", "（间奏）"], ["C"], ["D", "E", "结束"]])
        song["lines"] = ["different"]
        self.assertIsNotNone(validate_section(song, "song"))

    def test_invalid_pages_are_rejected(self):
        for pages in (None, [], [{}], [{"lines": []}], [{"lines": [" "]}],
                      [{"lines": ["A", "间奏"]}], [{"lines": ["A"], "footer": 3}],
                      [{"lines": ["A"], "footer": "unknown"}]):
            with self.subTest(pages=pages):
                self.assertIsNotNone(validate_section({"title": "Test", "pages": pages}, "song"))

    def test_sample_and_legacy_api_validation(self):
        legacy = json.loads((ROOT / "data/sample_worship_data.json").read_text())
        self.assertIsNone(_validate_data(legacy))
        sample = get_sample_worship_data()
        self.assertIn("pages", sample["hymns"][0])
        self.assertIsNone(_validate_data(sample))
        sample["response_hymn"]["pages"] = [{"lines": ["间奏"]}]
        self.assertIsNotNone(_validate_data(sample))

    def test_actual_pptx_contains_exact_pages_for_both_formats(self):
        base = get_sample_worship_data()
        for explicit in (False, True):
            data = copy.deepcopy(base)
            if explicit:
                song = {"title": "TEST_SONG", "pages": [
                    {"lines": ["TEST_A", "TEST_B"], "footer": "（间奏）"},
                    {"lines": ["TEST_C", "TEST_D"], "footer": "结束"},
                ]}
            else:
                song = {"title": "TEST_SONG", "lines": [
                    "TEST_A", "TEST_B", "（间奏）", "TEST_C", "TEST_D", "结束",
                ]}
            data["hymns"] = [song]
            data["response_hymn"] = song
            self.assertIsNone(_validate_data(data))
            with tempfile.TemporaryDirectory() as directory:
                output = Path(directory) / "test.pptx"
                result = generate_worship_powerpoint(data, "2026-10-25")
                output.write_bytes(base64.b64decode(result["pptx_base64"]))
                self.assertTrue(result["filename"].endswith("2026-10-25.pptx"))
                bodies = []
                for slide in Presentation(output).slides:
                    for shape in slide.shapes:
                        if shape.has_text_frame and (
                            "TEST_A" in shape.text or "TEST_C" in shape.text
                        ):
                            bodies.append(shape.text)
                self.assertEqual(bodies, ["TEST_A\nTEST_B\n（间奏）", "TEST_C\nTEST_D\n结束"] * 2)

    def test_hymn_page_numbers_are_centered_and_formatted_xy(self):
        from pptx.enum.text import PP_ALIGN
        from pptx.util import Pt
        from src.backend.worship import generate_worship_ppt

        base = get_sample_worship_data()
        song1 = {"title": "SONG_ONE", "pages": [
            {"lines": ["S1_P1_L1"]}, {"lines": ["S1_P2_L1"]}, {"lines": ["S1_P3_L1"]},
        ]}
        song2 = {"title": "SONG_TWO", "pages": [
            {"lines": ["S2_P1_L1"]}, {"lines": ["S2_P2_L1"]},
        ]}
        resp = {"title": "RESP_SONG", "pages": [
            {"lines": ["R_P1_L1"]}, {"lines": ["R_P2_L1"]}, {"lines": ["R_P3_L1"]}, {"lines": ["R_P4_L1"]},
        ]}
        data = copy.deepcopy(base)
        data["hymns"] = [song1, song2]
        data["response_hymn"] = resp

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "test_pages.pptx"
            generate_worship_ppt(str(ROOT / "data/template.pptx"), data, str(output), "2026-11-23")
            prs = Presentation(output)

            hymn_page_texts = []
            scripture_page_texts = []

            for slide in prs.slides:
                layout_name = slide.slide_layout.name
                page_ph = None
                for ph in slide.placeholders:
                    if ph.placeholder_format.idx == 12:
                        page_ph = ph
                        break

                if layout_name in ("Hymn", "response"):
                    self.assertIsNotNone(page_ph, f"Slide with layout {layout_name} should have page placeholder")
                    # Check centered alignment
                    p = page_ph.text_frame.paragraphs[0]
                    self.assertEqual(p.alignment, PP_ALIGN.CENTER)
                    # Check font size in 14-18 pt range
                    run = p.runs[0]
                    self.assertGreaterEqual(run.font.size, Pt(14))
                    self.assertLessEqual(run.font.size, Pt(18))
                    # Check horizontal centering
                    center_x = page_ph.left + page_ph.width / 2
                    self.assertAlmostEqual(center_x, prs.slide_width / 2, delta=1000)
                    hymn_page_texts.append(page_ph.text.strip())
                elif layout_name in ("call_to_scripture", "theme_scripture"):
                    if page_ph:
                        scripture_page_texts.append(page_ph.text.strip())

            # Song 1 (3 pages): 1/3, 2/3, 3/3; Song 2 (2 pages): 1/2, 2/2; Resp (4 pages): 1/4, 2/4, 3/4, 4/4
            expected = ["1/3", "2/3", "3/3", "1/2", "2/2", "1/4", "2/4", "3/4", "4/4"]
            self.assertEqual(hymn_page_texts, expected)
            self.assertEqual(scripture_page_texts, [])

    def test_edited_slides_hymn_page_numbers(self):
        from src.backend.worship import generate_worship_ppt

        base = get_sample_worship_data()
        data = copy.deepcopy(base)
        data["preview_slides"] = [
            {"kind": "hymn", "sectionLabel": "Hymn 1", "title": "Hymn1: A", "lines": ["Line 1"]},
            {"kind": "hymn", "sectionLabel": "Hymn 1", "title": "Hymn1: A", "lines": ["Line 2"]},
            {"kind": "hymn", "sectionLabel": "Hymn 2", "title": "Hymn2: B", "lines": ["Line 1"]},
            {"kind": "response_hymn", "sectionLabel": "Response Hymn", "title": "Resp: C", "lines": ["R1"]},
            {"kind": "response_hymn", "sectionLabel": "Response Hymn", "title": "Resp: C", "lines": ["R2"]},
        ]

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "test_edited_pages.pptx"
            generate_worship_ppt(str(ROOT / "data/template.pptx"), data, str(output), "2026-11-23")
            prs = Presentation(output)

            hymn_page_texts = []
            for slide in prs.slides:
                if slide.slide_layout.name in ("Hymn", "response"):
                    for ph in slide.placeholders:
                        if ph.placeholder_format.idx == 12:
                            hymn_page_texts.append(ph.text.strip())

            self.assertEqual(hymn_page_texts, ["1/2", "2/2", "1/1", "1/2", "2/2"])


if __name__ == "__main__":
    unittest.main()
