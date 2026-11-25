import base64
import copy
import hashlib
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE

from src.backend.app import (
    _is_allowed_onedrive_url,
    _validate_data,
    generate_worship_powerpoint,
    get_sample_worship_data,
)
from src.backend.content import section_pages, split_lines, validate_section
from src.backend.matters import copy_matters_slides
from src.backend.worship import generate_worship_ppt


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
                self.assertFalse(result["matters_copied"])
                self.assertEqual(result["matters_slides_copied"], 0)
                self.assertIn("1drv.ms", result["reminder"])
                self.assertTrue(any("1drv.ms" in w for w in result["warnings"]))
                self.assertIn("chrome://downloads", result["post_generation_prompt"])
                self.assertIn("download.aspx?UniqueId=", result["post_generation_prompt"])
                self.assertIn("chrome://downloads", result["reminder"])
                self.assertTrue(any("chrome://downloads" in w for w in result["warnings"]))
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
                    # Check font size is 28 pt
                    run = p.runs[0]
                    self.assertEqual(run.font.size, Pt(28))
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


class MattersCopyTests(unittest.TestCase):
    PNG_1PX = (
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8A"
        "AQUBAScY42YAAAAASUVORK5CYII="
    )

    def _build_last_week_deck(self, directory: Path, page_count: int = 1) -> Path:
        image_path = directory / "pixel.png"
        image_path.write_bytes(base64.b64decode(self.PNG_1PX))

        prs = Presentation(str(ROOT / "data/template.pptx"))
        layout = next(l for l in prs.slide_layouts if l.name == "matters")
        for index in range(page_count):
            slide = prs.slides.add_slide(layout)
            body = next(
                p for p in slide.placeholders if p.placeholder_format.idx == 10
            )
            body.text = f"MATTERS_TEXT_{index}"

            if index == 0:
                slide.shapes.add_picture(
                    str(image_path), 914400, 2743200, 914400, 914400
                )
                table = slide.shapes.add_table(
                    2, 2, 1828800, 2743200, 2743200, 914400
                ).table
                for row in range(2):
                    for column in range(2):
                        table.cell(row, column).text = f"T{row}{column}"

                group = slide.shapes.add_group_shape()
                group.shapes.add_textbox(4572000, 2743200, 1828800, 457200).text = (
                    "GROUP_TEXT"
                )
                group.shapes.add_picture(
                    str(image_path), 4572000, 3200400, 914400, 914400
                )

                chart_data = ChartData()
                chart_data.categories = ["A", "B"]
                chart_data.add_series("Series 1", (1, 2))
                slide.shapes.add_chart(
                    XL_CHART_TYPE.COLUMN_CLUSTERED,
                    914400,
                    4572000,
                    3657600,
                    2286000,
                    chart_data,
                )
            else:
                table = slide.shapes.add_table(
                    1, 1, 914400, 2743200, 2743200, 457200
                ).table
                table.cell(0, 0).text = f"TABLE_{index}"

        path = directory / "last_week.pptx"
        prs.save(str(path))
        return path

    def _generate_target(self, directory: Path) -> Path:
        target = directory / "target.pptx"
        generate_worship_ppt(
            str(ROOT / "data/template.pptx"),
            get_sample_worship_data(),
            str(target),
            "2026-11-29",
        )
        return target

    @staticmethod
    def _matters_slides(prs):
        return [
            slide
            for slide in prs.slides
            if slide.slide_layout.name.strip().lower() == "matters"
        ]

    def test_copy_preserves_pictures_tables_groups_and_charts(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            source = self._build_last_week_deck(directory)
            target = self._generate_target(directory)

            copied, warnings = copy_matters_slides(target, source.read_bytes())
            self.assertEqual(copied, 1)
            self.assertFalse(any("No announcement" in w for w in warnings))

            prs = Presentation(str(target))
            matters = self._matters_slides(prs)
            self.assertEqual(len(matters), 1)
            slide = matters[0]

            texts = [
                shape.text
                for shape in slide.shapes
                if shape.has_text_frame and shape.text.strip()
            ]
            self.assertIn("MATTERS_TEXT_0", texts)

            pictures = [
                shape
                for shape in slide.shapes
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
            ]
            self.assertEqual(len(pictures), 1)

            tables = [shape for shape in slide.shapes if shape.has_table]
            self.assertEqual(len(tables), 1)
            self.assertEqual(tables[0].table.cell(0, 0).text, "T00")

            groups = [
                shape
                for shape in slide.shapes
                if shape.shape_type == MSO_SHAPE_TYPE.GROUP
            ]
            self.assertEqual(len(groups), 1)
            self.assertIn(
                "GROUP_TEXT",
                [
                    shape.text
                    for shape in groups[0].shapes
                    if shape.has_text_frame
                ],
            )
            self.assertEqual(
                sum(
                    1
                    for shape in groups[0].shapes
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
                ),
                1,
            )

            charts = [shape for shape in slide.shapes if shape.has_chart]
            self.assertEqual(len(charts), 1)
            self.assertEqual(
                [series.name for series in charts[0].chart.series], ["Series 1"]
            )
            self.assertIsNotNone(
                charts[0].chart.part.chart_workbook.xlsx_part
            )

            source_hash = hashlib.sha256(
                base64.b64decode(self.PNG_1PX)
            ).hexdigest()
            with zipfile.ZipFile(target) as archive:
                media_hashes = {
                    hashlib.sha256(archive.read(name)).hexdigest()
                    for name in archive.namelist()
                    if name.startswith("ppt/media/")
                }
            self.assertIn(source_hash, media_hashes)

    def test_copy_supports_multiple_matters_slides(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            source = self._build_last_week_deck(directory, page_count=2)
            target = self._generate_target(directory)

            copied, _ = copy_matters_slides(target, source.read_bytes())
            self.assertEqual(copied, 2)

            prs = Presentation(str(target))
            matters = self._matters_slides(prs)
            self.assertEqual(len(matters), 2)
            first_texts = [
                shape.text
                for shape in matters[0].shapes
                if shape.has_text_frame and shape.text.strip()
            ]
            second_texts = [
                shape.text
                for shape in matters[1].shapes
                if shape.has_text_frame and shape.text.strip()
            ]
            self.assertIn("MATTERS_TEXT_0", first_texts)
            self.assertIn("MATTERS_TEXT_1", second_texts)
            second_tables = [
                shape for shape in matters[1].shapes if shape.has_table
            ]
            self.assertEqual(second_tables[0].table.cell(0, 0).text, "TABLE_1")

    def test_missing_matters_slide_returns_warning(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            source_path = directory / "no_matters.pptx"
            Presentation(str(ROOT / "data/template.pptx")).save(str(source_path))
            target = self._generate_target(directory)

            copied, warnings = copy_matters_slides(
                target, source_path.read_bytes()
            )
            self.assertEqual(copied, 0)
            self.assertTrue(any("No announcement" in w for w in warnings))

    def test_mcp_generate_copies_matters_from_base64(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            source = self._build_last_week_deck(directory)

            result = generate_worship_powerpoint(
                get_sample_worship_data(),
                "2026-11-29",
                last_week_pptx_base64=base64.b64encode(
                    source.read_bytes()
                ).decode("ascii"),
            )
            self.assertTrue(result["matters_copied"])
            self.assertEqual(result["matters_slides_copied"], 1)
            self.assertEqual(result["post_generation_prompt"], "")
            self.assertNotIn("默认为空白", result["reminder"])

            output = directory / "mcp_result.pptx"
            output.write_bytes(base64.b64decode(result["pptx_base64"]))
            matters = self._matters_slides(Presentation(str(output)))
            texts = [
                shape.text
                for shape in matters[0].shapes
                if shape.has_text_frame and shape.text.strip()
            ]
            self.assertIn("MATTERS_TEXT_0", texts)

    def test_onedrive_url_allowlist(self):
        self.assertTrue(
            _is_allowed_onedrive_url(
                "https://onedrive.live.com/personal/abc/_layouts/15/"
                "download.aspx?UniqueId=def"
            )
        )
        self.assertTrue(_is_allowed_onedrive_url("https://1drv.ms/f/abc"))
        self.assertTrue(
            _is_allowed_onedrive_url("https://public.dm.files.1drv.com/file")
        )
        self.assertFalse(
            _is_allowed_onedrive_url("http://onedrive.live.com/file")
        )
        self.assertFalse(_is_allowed_onedrive_url("https://example.com/file"))
        self.assertFalse(
            _is_allowed_onedrive_url(
                "https://onedrive.live.com.example.com/file"
            )
        )


if __name__ == "__main__":
    unittest.main()
