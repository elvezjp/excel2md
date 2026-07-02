"""Tests for WorkbookDrawingIndex (Issue #26).

The index consolidates the per-sheet DrawingML resolution that used to be
duplicated in image_extraction and mermaid_generator, parsing the xlsx ZIP
once per workbook.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from excel2md.drawing_index import WorkbookDrawingIndex

FIXTURES = Path(__file__).parent / "fixtures"


class TestWorkbookDrawingIndex:
    def test_resolves_drawing_path_for_sheet_with_images(self):
        with WorkbookDrawingIndex(str(FIXTURES / "test_standard.xlsx")) as index:
            drawing_path = index.drawing_path("画像")
            assert drawing_path is not None
            assert drawing_path.startswith("xl/drawings/")
            assert index.exists(drawing_path)
            assert index.read(drawing_path).startswith(b"<?xml")

    def test_image_rels_map_rid_to_media_path(self):
        with WorkbookDrawingIndex(str(FIXTURES / "test_standard.xlsx")) as index:
            image_rels = index.image_rels("画像")
            assert image_rels, "expected image relationships on the 画像 sheet"
            for media_path in image_rels.values():
                assert media_path.startswith("xl/media/")
                assert index.exists(media_path)

    def test_sheet_without_drawing_returns_none(self):
        with WorkbookDrawingIndex(str(FIXTURES / "test_standard.xlsx")) as index:
            assert index.drawing_path("基本テーブル") is None
            assert index.image_rels("基本テーブル") == {}

    def test_resolves_drawing_for_shapes_sheet(self):
        with WorkbookDrawingIndex(str(FIXTURES / "test_mermaid.xlsx")) as index:
            assert index.drawing_path("図形フロー") is not None
            # Shapes-only drawing: no image relationships
            assert index.image_rels("図形フロー") == {}

    def test_unknown_sheet_returns_empty(self):
        with WorkbookDrawingIndex(str(FIXTURES / "test_standard.xlsx")) as index:
            assert index.drawing_path("存在しないシート") is None
            assert index.image_rels("存在しないシート") == {}

    def test_unopenable_file_behaves_as_empty_index(self, tmp_path):
        bogus = tmp_path / "not_a_zip.xlsx"
        bogus.write_bytes(b"this is not a zip file")
        with WorkbookDrawingIndex(str(bogus)) as index:
            assert index.drawing_path("Sheet1") is None
            assert index.image_rels("Sheet1") == {}
            assert not index.exists("xl/workbook.xml")
