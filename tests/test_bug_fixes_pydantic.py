"""Tests for the three Pydantic bug fixes.

Bug 1: get_workbook_metadata now returns ``file_path`` in its result dict.
Bug 2: write_multi_sheet returns a dict with ``file_path`` and ``sheets_created``.
Bug 3: read_comment returns ``cell_ref`` alongside ``text`` and ``author``.
"""

from __future__ import annotations

from openpyxl import Workbook

from mcp_server.tools.comments import add_comment, read_comment
from mcp_server.tools.workbook import (
    get_workbook_metadata,
    write_multi_sheet,
)

# ---------------------------------------------------------------------------
# Bug 1 – get_workbook_metadata returns file_path
# ---------------------------------------------------------------------------


class TestGetWorkbookMetadataFilePath:
    """get_workbook_metadata must include ``file_path`` in its response."""

    def test_file_path_present(self, sample_xlsx: str) -> None:
        meta = get_workbook_metadata(sample_xlsx)
        assert meta.file_path is not None
        assert meta.file_path == sample_xlsx

    def test_active_sheet_type(self, sample_xlsx: str) -> None:
        meta = get_workbook_metadata(sample_xlsx)
        assert meta.active_sheet is not None
        assert isinstance(meta.active_sheet, str) or meta.active_sheet is None

    def test_basic_fields_present(self, sample_xlsx: str) -> None:
        meta = get_workbook_metadata(sample_xlsx)
        assert meta.sheets is not None
        assert meta.named_ranges is not None
        assert isinstance(meta.sheets, list)
        assert isinstance(meta.named_ranges, list)

    def test_empty_workbook(self, empty_xlsx: str) -> None:
        meta = get_workbook_metadata(empty_xlsx)
        assert meta.file_path == empty_xlsx
        assert len(meta.sheets) == 1
        assert meta.sheets[0].name == "Sheet1"

    def test_multi_sheet_workbook(self, tmp_path) -> None:
        path = str(tmp_path / "multi.xlsx")
        wb = Workbook()
        wb.active.title = "Alpha"
        wb.create_sheet("Beta")
        wb.create_sheet("Gamma")
        wb.save(path)
        wb.close()

        meta = get_workbook_metadata(path)
        assert meta.file_path == path
        names = [s.name for s in meta.sheets]
        assert "Alpha" in names
        assert "Beta" in names
        assert "Gamma" in names


# ---------------------------------------------------------------------------
# Bug 2 – write_multi_sheet response structure
# ---------------------------------------------------------------------------


class TestWriteMultiSheetResponse:
    """write_multi_sheet must return ``file_path`` and ``sheets_created``."""

    def test_result_keys(self, tmp_path) -> None:
        path = str(tmp_path / "out.xlsx")
        sheets = [{"name": "S1", "headers": ["A"], "data": [["x"]]}]
        result = write_multi_sheet(path, sheets)
        assert isinstance(result, dict)
        assert "file_path" in result
        assert "sheets_created" in result

    def test_sheets_created_entry_keys(self, tmp_path) -> None:
        path = str(tmp_path / "out.xlsx")
        sheets = [{"name": "S1", "headers": ["H1", "H2"], "data": [["a", "b"]]}]
        result = write_multi_sheet(path, sheets)
        entry = result["sheets_created"][0]
        assert "name" in entry
        assert "header_count" in entry
        assert "row_count" in entry
        assert "column_widths_set" in entry

    def test_single_sheet(self, tmp_path) -> None:
        path = str(tmp_path / "single.xlsx")
        sheets = [{"name": "Only", "headers": ["Col"], "data": [["v1"], ["v2"]]}]
        result = write_multi_sheet(path, sheets)
        assert len(result["sheets_created"]) == 1
        entry = result["sheets_created"][0]
        assert entry["name"] == "Only"
        assert entry["header_count"] == 1
        assert entry["row_count"] == 2

    def test_multiple_sheets(self, tmp_path) -> None:
        path = str(tmp_path / "multi.xlsx")
        sheets = [
            {"name": "Sales", "headers": ["Product", "Revenue"], "data": [["Widget", 100]]},
            {"name": "Costs", "headers": ["Item", "Amount"], "data": [["Rent", 500], ["Utils", 100]]},
        ]
        result = write_multi_sheet(path, sheets)
        assert len(result["sheets_created"]) == 2
        names = [s["name"] for s in result["sheets_created"]]
        assert "Sales" in names
        assert "Costs" in names

    def test_file_path_matches_input(self, tmp_path) -> None:
        path = str(tmp_path / "match.xlsx")
        sheets = [{"name": "S1", "headers": ["X"], "data": []}]
        result = write_multi_sheet(path, sheets)
        assert result["file_path"] == path


# ---------------------------------------------------------------------------
# Bug 3 – read_comment returns cell_ref
# ---------------------------------------------------------------------------


class TestReadCommentCellRef:
    """read_comment must include ``cell_ref`` in the returned dict."""

    def test_read_comment_keys(self, sample_xlsx: str) -> None:
        add_comment(sample_xlsx, "Sheet1", "A1", "hello", "tester")
        result = read_comment(sample_xlsx, "Sheet1", "A1")
        assert result is not None
        assert "cell_ref" in result
        assert "text" in result
        assert "author" in result

    def test_cell_ref_matches_queried_cell(self, sample_xlsx: str) -> None:
        add_comment(sample_xlsx, "Sheet1", "B2", "note", "author1")
        result = read_comment(sample_xlsx, "Sheet1", "B2")
        assert result is not None
        assert result["cell_ref"] == "B2"

    def test_no_comment_returns_none(self, sample_xlsx: str) -> None:
        result = read_comment(sample_xlsx, "Sheet1", "C3")
        assert result is None

    def test_round_trip(self, sample_xlsx: str) -> None:
        add_comment(sample_xlsx, "Sheet1", "D4", "round-trip", "bot")
        result = read_comment(sample_xlsx, "Sheet1", "D4")
        assert result is not None
        assert result["cell_ref"] == "D4"
        assert result["text"] == "round-trip"
        assert result["author"] == "bot"
