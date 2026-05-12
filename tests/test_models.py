from __future__ import annotations

from mcp_server.models import (
    ChunkReadResult,
    ColumnStats,
    FilterResult,
    SheetInfo,
    SheetSummary,
    WorkbookCreatedResult,
    WorkbookMetadata,
)


def test_sheet_info():
    si = SheetInfo(name="Sheet1", min_row=1, max_row=10, min_col=1, max_col=5)
    assert si.name == "Sheet1"
    assert si.min_row == 1
    si_minimal = SheetInfo(name="Sheet2")
    assert si_minimal.name == "Sheet2"
    assert si_minimal.min_row is None


def test_sheet_summary():
    ss = SheetSummary(name="Sheet1", row_count=10, col_count=5, headers=["A", "B", "C"], used_range="A1:C10")
    assert ss.name == "Sheet1"
    assert ss.row_count == 10
    assert "A" in ss.headers


def test_workbook_created_result():
    wcr = WorkbookCreatedResult(file_path="/path/to/file.xlsx", sheets=["Sheet1", "Sheet2"])
    assert wcr.file_path == "/path/to/file.xlsx"
    assert len(wcr.sheets) == 2


def test_filter_result():
    fr = FilterResult(matched_rows=[[1, 2]], total_rows=10, matched_count=1, headers=["A", "B"])
    assert fr.matched_count == 1
    assert fr.total_rows == 10


def test_column_stats():
    cs = ColumnStats(column="A", count=5, mean=2.5, sum_val=12.5)
    assert cs.column == "A"
    assert cs.mean == 2.5
    cs_empty = ColumnStats(column="B", count=0)
    assert cs_empty.mean is None


def test_workbook_metadata():
    si = SheetInfo(name="Sheet1")
    wm = WorkbookMetadata(
        file_path="/path/test.xlsx", sheets=[si], active_sheet="Sheet1", named_ranges=[{"name": "test", "range": "A1"}]
    )
    assert wm.active_sheet == "Sheet1"
    assert wm.sheets[0].name == "Sheet1"


def test_chunk_read_result():
    crr = ChunkReadResult(rows=[{"A": 1}], chunk_start=0, chunk_size=1, has_more=True, next_start_row=1)
    assert crr.has_more is True
    assert crr.next_start_row == 1
