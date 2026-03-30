from __future__ import annotations

import pytest
from pydantic import ValidationError

from mcp_server.models import (
    AggregateResult,
    CellValue,
    ChartConfig,
    ChunkReadResult,
    ColumnStats,
    ConditionalFormatRule,
    CsvPreview,
    DataProfile,
    DuplicateResult,
    FilterResult,
    FormatOptions,
    PivotResult,
    RangeData,
    SheetInfo,
    SheetSummary,
    SortCriteria,
    ValidationResult,
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


def test_cell_value():
    cv = CellValue(cell_ref="A1", value="Hello", data_type="s")
    assert cv.cell_ref == "A1"
    assert cv.value == "Hello"


def test_range_data():
    rd = RangeData(rows=[[1, 2], [3, 4]], row_count=2, col_count=2)
    assert rd.row_count == 2
    assert rd.rows[0][0] == 1


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


def test_aggregate_result():
    ar = AggregateResult(groups=[{"Category": "A", "Sum": 10}], group_by="Category", operation="sum")
    assert ar.group_by == "Category"
    assert ar.operation == "sum"


def test_duplicate_result():
    dr = DuplicateResult(duplicates=[[1, 1]], count=1, headers=["ID", "Val"])
    assert dr.count == 1


def test_csv_preview():
    cp = CsvPreview(headers=["A", "B"], rows=[[1, 2]], total_rows=100)
    assert cp.total_rows == 100
    assert len(cp.headers) == 2


def test_workbook_metadata():
    si = SheetInfo(name="Sheet1")
    wm = WorkbookMetadata(
        file_path="/path/test.xlsx", sheets=[si], active_sheet="Sheet1", named_ranges=[{"name": "test", "range": "A1"}]
    )
    assert wm.active_sheet == "Sheet1"
    assert wm.sheets[0].name == "Sheet1"


def test_data_profile():
    dp = DataProfile(
        columns=["A"], row_count=1, missing_values={"A": 0}, duplicates=0, summary_statistics={"A": {"mean": 1.0}}
    )
    assert dp.row_count == 1


def test_pivot_result():
    pr = PivotResult(data=[{"A": 1}], index_columns=["Col1"], value_columns=["Col2"], operation="sum")
    assert pr.operation == "sum"


def test_validation_result():
    vr = ValidationResult(valid=True, tokens=[{"type": "num", "value": "1"}])
    assert vr.valid is True
    vr_err = ValidationResult(valid=False, tokens=[], error="Invalid")
    assert vr_err.valid is False
    assert vr_err.error == "Invalid"


def test_chunk_read_result():
    crr = ChunkReadResult(rows=[{"A": 1}], chunk_start=0, chunk_size=1, has_more=True, next_start_row=1)
    assert crr.has_more is True
    assert crr.next_start_row == 1


def test_format_options():
    fo = FormatOptions(bold=True, font_size=12)
    assert fo.bold is True
    assert fo.font_size == 12
    assert fo.italic is False


def test_conditional_format_rule():
    cfr = ConditionalFormatRule(format_type="color_scale", start_color="FF0000", end_color="00FF00")
    assert cfr.format_type == "color_scale"

    with pytest.raises(ValidationError):
        ConditionalFormatRule(format_type="invalid_type")


def test_chart_config():
    cc = ChartConfig(chart_type="bar", title="My Chart")
    assert cc.chart_type == "bar"
    assert cc.title == "My Chart"
    assert cc.style == 10  # default


def test_sort_criteria():
    sc = SortCriteria(column="A")
    assert sc.column == "A"
