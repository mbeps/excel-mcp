from __future__ import annotations

import openpyxl

from mcp_server.tools.formatting import (
    auto_fit_columns,
    format_cells,
    merge_cells,
    set_column_width,
    set_row_height,
    unmerge_cells,
)


def test_format_cells(sample_xlsx: str) -> None:
    format_cells(sample_xlsx, "Sheet1", "A1:D1", bold=True, font_size=14, bg_color="FFFF00")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    cell = ws["A1"]
    assert cell.font.bold is True
    assert cell.font.size == 14
    assert cell.fill.start_color.rgb == "00FFFF00"
    wb.close()


def test_set_column_width(sample_xlsx: str) -> None:
    set_column_width(sample_xlsx, "Sheet1", "B", 25.0)
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws.column_dimensions["B"].width == 25.0
    wb.close()


def test_set_row_height(sample_xlsx: str) -> None:
    set_row_height(sample_xlsx, "Sheet1", 1, 30.0)
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws.row_dimensions[1].height == 30.0
    wb.close()


def test_merge_cells(sample_xlsx: str) -> None:
    merge_cells(sample_xlsx, "Sheet1", "A1:D1")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert "A1:D1" in [str(m) for m in ws.merged_cells.ranges]
    wb.close()


def test_unmerge_cells(sample_xlsx: str) -> None:
    merge_cells(sample_xlsx, "Sheet1", "A1:D1")
    unmerge_cells(sample_xlsx, "Sheet1", "A1:D1")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert "A1:D1" not in [str(m) for m in ws.merged_cells.ranges]
    wb.close()


def test_auto_fit_columns(sample_xlsx: str) -> None:
    result = auto_fit_columns(sample_xlsx, "Sheet1")
    assert "Auto-fitted" in result
