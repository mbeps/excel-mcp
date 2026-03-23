from __future__ import annotations

import openpyxl

from mcp_server.tools.worksheet_ops import (
    freeze_panes,
    group_columns,
    group_rows,
    hide_columns,
    hide_rows,
    hide_sheet,
    move_sheet,
    remove_auto_filter,
    set_auto_filter,
    set_header_footer,
    set_page_margins,
    set_page_setup,
    set_print_area,
    set_sheet_tab_color,
    set_zoom,
    show_gridlines,
    unfreeze_panes,
    ungroup_columns,
    ungroup_rows,
    unhide_columns,
    unhide_rows,
    unhide_sheet,
)


def test_freeze_panes(sample_xlsx: str) -> None:
    freeze_panes(sample_xlsx, "Sheet1", "B2")
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"].freeze_panes == "B2"


def test_unfreeze_panes(sample_xlsx: str) -> None:
    freeze_panes(sample_xlsx, "Sheet1", "B2")
    unfreeze_panes(sample_xlsx, "Sheet1")
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"].freeze_panes is None


def test_set_auto_filter(sample_xlsx: str) -> None:
    set_auto_filter(sample_xlsx, "Sheet1", "A1:D6")
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"].auto_filter.ref == "A1:D6"


def test_remove_auto_filter(sample_xlsx: str) -> None:
    set_auto_filter(sample_xlsx, "Sheet1", "A1:D6")
    remove_auto_filter(sample_xlsx, "Sheet1")
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"].auto_filter.ref is None


def test_hide_rows(sample_xlsx: str) -> None:
    hide_rows(sample_xlsx, "Sheet1", 2, 3)
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws.row_dimensions[2].hidden is True
    assert ws.row_dimensions[3].hidden is True


def test_unhide_rows(sample_xlsx: str) -> None:
    hide_rows(sample_xlsx, "Sheet1", 2, 3)
    unhide_rows(sample_xlsx, "Sheet1", 2, 3)
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws.row_dimensions[2].hidden is False
    assert ws.row_dimensions[3].hidden is False


def test_hide_columns(sample_xlsx: str) -> None:
    hide_columns(sample_xlsx, "Sheet1", "B", "C")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws.column_dimensions["B"].hidden is True
    assert ws.column_dimensions["C"].hidden is True


def test_unhide_columns(sample_xlsx: str) -> None:
    hide_columns(sample_xlsx, "Sheet1", "B", "C")
    unhide_columns(sample_xlsx, "Sheet1", "B", "C")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws.column_dimensions["B"].hidden is False
    assert ws.column_dimensions["C"].hidden is False


def test_group_rows(sample_xlsx: str) -> None:
    group_rows(sample_xlsx, "Sheet1", 2, 5)
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    for r in range(2, 6):
        assert ws.row_dimensions[r].outlineLevel == 1


def test_ungroup_rows(sample_xlsx: str) -> None:
    group_rows(sample_xlsx, "Sheet1", 2, 5)
    ungroup_rows(sample_xlsx, "Sheet1", 2, 5)
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    for r in range(2, 6):
        assert ws.row_dimensions[r].outlineLevel == 0


def test_group_columns(sample_xlsx: str) -> None:
    group_columns(sample_xlsx, "Sheet1", "B", "D")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    for col in ("B", "C", "D"):
        assert ws.column_dimensions[col].outlineLevel == 1


def test_ungroup_columns(sample_xlsx: str) -> None:
    group_columns(sample_xlsx, "Sheet1", "B", "D")
    ungroup_columns(sample_xlsx, "Sheet1", "B", "D")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    for col in ("B", "C", "D"):
        assert ws.column_dimensions[col].outlineLevel == 0


def test_set_sheet_tab_color(sample_xlsx: str) -> None:
    set_sheet_tab_color(sample_xlsx, "Sheet1", "FF0000")
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"].sheet_properties.tabColor.rgb == "00FF0000"


def test_hide_sheet(tmp_path) -> None:
    path = str(tmp_path / "multi.xlsx")
    wb = openpyxl.Workbook()
    wb.active.title = "Sheet1"
    wb.create_sheet("Sheet2")
    wb.save(path)

    hide_sheet(path, "Sheet2")
    wb = openpyxl.load_workbook(path)
    assert wb["Sheet2"].sheet_state == "hidden"


def test_unhide_sheet(tmp_path) -> None:
    path = str(tmp_path / "multi.xlsx")
    wb = openpyxl.Workbook()
    wb.active.title = "Sheet1"
    wb.create_sheet("Sheet2")
    wb.save(path)

    hide_sheet(path, "Sheet2")
    unhide_sheet(path, "Sheet2")
    wb = openpyxl.load_workbook(path)
    assert wb["Sheet2"].sheet_state == "visible"


def test_move_sheet(tmp_path) -> None:
    path = str(tmp_path / "multi.xlsx")
    wb = openpyxl.Workbook()
    wb.active.title = "Sheet1"
    wb.create_sheet("Sheet2")
    wb.create_sheet("Sheet3")
    wb.save(path)

    move_sheet(path, "Sheet1", 2)
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == ["Sheet2", "Sheet3", "Sheet1"]


def test_set_zoom(sample_xlsx: str) -> None:
    set_zoom(sample_xlsx, "Sheet1", 150)
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"].sheet_view.zoomScale == 150


def test_show_gridlines(sample_xlsx: str) -> None:
    show_gridlines(sample_xlsx, "Sheet1", show=False)
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"].sheet_view.showGridLines is False


def test_set_print_area(sample_xlsx: str) -> None:
    set_print_area(sample_xlsx, "Sheet1", "A1:D10")
    wb = openpyxl.load_workbook(sample_xlsx)
    assert "$A$1" in wb["Sheet1"].print_area
    assert "$D$10" in wb["Sheet1"].print_area


def test_set_page_setup(sample_xlsx: str) -> None:
    set_page_setup(sample_xlsx, "Sheet1", orientation="landscape")
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"].page_setup.orientation == "landscape"


def test_set_page_margins(sample_xlsx: str) -> None:
    set_page_margins(sample_xlsx, "Sheet1", top=1.0, bottom=1.0, left=0.5, right=0.5)
    wb = openpyxl.load_workbook(sample_xlsx)
    margins = wb["Sheet1"].page_margins
    assert float(margins.top) == 1.0
    assert float(margins.bottom) == 1.0
    assert float(margins.left) == 0.5
    assert float(margins.right) == 0.5


def test_set_header_footer(sample_xlsx: str) -> None:
    set_header_footer(sample_xlsx, "Sheet1", header_center="My Report")
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"].oddHeader.center.text == "My Report"
