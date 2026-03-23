from __future__ import annotations

from mcp_server.tools.workbook import (
    copy_sheet,
    create_workbook,
    delete_sheet,
    get_sheet_summary,
    get_workbook_metadata,
    list_sheets,
    rename_sheet,
)


def test_list_sheets(sample_xlsx: str) -> None:
    sheets = list_sheets(sample_xlsx)
    assert len(sheets) == 1
    assert sheets[0]["name"] == "Sheet1"


def test_create_workbook(tmp_path) -> None:
    path = str(tmp_path / "new.xlsx")
    result = create_workbook(path, sheet_names=["Sales", "Inventory"])
    assert "Sales" in result["sheets"]
    assert "Inventory" in result["sheets"]
    assert "Sheet" not in result["sheets"]


def test_get_workbook_metadata(sample_xlsx: str) -> None:
    meta = get_workbook_metadata(sample_xlsx)
    assert len(meta["sheets"]) == 1
    assert meta["sheets"][0]["name"] == "Sheet1"
    assert meta["active_sheet"] == "Sheet1"
    assert isinstance(meta["named_ranges"], list)


def test_get_sheet_summary(sample_xlsx: str) -> None:
    summary = get_sheet_summary(sample_xlsx, "Sheet1")
    assert summary["name"] == "Sheet1"
    assert summary["row_count"] == 6  # 1 header + 5 data
    assert summary["col_count"] == 4
    assert "Name" in summary["headers"]


def test_rename_sheet(sample_xlsx: str) -> None:
    rename_sheet(sample_xlsx, "Sheet1", "Data")
    sheets = list_sheets(sample_xlsx)
    assert sheets[0]["name"] == "Data"


def test_delete_sheet(sample_xlsx: str) -> None:
    copy_sheet(sample_xlsx, "Sheet1", "Sheet2")
    delete_sheet(sample_xlsx, "Sheet2")
    sheets = list_sheets(sample_xlsx)
    assert len(sheets) == 1
    assert sheets[0]["name"] == "Sheet1"


def test_copy_sheet(sample_xlsx: str) -> None:
    copy_sheet(sample_xlsx, "Sheet1", "Sheet1_Copy")
    sheets = list_sheets(sample_xlsx)
    names = [s["name"] for s in sheets]
    assert "Sheet1_Copy" in names
    assert len(sheets) == 2


def test_write_multi_sheet(tmp_path) -> None:
    """Test creating a workbook with multiple sheets, headers, and data."""
    from mcp_server.tools.workbook import write_multi_sheet

    file_path = str(tmp_path / "multi.xlsx")
    sheets = [
        {"name": "Sales", "headers": ["Product", "Revenue"], "data": [["Widget", 100], ["Gadget", 200]]},
        {"name": "Costs", "headers": ["Item", "Amount"], "data": [["Rent", 500], ["Utils", 100]]},
    ]
    result = write_multi_sheet(file_path, sheets)
    assert "sheets_created" in result
    assert len(result["sheets_created"]) == 2
    from openpyxl import load_workbook

    wb = load_workbook(file_path)
    assert "Sales" in wb.sheetnames
    assert "Costs" in wb.sheetnames
    wb.close()
