from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest
from openpyxl import Workbook

from mcp_server.tools.tables import (
    create_table,
    get_table_data,
    list_tables,
    resize_table,
    set_table_totals_row,
)


def _make_table_workbook(tmp_path: Path, name: str = "table.xlsx") -> str:
    """Create a workbook with data suitable for table creation."""
    path = str(tmp_path / name)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Product", "Qty", "Price", "Total"])
    ws.append(["Widget", 10, 5.00, 50.00])
    ws.append(["Gadget", 20, 8.50, 170.00])
    ws.append(["Gizmo", 15, 3.25, 48.75])
    ws.append(["Doohickey", 5, 12.00, 60.00])
    wb.save(path)
    wb.close()
    return path


# ── create_table ──────────────────────────────────────────────────────────


def test_create_table(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    result = create_table(path, "Sheet1", "A1:D5", "ProductTable")
    assert "ProductTable" in result
    assert "A1:D5" in result


def test_create_table_custom_style(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    result = create_table(path, "Sheet1", "A1:D5", "StyledTable", style_name="TableStyleLight1")
    assert "StyledTable" in result
    tables = list_tables(path, "Sheet1")
    assert tables[0]["style"] == "TableStyleLight1"


def test_create_table_verify_openpyxl(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    create_table(path, "Sheet1", "A1:D5", "VerifyTable")
    wb = openpyxl.load_workbook(path)
    ws = wb["Sheet1"]
    assert "VerifyTable" in ws.tables
    wb.close()


def test_create_table_on_sample_data(sample_xlsx: str) -> None:
    result = create_table(sample_xlsx, "Sheet1", "A1:D6", "SalesTable")
    assert "SalesTable" in result


# ── list_tables ───────────────────────────────────────────────────────────


def test_list_tables_empty(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    tables = list_tables(path, "Sheet1")
    assert tables == []


def test_list_tables_single(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    create_table(path, "Sheet1", "A1:D5", "OnlyTable")
    tables = list_tables(path, "Sheet1")
    assert len(tables) == 1
    assert tables[0]["name"] == "OnlyTable"
    assert tables[0]["ref"] == "A1:D5"
    assert tables[0]["style"] == "TableStyleMedium9"


def test_list_tables_multiple(tmp_path: Path) -> None:
    path = str(tmp_path / "multi.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["A", "B", "C", "D", "E", "F"])
    ws.append([1, 2, 3, 4, 5, 6])
    ws.append([7, 8, 9, 10, 11, 12])
    wb.save(path)
    wb.close()
    create_table(path, "Sheet1", "A1:C3", "Table1")
    create_table(path, "Sheet1", "D1:F3", "Table2")
    tables = list_tables(path, "Sheet1")
    assert len(tables) == 2
    names = {t["name"] for t in tables}
    assert names == {"Table1", "Table2"}


# ── resize_table ──────────────────────────────────────────────────────────


def test_resize_table(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    create_table(path, "Sheet1", "A1:D5", "ResizeMe")
    result = resize_table(path, "Sheet1", "ResizeMe", "A1:D3")
    assert "A1:D3" in result
    tables = list_tables(path, "Sheet1")
    match = next(t for t in tables if t["name"] == "ResizeMe")
    assert match["ref"] == "A1:D3"


def test_resize_table_expand(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    create_table(path, "Sheet1", "A1:C5", "ExpandMe")
    result = resize_table(path, "Sheet1", "ExpandMe", "A1:D5")
    assert "A1:D5" in result


def test_resize_table_not_found(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    with pytest.raises(ValueError, match="not found"):
        resize_table(path, "Sheet1", "Nonexistent", "A1:D5")


# ── set_table_totals_row ─────────────────────────────────────────────────


def test_set_table_totals_row_enable(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    create_table(path, "Sheet1", "A1:D5", "TotalsTable")
    result = set_table_totals_row(path, "Sheet1", "TotalsTable", show_totals=True)
    assert "enabled" in result


def test_set_table_totals_row_disable(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    create_table(path, "Sheet1", "A1:D5", "TotalsOff")
    set_table_totals_row(path, "Sheet1", "TotalsOff", show_totals=True)
    result = set_table_totals_row(path, "Sheet1", "TotalsOff", show_totals=False)
    assert "disabled" in result


def test_set_table_totals_row_with_functions(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    create_table(path, "Sheet1", "A1:D5", "FuncTable")
    result = set_table_totals_row(
        path,
        "Sheet1",
        "FuncTable",
        show_totals=True,
        column_totals={"Qty": "sum", "Price": "average", "Total": "sum"},
    )
    assert "enabled" in result


def test_set_table_totals_row_not_found(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    with pytest.raises(ValueError, match="not found"):
        set_table_totals_row(path, "Sheet1", "Ghost", show_totals=True)


# ── get_table_data ────────────────────────────────────────────────────────


def test_get_table_data(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    create_table(path, "Sheet1", "A1:D5", "DataTable")
    data = get_table_data(path, "Sheet1", "DataTable")
    assert data["table_name"] == "DataTable"
    assert data["ref"] == "A1:D5"
    assert data["headers"] == ["Product", "Qty", "Price", "Total"]
    assert data["row_count"] == 4
    assert data["rows"][0] == ["Widget", 10, 5.00, 50.00]


def test_get_table_data_all_rows(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    create_table(path, "Sheet1", "A1:D5", "AllRows")
    data = get_table_data(path, "Sheet1", "AllRows")
    assert len(data["rows"]) == 4
    last_row = data["rows"][-1]
    assert last_row[0] == "Doohickey"


def test_get_table_data_not_found(tmp_path: Path) -> None:
    path = _make_table_workbook(tmp_path)
    with pytest.raises(ValueError, match="not found"):
        get_table_data(path, "Sheet1", "Missing")


# ── edge cases ────────────────────────────────────────────────────────────


def test_table_lifecycle(tmp_path: Path) -> None:
    """Create → list → resize → totals → data: full lifecycle."""
    path = _make_table_workbook(tmp_path)
    create_table(path, "Sheet1", "A1:D5", "Lifecycle")
    tables = list_tables(path, "Sheet1")
    assert len(tables) == 1

    resize_table(path, "Sheet1", "Lifecycle", "A1:D4")
    set_table_totals_row(path, "Sheet1", "Lifecycle", show_totals=True)

    data = get_table_data(path, "Sheet1", "Lifecycle")
    assert data["table_name"] == "Lifecycle"


def test_create_table_on_empty_sheet(empty_xlsx: str) -> None:
    """Table creation on sheet with no data (headers only in range)."""
    wb = openpyxl.load_workbook(empty_xlsx)
    ws = wb["Sheet1"]
    ws["A1"] = "Col1"
    ws["B1"] = "Col2"
    ws["A2"] = 1
    ws["B2"] = 2
    wb.save(empty_xlsx)
    wb.close()

    result = create_table(empty_xlsx, "Sheet1", "A1:B2", "EmptyTable")
    assert "EmptyTable" in result
