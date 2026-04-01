"""Tests for pivot table persistence and refresh feature."""

from __future__ import annotations

import json
from pathlib import Path

import openpyxl
import pytest

from mcp_server.tools.pivot_etl import _PIVOTS_SHEET, create_pivot_table, refresh_pivot_table


def make_source_workbook(tmp_path: Path, file_name: str = "source.xlsx") -> str:
    """Create a workbook with simple data for pivot testing."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Region", "Product", "Sales"])
    ws.append(["North", "Apple", 100])
    ws.append(["North", "Banana", 200])
    ws.append(["South", "Apple", 150])
    ws.append(["South", "Banana", 300])
    path = str(tmp_path / file_name)
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# create_pivot_table — pivot persistence
# ---------------------------------------------------------------------------


def test_create_pivot_persists_definition(tmp_path: Path) -> None:
    path = make_source_workbook(tmp_path)
    create_pivot_table(
        path,
        "Data",
        index_cols=["Region"],
        value_cols=["Sales"],
        aggfunc="sum",
        output_sheet="Pivot1",
    )

    wb = openpyxl.load_workbook(path)
    assert _PIVOTS_SHEET in wb.sheetnames
    raw = wb[_PIVOTS_SHEET]["A1"].value
    assert raw is not None
    pivots = json.loads(str(raw))
    assert "Pivot1" in pivots
    wb.close()


def test_create_pivot_persists_all_params(tmp_path: Path) -> None:
    path = make_source_workbook(tmp_path)
    create_pivot_table(
        path,
        "Data",
        index_cols=["Region"],
        value_cols=["Sales"],
        aggfunc="mean",
        output_sheet="Summary",
    )

    wb = openpyxl.load_workbook(path)
    pivots = json.loads(str(wb[_PIVOTS_SHEET]["A1"].value))
    defn = pivots["Summary"]
    wb.close()

    assert defn["file_path"] == path
    assert defn["sheet_name"] == "Data"
    assert defn["index_cols"] == ["Region"]
    assert defn["value_cols"] == ["Sales"]
    assert defn["aggfunc"] == "mean"
    assert defn["output_sheet"] == "Summary"


def test_mcp_pivots_sheet_is_hidden(tmp_path: Path) -> None:
    path = make_source_workbook(tmp_path)
    create_pivot_table(
        path,
        "Data",
        index_cols=["Region"],
        value_cols=["Sales"],
        output_sheet="PivotHidden",
    )

    wb = openpyxl.load_workbook(path)
    ws = wb[_PIVOTS_SHEET]
    assert ws.sheet_state == "hidden"
    wb.close()


def test_multiple_pivots_coexist(tmp_path: Path) -> None:
    path = make_source_workbook(tmp_path)
    create_pivot_table(
        path,
        "Data",
        index_cols=["Region"],
        value_cols=["Sales"],
        output_sheet="ByRegion",
    )
    create_pivot_table(
        path,
        "Data",
        index_cols=["Product"],
        value_cols=["Sales"],
        output_sheet="ByProduct",
    )

    wb = openpyxl.load_workbook(path)
    pivots = json.loads(str(wb[_PIVOTS_SHEET]["A1"].value))
    wb.close()

    assert "ByRegion" in pivots
    assert "ByProduct" in pivots
    assert pivots["ByRegion"]["index_cols"] == ["Region"]
    assert pivots["ByProduct"]["index_cols"] == ["Product"]


# ---------------------------------------------------------------------------
# refresh_pivot_table
# ---------------------------------------------------------------------------


def test_refresh_pivot_basic(tmp_path: Path) -> None:
    path = make_source_workbook(tmp_path)
    create_pivot_table(
        path,
        "Data",
        index_cols=["Region"],
        value_cols=["Sales"],
        aggfunc="sum",
        output_sheet="PivotOut",
    )

    # Add a new row to the source data
    wb = openpyxl.load_workbook(path)
    wb["Data"].append(["North", "Cherry", 500])
    wb.save(path)
    wb.close()

    refresh_pivot_table(path, "PivotOut")

    wb2 = openpyxl.load_workbook(path)
    ws = wb2["PivotOut"]
    # Collect all North values from the pivot sheet
    north_sales: list[float] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == "North":
            north_sales.append(row[1])
    wb2.close()

    # North total is now 100 + 200 + 500 = 800
    assert north_sales, "North row missing from refreshed pivot"
    assert north_sales[0] == pytest.approx(800)


def test_refresh_pivot_return_dict(tmp_path: Path) -> None:
    path = make_source_workbook(tmp_path)
    create_pivot_table(
        path,
        "Data",
        index_cols=["Region"],
        value_cols=["Sales"],
        output_sheet="PivotReturn",
    )

    result = refresh_pivot_table(path, "PivotReturn")

    assert "refreshed" in result
    assert "source_file" in result
    assert "source_sheet" in result
    assert result["refreshed"] == "PivotReturn"


def test_refresh_pivot_not_found_raises(tmp_path: Path) -> None:
    path = make_source_workbook(tmp_path)
    # No pivot definitions created — just an empty workbook
    with pytest.raises(ValueError, match="No pivot definition found"):
        refresh_pivot_table(path, "NonExistentPivot")


def test_refresh_pivot_override_source_sheet(tmp_path: Path) -> None:
    path = make_source_workbook(tmp_path)
    create_pivot_table(
        path,
        "Data",
        index_cols=["Region"],
        value_cols=["Sales"],
        output_sheet="PivotOverride",
    )

    # Add a second sheet with different data
    wb = openpyxl.load_workbook(path)
    ws2 = wb.create_sheet("Data2")
    ws2.append(["Region", "Product", "Sales"])
    ws2.append(["East", "Mango", 999])
    wb.save(path)
    wb.close()

    refresh_pivot_table(path, "PivotOverride", source_sheet="Data2")

    wb2 = openpyxl.load_workbook(path)
    ws_pivot = wb2["PivotOverride"]
    regions = [ws_pivot.cell(row=r, column=1).value for r in range(2, ws_pivot.max_row + 1)]
    wb2.close()

    assert "East" in regions
    # Old regions should no longer be present in pure Data2-based pivot
    assert "North" not in regions
    assert "South" not in regions


def test_refresh_pivot_updates_stored_definition(tmp_path: Path) -> None:
    path = make_source_workbook(tmp_path)
    create_pivot_table(
        path,
        "Data",
        index_cols=["Region"],
        value_cols=["Sales"],
        output_sheet="PivotUpdate",
    )

    # Add Data2 sheet
    wb = openpyxl.load_workbook(path)
    ws2 = wb.create_sheet("Data2")
    ws2.append(["Region", "Product", "Sales"])
    ws2.append(["West", "Grape", 77])
    wb.save(path)
    wb.close()

    refresh_pivot_table(path, "PivotUpdate", source_sheet="Data2")

    # The stored definition should now reference Data2
    wb2 = openpyxl.load_workbook(path)
    pivots = json.loads(str(wb2[_PIVOTS_SHEET]["A1"].value))
    wb2.close()

    assert pivots["PivotUpdate"]["sheet_name"] == "Data2"
