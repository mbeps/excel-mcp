from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from mcp_server.tools.pivot_etl import (
    add_computed_column,
    append_datasets,
    create_pivot_table,
    deduplicate_data,
    find_differences,
    merge_datasets,
    unpivot_data,
)


def _make_two_sheet_workbook(tmp_path: Path) -> str:
    path = str(tmp_path / "multi.xlsx")
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Orders"
    ws1.append(["OrderID", "Product", "Qty"])
    ws1.append([1, "Widget", 10])
    ws1.append([2, "Gadget", 5])
    ws1.append([3, "Widget", 7])

    ws2 = wb.create_sheet("Products")
    ws2.append(["Product", "Price"])
    ws2.append(["Widget", 25.0])
    ws2.append(["Gadget", 40.0])
    wb.save(path)
    return path


def test_create_pivot_table(sample_xlsx: str, tmp_path: Path) -> None:
    out = str(tmp_path / "pivot_out.xlsx")
    result = create_pivot_table(
        sample_xlsx,
        "Sheet1",
        index_cols=["City"],
        value_cols=["Salary"],
        aggfunc="sum",
        output_file=out,
    )
    assert "data" in result
    city_sums = {r["City"]: r["Salary"] for r in result["data"]}
    assert city_sums["New York"] == 160000


def test_unpivot_data(sample_xlsx: str) -> None:
    result = unpivot_data(
        sample_xlsx,
        "Sheet1",
        id_vars=["Name"],
        value_vars=["Age", "Salary"],
    )
    assert result["row_count"] == 10  # 5 rows × 2 value vars


def test_merge_datasets(tmp_path: Path) -> None:
    path = _make_two_sheet_workbook(tmp_path)
    result = merge_datasets(path, "Orders", "Products", join_key="Product", how="left")
    assert result["row_count"] == 3
    assert all("Price" in r for r in result["data"])


def test_add_computed_column(tmp_path: Path) -> None:
    path = _make_two_sheet_workbook(tmp_path)
    add_computed_column(path, "Orders", "DoubleQty", "Qty * 2")
    from mcp_server.tools.cell_ops import read_cell

    cell = read_cell(path, "Orders", "D1")
    assert cell["value"] == "DoubleQty"
    cell_val = read_cell(path, "Orders", "D2")
    assert cell_val["value"] == 20  # 10 * 2


def test_deduplicate_data(tmp_path: Path) -> None:
    path = str(tmp_path / "dupes.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Value"])
    ws.append(["A", 1])
    ws.append(["B", 2])
    ws.append(["A", 1])
    ws.append(["C", 3])
    wb.save(path)

    result = deduplicate_data(path, "Sheet1")
    assert "Removed 1" in result


def test_append_datasets(tmp_path: Path) -> None:
    path = str(tmp_path / "append.xlsx")
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Sheet1"
    ws1.append(["Name", "Value"])
    ws1.append(["A", 1])
    ws1.append(["B", 2])
    ws2 = wb.create_sheet("Sheet2")
    ws2.append(["Name", "Value"])
    ws2.append(["C", 3])
    ws2.append(["D", 4])
    wb.save(path)

    result = append_datasets(path, "Sheet1", "Sheet2", "Output")
    assert result["rows"] == 4
    assert result["output_sheet"] == "Output"


def test_find_differences(tmp_path: Path) -> None:
    path = str(tmp_path / "diff.xlsx")
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "SheetA"
    ws1.append(["ID", "Name"])
    ws1.append([1, "Alice"])
    ws1.append([2, "Bob"])
    ws1.append([3, "Charlie"])
    ws2 = wb.create_sheet("SheetB")
    ws2.append(["ID", "Name"])
    ws2.append([1, "Alice"])
    ws2.append([2, "Bob"])
    wb.save(path)

    result = find_differences(path, "SheetA", "SheetB", key_columns=["ID"], output_sheet="Diff")
    assert result["rows_only_in_a"] == 1
    assert result["output_sheet"] == "Diff"


def test_create_pivot_table_margins(sample_xlsx: str, tmp_path: Path) -> None:
    out = str(tmp_path / "pivot_margins.xlsx")
    result = create_pivot_table(
        sample_xlsx,
        "Sheet1",
        index_cols=["City"],
        value_cols=["Salary"],
        aggfunc="sum",
        output_file=out,
        include_margins=True,
    )
    cities = [r["City"] for r in result["data"]]
    assert "Total" in cities
