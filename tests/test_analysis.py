from __future__ import annotations

import openpyxl
import pandas as pd
from openpyxl.styles import Font

from mcp_server.tools.analysis import (
    aggregate_data,
    column_statistics,
    export_analysis,
    filter_data,
    filter_data_advanced,
    find_cells_by_format,
    find_duplicates,
    normalize_data,
    profile_data,
    search_replace,
    sort_data,
    transpose_data,
)
from mcp_server.tools.cell_ops import read_cell


def test_filter_data_equals(sample_xlsx: str) -> None:
    result = filter_data(sample_xlsx, "Sheet1", "City", "==", "New York")
    assert result["matched_count"] == 2
    assert all(row[2] == "New York" for row in result["matched_rows"])


def test_filter_data_contains(sample_xlsx: str) -> None:
    result = filter_data(sample_xlsx, "Sheet1", "Name", "contains", "li")
    assert result["matched_count"] == 2  # Alice, Charlie


def test_sort_data(sample_xlsx: str) -> None:
    sort_data(sample_xlsx, "Sheet1", column="Age", ascending=True)
    cell = read_cell(sample_xlsx, "Sheet1", "A2")
    assert cell["value"] == "Bob"  # Bob is youngest at 25


def test_column_statistics(sample_xlsx: str) -> None:
    stats = column_statistics(sample_xlsx, "Sheet1", "Salary")
    assert stats["count"] == 5
    assert stats["min_val"] == 55000.0
    assert stats["max_val"] == 90000.0
    assert stats["sum_val"] == 357000.0


def test_aggregate_data(sample_xlsx: str) -> None:
    result = aggregate_data(sample_xlsx, "Sheet1", "City", "Salary", "sum")
    groups = result["groups"]
    city_sums = {g["City"]: g["Salary_sum"] for g in groups}
    assert city_sums["New York"] == 160000
    assert city_sums["Chicago"] == 117000


def test_find_duplicates(sample_xlsx: str) -> None:
    result = find_duplicates(sample_xlsx, "Sheet1", ["City"])
    # New York (2) + Chicago (2) = 4 duplicate rows
    assert result["count"] == 4


def test_profile_data(sample_xlsx: str) -> None:
    result = profile_data(sample_xlsx, "Sheet1")
    assert result["row_count"] == 5
    assert len(result["columns"]) == 4
    col_names = [c["name"] for c in result["columns"]]
    assert "Name" in col_names
    assert "Salary" in col_names


def test_search_replace(sample_xlsx: str) -> None:
    search_replace(sample_xlsx, "Sheet1", "Alice", "Alicia")
    cell = read_cell(sample_xlsx, "Sheet1", "A2")
    assert cell["value"] == "Alicia"


def test_column_statistics_skewness_kurtosis(sample_xlsx: str) -> None:
    """Test that column_statistics includes skewness and kurtosis."""
    result = column_statistics(sample_xlsx, "Sheet1", "Salary")
    assert "skewness" in result
    assert "kurtosis" in result


def test_vlookup_helper(tmp_path) -> None:
    """Test vlookup helper with exact matching."""
    from openpyxl import Workbook

    from mcp_server.tools.analysis import vlookup_helper

    lookup_path = str(tmp_path / "lookup.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ID", "Name"])
    ws.append([1, "Alice"])
    ws.append([2, "Bob"])
    ws.append([3, "Charlie"])
    wb.save(lookup_path)

    data_path = str(tmp_path / "data.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ID", "Score", "Grade"])
    ws.append([1, 95, "A"])
    ws.append([2, 87, "B"])
    ws.append([3, 92, "A"])
    wb.save(data_path)

    result = vlookup_helper(
        lookup_file=lookup_path,
        data_file=data_path,
        lookup_column="A",
        data_key_column="A",
        data_return_columns=["B", "C"],
    )
    assert result["matched"] == 3
    assert result["unmatched"] == 0
    assert len(result["results"]) == 3


def test_find_cells_by_format(sample_xlsx: str) -> None:
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    ws["A1"].font = Font(bold=True)
    wb.save(sample_xlsx)
    wb.close()

    result = find_cells_by_format(sample_xlsx, "Sheet1", bold=True)
    assert any(r["cell_ref"] == "A1" for r in result)


def test_find_cells_by_format_no_conditions(sample_xlsx: str) -> None:
    import pytest

    with pytest.raises(ValueError, match="At least one"):
        find_cells_by_format(sample_xlsx, "Sheet1")


def test_normalize_data_min_max(sample_xlsx: str) -> None:
    result = normalize_data(sample_xlsx, "Sheet1", columns=["Age"], method="min_max")
    assert "normalized" in result.lower() or "Age" in result
    df = pd.read_excel(sample_xlsx, sheet_name="Sheet1")
    assert df["Age"].between(0, 1).all()


def test_normalize_data_zscore(sample_xlsx: str) -> None:
    result = normalize_data(sample_xlsx, "Sheet1", columns=["Salary"], method="zscore")
    assert "normalized" in result.lower() or "Salary" in result
    df = pd.read_excel(sample_xlsx, sheet_name="Sheet1")
    # z-score of entire column should have mean ~0
    assert abs(df["Salary"].mean()) < 1e-9


def test_normalize_data_output_sheet(sample_xlsx: str) -> None:
    normalize_data(sample_xlsx, "Sheet1", columns=["Age"], method="min_max", output_sheet="Normalized")
    df = pd.read_excel(sample_xlsx, sheet_name="Normalized")
    assert df["Age"].between(0, 1).all()


def test_export_analysis(tmp_path: str) -> None:
    output = tmp_path / "export.xlsx"
    data = [{"Name": "Alice", "Score": 95}, {"Name": "Bob", "Score": 87}]
    result = export_analysis(data, str(output))
    assert str(output) in result or "export.xlsx" in result
    wb = openpyxl.load_workbook(str(output))
    ws = wb.active
    assert ws["A1"].value == "Name"
    assert ws["A2"].value == "Alice"
    wb.close()


def test_transpose_data_correct(tmp_path) -> None:
    from openpyxl import Workbook as WB

    path = str(tmp_path / "transpose.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Score"])
    ws.append(["Alice", 95])
    ws.append(["Bob", 87])
    wb.save(path)

    transpose_data(path, "Sheet1")

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2["Sheet1"]
    # Row 2 col 1 = original column name "Name"; cols 2,3 = data values
    assert ws2.cell(row=2, column=1).value == "Name"
    assert ws2.cell(row=2, column=2).value == "Alice"
    assert ws2.cell(row=2, column=3).value == "Bob"
    assert ws2.cell(row=3, column=1).value == "Score"
    assert ws2.cell(row=3, column=2).value == 95
    wb2.close()


def test_search_replace_regex(tmp_path) -> None:
    from openpyxl import Workbook as WB

    path = str(tmp_path / "regex.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["apple", "orange", "banana"])
    wb.save(path)

    search_replace(path, "Sheet1", "apple|orange", "fruit", use_regex=True)

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2.active
    assert ws2.cell(row=1, column=1).value == "fruit"
    assert ws2.cell(row=1, column=2).value == "fruit"
    assert ws2.cell(row=1, column=3).value == "banana"
    wb2.close()


def test_search_replace_numeric(tmp_path) -> None:
    from openpyxl import Workbook as WB

    path = str(tmp_path / "numeric.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.cell(row=1, column=1, value=42)
    wb.save(path)

    search_replace(path, "Sheet1", "42", "0")

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2.active
    assert ws2.cell(row=1, column=1).value == "0"
    wb2.close()


def test_filter_data_case_sensitive(sample_xlsx: str) -> None:
    # Case-insensitive (default): "new york" matches "New York"
    result_ci = filter_data(sample_xlsx, "Sheet1", "City", "contains", "new york", case_sensitive=False)
    # Case-sensitive: "new york" does NOT match "New York"
    result_cs = filter_data(sample_xlsx, "Sheet1", "City", "contains", "new york", case_sensitive=True)
    assert result_ci["matched_count"] == 2
    assert result_cs["matched_count"] == 0


def test_filter_data_advanced_and(tmp_path) -> None:
    from openpyxl import Workbook as WB

    path = str(tmp_path / "adv.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["City", "Salary"])
    ws.append(["New York", 70000])
    ws.append(["Chicago", 55000])
    ws.append(["New York", 90000])
    ws.append(["Chicago", 62000])
    ws.append(["Boston", 80000])
    wb.save(path)

    result = filter_data_advanced(
        path,
        "Sheet1",
        conditions=[
            {"column": "City", "operator": "==", "value": "New York"},
            {"column": "Salary", "operator": ">", "value": 75000},
        ],
        logic="AND",
    )
    # New York AND Salary > 75000: only New York/90000
    assert result["rows"] == 1


def test_filter_data_advanced_or(tmp_path) -> None:
    from openpyxl import Workbook as WB

    path = str(tmp_path / "adv.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["City", "Salary"])
    ws.append(["New York", 70000])
    ws.append(["Chicago", 55000])
    ws.append(["New York", 90000])
    ws.append(["Chicago", 62000])
    ws.append(["Boston", 80000])
    wb.save(path)

    result = filter_data_advanced(
        path,
        "Sheet1",
        conditions=[
            {"column": "City", "operator": "==", "value": "Boston"},
            {"column": "Salary", "operator": ">", "value": 80000},
        ],
        logic="OR",
    )
    # Boston (1) OR Salary > 80000 (New York/90000): 2 rows
    assert result["rows"] == 2


def test_profile_data_categorical(sample_xlsx: str) -> None:
    result = profile_data(sample_xlsx, "Sheet1")
    assert "categorical_columns" in result
    cat = result["categorical_columns"]
    # Name and City are string/object columns
    assert "Name" in cat
    assert "City" in cat
    assert cat["Name"]["unique_count"] == 5
    assert cat["City"]["top_value"] in ("New York", "Chicago")
    assert cat["City"]["null_count"] == 0
