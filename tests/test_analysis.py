from __future__ import annotations

from mcp_server.tools.analysis import (
    aggregate_data,
    column_statistics,
    filter_data,
    find_duplicates,
    profile_data,
    search_replace,
    sort_data,
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
