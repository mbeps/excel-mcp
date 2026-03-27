from __future__ import annotations

from mcp_server.tools.analysis import (
    aggregate_data,
    column_statistics,
    filter_data_advanced,
    find_duplicates,
    sort_data,
)
from mcp_server.tools.cell_ops import read_cell


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


# ============================================================
# Additional comprehensive tests
# ============================================================

import pytest


def test_sort_data_descending(sample_xlsx: str) -> None:
    """Sort by Salary descending — Charlie (90000) should be first data row."""
    sort_data(sample_xlsx, "Sheet1", column="Salary", ascending=False)
    cell = read_cell(sample_xlsx, "Sheet1", "A2")
    assert cell["value"] == "Charlie"  # highest salary = 90000


def test_sort_data_multi_column(sample_xlsx: str) -> None:
    """Sort by City ASC then Age DESC; checks two-level ordering."""
    sort_data(
        sample_xlsx,
        "Sheet1",
        sort_by=[
            {"column": "City", "ascending": True},
            {"column": "Age", "ascending": False},
        ],
    )
    # Sorted: Boston(Eve,32), Chicago(Diana,28), Chicago(Bob,25), New York(Charlie,35), New York(Alice,30)
    assert read_cell(sample_xlsx, "Sheet1", "A2")["value"] == "Eve"  # Boston
    assert read_cell(sample_xlsx, "Sheet1", "A3")["value"] == "Diana"  # Chicago, age 28 > 25


def test_sort_data_with_missing_values(tmp_path) -> None:
    """NaN values in sort column are pushed to the end when sorting ascending."""
    from openpyxl import Workbook as WB

    path = str(tmp_path / "missing.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Score"])
    ws.append(["Alice", 90])
    ws.append(["Bob", None])  # missing value — pandas pushes NaN to end
    ws.append(["Charlie", 70])
    wb.save(path)

    sort_data(path, "Sheet1", column="Score", ascending=True)
    # Charlie(70), Alice(90), Bob(NaN)
    assert read_cell(path, "Sheet1", "A2")["value"] == "Charlie"
    assert read_cell(path, "Sheet1", "A4")["value"] == "Bob"


def test_column_statistics_mean_median_std(sample_xlsx: str) -> None:
    """Verify computed mean, median, and std for the Salary column."""
    # Salaries: 70000, 55000, 90000, 62000, 80000
    stats = column_statistics(sample_xlsx, "Sheet1", "Salary")
    assert stats["mean"] == pytest.approx(71400.0)
    assert stats["median"] == pytest.approx(70000.0)
    assert stats["std"] > 0


def test_column_statistics_text_column(sample_xlsx: str) -> None:
    """Non-numeric column returns a descriptive message, not statistics."""
    result = column_statistics(sample_xlsx, "Sheet1", "Name")
    assert "message" in result
    assert "Cannot compute statistics" in result["message"]


def test_aggregate_data_mean(sample_xlsx: str) -> None:
    """Aggregate mean of Salary grouped by City."""
    result = aggregate_data(sample_xlsx, "Sheet1", "City", "Salary", "mean")
    groups = {g["City"]: g["Salary_mean"] for g in result["groups"]}
    assert groups["New York"] == pytest.approx(80000.0)  # (70000+90000)/2
    assert groups["Chicago"] == pytest.approx(58500.0)  # (55000+62000)/2
    assert groups["Boston"] == pytest.approx(80000.0)


def test_aggregate_data_count(sample_xlsx: str) -> None:
    """Aggregate count of Name grouped by City."""
    result = aggregate_data(sample_xlsx, "Sheet1", "City", "Name", "count")
    groups = {g["City"]: g["Name_count"] for g in result["groups"]}
    assert groups["New York"] == 2
    assert groups["Chicago"] == 2
    assert groups["Boston"] == 1


def test_aggregate_data_max(sample_xlsx: str) -> None:
    """Aggregate max of Salary grouped by City."""
    result = aggregate_data(sample_xlsx, "Sheet1", "City", "Salary", "max")
    groups = {g["City"]: g["Salary_max"] for g in result["groups"]}
    assert groups["New York"] == 90000  # max(70000, 90000)
    assert groups["Chicago"] == 62000  # max(55000, 62000)
    assert groups["Boston"] == 80000


def test_find_duplicates_no_duplicates(sample_xlsx: str) -> None:
    """All names are unique — find_duplicates returns count=0."""
    result = find_duplicates(sample_xlsx, "Sheet1", ["Name"])
    assert result["count"] == 0
    assert result["duplicates"] == []


def test_find_duplicates_multi_column(tmp_path) -> None:
    """Duplicate detection on a multi-column subset (City + Department)."""
    from openpyxl import Workbook as WB

    path = str(tmp_path / "dup_multi.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["City", "Dept", "Score"])
    ws.append(["NY", "Sales", 100])
    ws.append(["NY", "Sales", 200])  # City+Dept duplicated → both rows flagged
    ws.append(["NY", "HR", 100])
    ws.append(["LA", "Sales", 100])
    wb.save(path)

    result = find_duplicates(path, "Sheet1", ["City", "Dept"])
    assert result["count"] == 2  # NY/Sales appears twice


def test_vlookup_helper_no_match(tmp_path) -> None:
    """vlookup_helper returns unmatched entry when lookup key is absent in data."""
    from openpyxl import Workbook as WB
    from mcp_server.tools.analysis import vlookup_helper

    lookup_path = str(tmp_path / "lookup_nm.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ID"])
    ws.append([999])  # no corresponding row in data
    wb.save(lookup_path)

    data_path = str(tmp_path / "data_nm.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ID", "Score"])
    ws.append([1, 100])
    ws.append([2, 200])
    wb.save(data_path)

    result = vlookup_helper(
        lookup_file=lookup_path,
        data_file=data_path,
        lookup_column="A",
        data_key_column="A",
        data_return_columns=["B"],
    )
    assert result["matched"] == 0
    assert result["unmatched"] == 1
    assert result["results"][0]["matched_value"] is None


def test_vlookup_helper_partial_match(tmp_path) -> None:
    """vlookup_helper reports correct matched/unmatched counts for mixed results."""
    from openpyxl import Workbook as WB
    from mcp_server.tools.analysis import vlookup_helper

    lookup_path = str(tmp_path / "lookup_part.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ID"])
    ws.append([1])  # match
    ws.append([999])  # no match
    ws.append([2])  # match
    wb.save(lookup_path)

    data_path = str(tmp_path / "data_part.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ID", "Product"])
    ws.append([1, "Widget"])
    ws.append([2, "Gadget"])
    wb.save(data_path)

    result = vlookup_helper(
        lookup_file=lookup_path,
        data_file=data_path,
        lookup_column="A",
        data_key_column="A",
        data_return_columns=["B"],
    )
    assert result["matched"] == 2
    assert result["unmatched"] == 1


def test_filter_data_advanced_three_conditions_and(tmp_path) -> None:
    """filter_data_advanced with three AND conditions narrows to one row."""
    from openpyxl import Workbook as WB

    path = str(tmp_path / "three_cond.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["City", "Salary", "Dept"])
    ws.append(["New York", 70000, "Sales"])
    ws.append(["New York", 90000, "Engineering"])
    ws.append(["Chicago", 55000, "Sales"])
    ws.append(["Chicago", 62000, "HR"])
    ws.append(["Boston", 80000, "Sales"])
    wb.save(path)

    result = filter_data_advanced(
        path,
        "Sheet1",
        conditions=[
            {"column": "City", "operator": "==", "value": "New York"},
            {"column": "Salary", "operator": ">", "value": 60000},
            {"column": "Dept", "operator": "==", "value": "Engineering"},
        ],
        logic="AND",
    )
    assert result["rows"] == 1  # only New York, Salary>60000, Dept=Engineering


def test_filter_data_advanced_three_conditions_or(tmp_path) -> None:
    """filter_data_advanced with three OR conditions widens to all rows here."""
    from openpyxl import Workbook as WB

    path = str(tmp_path / "or_cond.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["City", "Salary"])
    ws.append(["Alpha", 10])
    ws.append(["Beta", 20])
    ws.append(["Gamma", 30])
    wb.save(path)

    result = filter_data_advanced(
        path,
        "Sheet1",
        conditions=[
            {"column": "City", "operator": "==", "value": "Alpha"},
            {"column": "City", "operator": "==", "value": "Beta"},
            {"column": "Salary", "operator": "==", "value": 30},
        ],
        logic="OR",
    )
    # Alpha OR Beta OR Salary==30 (Gamma) — all 3 rows match
    assert result["rows"] == 3


def test_filter_data_advanced_single_row(tmp_path) -> None:
    """filter_data_advanced works correctly on a sheet with a single data row."""
    from openpyxl import Workbook as WB

    path = str(tmp_path / "single.xlsx")
    wb = WB()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Age"])
    ws.append(["Alice", 30])
    wb.save(path)

    res_match = filter_data_advanced(path, "Sheet1", conditions=[{"column": "Age", "operator": "==", "value": 30}])
    assert res_match["rows"] == 1

    res_no_match = filter_data_advanced(path, "Sheet1", conditions=[{"column": "Age", "operator": "==", "value": 99}])
    assert res_no_match["rows"] == 0
