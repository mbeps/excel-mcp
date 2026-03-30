from __future__ import annotations
import pytest
from openpyxl import Workbook
from mcp_server.tools.analysis import (
    aggregate_data,
    column_statistics,
    filter_data_advanced,
    find_duplicates,
    sort_data,
    vlookup_helper,
)


def test_filter_data_advanced_errors(sample_xlsx: str):
    """Test error handling in filter_data_advanced."""
    with pytest.raises(ValueError, match="Logic must be 'AND' or 'OR'"):
        filter_data_advanced(
            sample_xlsx, "Sheet1", conditions=[{"column": "Age", "operator": "==", "value": 30}], logic="XOR"
        )

    with pytest.raises(ValueError, match="At least one condition is required"):
        filter_data_advanced(sample_xlsx, "Sheet1", conditions=[], logic="AND")

    with pytest.raises(ValueError, match="Unsupported operator"):
        filter_data_advanced(sample_xlsx, "Sheet1", conditions=[{"column": "Age", "operator": "INVALID", "value": 30}])

    with pytest.raises(ValueError, match="Column 'NON_EXISTENT' not found"):
        filter_data_advanced(
            sample_xlsx, "Sheet1", conditions=[{"column": "NON_EXISTENT", "operator": "==", "value": 30}]
        )


def test_filter_data_advanced_operators(sample_xlsx: str):
    """Test all operators in filter_data_advanced."""
    # !=
    res = filter_data_advanced(
        sample_xlsx, "Sheet1", conditions=[{"column": "City", "operator": "!=", "value": "New York"}]
    )
    assert (
        res["rows"] == 3
    )  # Alice (New York), Bob (Chicago), Charlie (New York), David (Chicago), Eve (Boston) -> 3 not New York

    # <=
    res = filter_data_advanced(sample_xlsx, "Sheet1", conditions=[{"column": "Age", "operator": "<=", "value": 30}])
    assert res["rows"] == 3  # 30, 25, 35, 28, 40 -> 30, 25, 28 are <= 30

    # contains
    res = filter_data_advanced(
        sample_xlsx, "Sheet1", conditions=[{"column": "City", "operator": "contains", "value": "York"}]
    )
    assert res["rows"] == 2

    # startswith
    res = filter_data_advanced(
        sample_xlsx, "Sheet1", conditions=[{"column": "City", "operator": "startswith", "value": "Chi"}]
    )
    assert res["rows"] == 2

    # endswith
    res = filter_data_advanced(
        sample_xlsx, "Sheet1", conditions=[{"column": "City", "operator": "endswith", "value": "cago"}]
    )
    assert res["rows"] == 2


def test_filter_data_advanced_output_sheet(sample_xlsx: str):
    """Test filtering with output_sheet."""
    res = filter_data_advanced(
        sample_xlsx, "Sheet1", conditions=[{"column": "Age", "operator": ">", "value": 30}], output_sheet="Filtered"
    )
    assert res["rows"] == 2
    from openpyxl import load_workbook

    wb = load_workbook(sample_xlsx)
    assert "Filtered" in wb.sheetnames
    ws = wb["Filtered"]
    assert ws["A1"].value == "Name"
    assert ws["B2"].value is not None
    wb.close()


def test_sort_data_extended(sample_xlsx: str):
    """Test sort_data with sort_by and errors."""
    # sort_by
    sort_data(
        sample_xlsx, "Sheet1", sort_by=[{"column": "City", "ascending": True}, {"column": "Age", "ascending": False}]
    )
    from mcp_server.tools.cell_ops import read_cell

    # Cities: Boston, Chicago, Chicago, New York, New York
    # Ages for Chicago: 25, 28 -> DESC: 28, 25
    assert read_cell(sample_xlsx, "Sheet1", "A2")["value"] == "Eve"  # Boston
    assert read_cell(sample_xlsx, "Sheet1", "A3")["value"] == "Diana"  # Chicago, 28

    # Errors
    with pytest.raises(ValueError, match="Provide either 'sort_by' or 'column'"):
        sort_data(sample_xlsx, "Sheet1")

    with pytest.raises(ValueError, match="Column 'INVALID' not found"):
        sort_data(sample_xlsx, "Sheet1", column="INVALID")


def test_column_statistics_non_numeric(sample_xlsx: str):
    """Test column_statistics with a non-numeric column."""
    res = column_statistics(sample_xlsx, "Sheet1", "City")
    assert "Cannot compute statistics" in res.message


def test_aggregate_data_extended(sample_xlsx: str):
    """Test aggregate_data with dict aggfunc and errors."""
    # dict aggfunc
    res = aggregate_data(
        sample_xlsx, "Sheet1", group_by="City", value_column="Salary", aggfunc={"Salary": "sum", "Age": "mean"}
    )
    assert "Salary" in res["groups"][0]
    assert "Age" in res["groups"][0]

    # errors
    with pytest.raises(ValueError, match="Unsupported operation"):
        aggregate_data(sample_xlsx, "Sheet1", group_by="City", value_column="Salary", operation="INVALID")

    with pytest.raises(ValueError, match="Column 'INVALID' not found"):
        aggregate_data(sample_xlsx, "Sheet1", group_by="INVALID", value_column="Salary")


def test_vlookup_helper_fuzzy(tmp_path):
    """Test fuzzy vlookup."""
    path1 = str(tmp_path / "lookup.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "ID"])
    ws.append(["Alis", 101])  # misspelled Alice
    wb.save(path1)

    path2 = str(tmp_path / "data.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Score"])
    ws.append(["Alice", 100])
    wb.save(path2)

    res = vlookup_helper(
        path1, path2, lookup_column="A", data_key_column="A", data_return_columns=["B"], fuzzy=True, fuzzy_threshold=0.5
    )
    assert res["matched"] == 1
    assert res["results"][0]["matched_value"] == "Alice"


def test_vlookup_helper_output_file(tmp_path):
    """Test vlookup with output file."""
    path1 = str(tmp_path / "lookup.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "ID"])
    ws.append(["Alice", 101])
    wb.save(path1)

    path2 = str(tmp_path / "data.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Score"])
    ws.append(["Alice", 100])
    wb.save(path2)

    out_path = str(tmp_path / "out.xlsx")
    vlookup_helper(
        path1, path2, lookup_column="A", data_key_column="A", data_return_columns=["B"], output_file=out_path
    )

    from openpyxl import load_workbook

    wb = load_workbook(out_path)
    assert "VLookup Results" in wb.sheetnames
    wb.close()
