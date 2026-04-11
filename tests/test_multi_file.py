"""Tests for multi_file.py — cross-file operations."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from mcp_server.tools.multi_file import (
    bulk_aggregate_multi_files,
    bulk_filter_multi_files,
    validate_data_consistency,
)


def _create_file(path: str, headers: list, rows: list) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)
    return path


@pytest.fixture()
def multi_files(tmp_path: Path) -> list[str]:
    """Create 3 files with overlapping data."""
    f1 = _create_file(
        str(tmp_path / "sales_q1.xlsx"),
        ["Product", "Revenue", "Region"],
        [["Widget", 100, "East"], ["Gadget", 200, "West"], ["Widget", 150, "East"]],
    )
    f2 = _create_file(
        str(tmp_path / "sales_q2.xlsx"),
        ["Product", "Revenue", "Region"],
        [["Widget", 120, "East"], ["Gadget", 180, "West"], ["Doohickey", 300, "North"]],
    )
    f3 = _create_file(
        str(tmp_path / "sales_q3.xlsx"),
        ["Product", "Revenue", "Region"],
        [["Widget", 130, "East"], ["Gadget", 220, "West"]],
    )
    return [f1, f2, f3]


@pytest.fixture()
def consistency_files(tmp_path: Path) -> list[str]:
    """Create files with intentional inconsistencies."""
    f1 = _create_file(
        str(tmp_path / "master.xlsx"),
        ["ID", "Name", "Price"],
        [[1, "Widget", 10.0], [2, "Gadget", 20.0], [3, "Doohickey", 30.0]],
    )
    f2 = _create_file(
        str(tmp_path / "branch_a.xlsx"),
        ["ID", "Name", "Price"],
        [[1, "Widget", 10.0], [2, "Gadget", 25.0]],  # Price mismatch for ID=2, ID=3 missing
    )
    return [f1, f2]


def test_bulk_aggregate_sum(multi_files: list[str]) -> None:
    result = bulk_aggregate_multi_files(multi_files, column="Revenue", operation="sum")
    assert result["total_files"] == 3
    assert result["operation"] == "sum"
    assert len(result["per_file"]) == 3
    assert result["aggregate"] == 100 + 200 + 150 + 120 + 180 + 300 + 130 + 220


def test_bulk_aggregate_mean(multi_files: list[str]) -> None:
    result = bulk_aggregate_multi_files(multi_files, column="Revenue", operation="mean")
    assert result["operation"] == "mean"
    assert result["aggregate"] > 0


def test_bulk_aggregate_output_file(multi_files: list[str], tmp_path: Path) -> None:
    output = str(tmp_path / "aggregate_summary.xlsx")
    result = bulk_aggregate_multi_files(multi_files, column="Revenue", operation="sum", output_file=output)
    assert Path(output).exists()
    assert result["aggregate"] > 0


def test_bulk_aggregate_invalid_operation(multi_files: list[str]) -> None:
    with pytest.raises(ValueError, match="Unsupported operation"):
        bulk_aggregate_multi_files(multi_files, column="Revenue", operation="mode")


def test_bulk_filter_equals(multi_files: list[str]) -> None:
    result = bulk_filter_multi_files(multi_files, column="Product", operator="equals", value="Widget")
    assert result["total_matched"] >= 3  # Widget appears in all 3 files
    assert result["total_files"] == 3


def test_bulk_filter_greater_than(multi_files: list[str]) -> None:
    result = bulk_filter_multi_files(multi_files, column="Revenue", operator="greater_than", value=150)
    assert result["total_matched"] > 0
    for pf in result["per_file"]:
        for row in pf["data"]:
            assert row[1] > 150  # Revenue is column index 1


def test_bulk_filter_contains(multi_files: list[str]) -> None:
    result = bulk_filter_multi_files(multi_files, column="Product", operator="contains", value="get")
    assert result["total_matched"] >= 3  # "Gadget" contains "get"


def test_bulk_filter_output_file(multi_files: list[str], tmp_path: Path) -> None:
    output = str(tmp_path / "filtered.xlsx")
    result = bulk_filter_multi_files(
        multi_files, column="Product", operator="equals", value="Widget", output_file=output
    )
    assert Path(output).exists()
    assert result["total_matched"] > 0


def test_validate_consistency_all_match(tmp_path: Path) -> None:
    """All files have the same keys and values — should be consistent."""
    f1 = _create_file(
        str(tmp_path / "a.xlsx"),
        ["ID", "Name"],
        [[1, "Alice"], [2, "Bob"]],
    )
    f2 = _create_file(
        str(tmp_path / "b.xlsx"),
        ["ID", "Name"],
        [[1, "Alice"], [2, "Bob"]],
    )
    result = validate_data_consistency([f1, f2], key_column="ID", check_columns=["Name"])
    assert result["consistent"] is True
    assert len(result["missing_keys"]) == 0
    assert len(result["mismatched_values"]) == 0


def test_validate_consistency_missing_keys(consistency_files: list[str]) -> None:
    result = validate_data_consistency(consistency_files, key_column="ID")
    assert result["consistent"] is False
    assert len(result["missing_keys"]) > 0  # ID=3 missing from branch_a


def test_validate_consistency_mismatched_values(consistency_files: list[str]) -> None:
    result = validate_data_consistency(consistency_files, key_column="ID", check_columns=["Price"])
    assert result["consistent"] is False
    assert len(result["mismatched_values"]) > 0  # Price mismatch for ID=2


def test_validate_consistency_invalid_column(consistency_files: list[str]) -> None:
    result = validate_data_consistency(consistency_files, key_column="NonExistent")
    assert result["consistent"] is False
    assert len(result["schema_mismatches"]) > 0


from mcp_server.tools.multi_file import compare_workbooks


def test_compare_workbooks(tmp_path: Path) -> None:
    f1 = _create_file(str(tmp_path / "comp1.xlsx"), ["A", "B"], [[1, 2]])
    f2 = _create_file(str(tmp_path / "comp2.xlsx"), ["A", "B"], [[1, 3]])

    result = compare_workbooks(f1, f2, sheet_name="Sheet1")
    assert result["total_differences"] == 1


# ============================================================
# Additional comprehensive tests
# ============================================================


def test_bulk_aggregate_two_files(tmp_path: Path) -> None:
    """Sum a column across exactly two files, verifying per-file and aggregate values."""
    f1 = _create_file(str(tmp_path / "agg_a.xlsx"), ["Sales"], [[100], [200]])
    f2 = _create_file(str(tmp_path / "agg_b.xlsx"), ["Sales"], [[300]])

    result = bulk_aggregate_multi_files([f1, f2], column="Sales", operation="sum")
    assert result["total_files"] == 2
    assert result["aggregate"] == pytest.approx(600.0)  # 100+200+300
    assert len(result["per_file"]) == 2


def test_bulk_aggregate_missing_column(tmp_path: Path) -> None:
    """Raises ValueError when the requested column is absent in one of the files."""
    f1 = _create_file(str(tmp_path / "mc_f1.xlsx"), ["Revenue"], [[100], [200]])
    f2 = _create_file(str(tmp_path / "mc_f2.xlsx"), ["Amount"], [[300], [400]])  # different col

    with pytest.raises(ValueError, match="Column 'Revenue' not found"):
        bulk_aggregate_multi_files([f1, f2], column="Revenue", operation="sum")


def test_bulk_aggregate_min_max(multi_files: list[str]) -> None:
    """Test min and max aggregations across multi_files fixture."""
    # Revenues in fixture: 100, 200, 150, 120, 180, 300, 130, 220
    min_result = bulk_aggregate_multi_files(multi_files, column="Revenue", operation="min")
    assert min_result["aggregate"] == pytest.approx(100.0)

    max_result = bulk_aggregate_multi_files(multi_files, column="Revenue", operation="max")
    assert max_result["aggregate"] == pytest.approx(300.0)


def test_bulk_aggregate_nonexistent_file(tmp_path: Path) -> None:
    """Raises an exception when a file path does not exist."""
    f1 = _create_file(str(tmp_path / "exists.xlsx"), ["Revenue"], [[100]])
    nonexistent = str(tmp_path / "missing_file.xlsx")

    with pytest.raises(Exception):
        bulk_aggregate_multi_files([f1, nonexistent], column="Revenue", operation="sum")


def test_bulk_filter_not_equals(multi_files: list[str]) -> None:
    """Filter rows where Product is NOT 'Widget'; result contains no Widget rows."""
    result = bulk_filter_multi_files(multi_files, column="Product", operator="not_equals", value="Widget")
    assert result["total_matched"] > 0
    for pf in result["per_file"]:
        for row in pf["data"]:
            assert row[0] != "Widget"  # Product column is index 0


def test_bulk_filter_less_than(multi_files: list[str]) -> None:
    """Filter rows where Revenue < 150; all matched rows satisfy the condition."""
    result = bulk_filter_multi_files(multi_files, column="Revenue", operator="less_than", value=150)
    assert result["total_matched"] > 0  # 100, 120, 130 are < 150
    for pf in result["per_file"]:
        for row in pf["data"]:
            assert row[1] < 150  # Revenue is index 1


def test_validate_consistency_no_check_columns(tmp_path: Path) -> None:
    """With no check_columns, only key presence is checked; matching keys → consistent."""
    f1 = _create_file(str(tmp_path / "vcc_a.xlsx"), ["ID", "Name"], [[1, "X"], [2, "Y"]])
    f2 = _create_file(str(tmp_path / "vcc_b.xlsx"), ["ID", "Name"], [[1, "X"], [2, "Y"]])

    result = validate_data_consistency([f1, f2], key_column="ID")
    assert result["consistent"] is True
    assert result["total_keys"] == 2
    assert len(result["missing_keys"]) == 0


def test_compare_workbooks_identical(tmp_path: Path) -> None:
    """Identical workbooks produce zero differences and identical=True."""
    f1 = _create_file(str(tmp_path / "id1.xlsx"), ["A", "B"], [[1, 2], [3, 4]])
    f2 = _create_file(str(tmp_path / "id2.xlsx"), ["A", "B"], [[1, 2], [3, 4]])

    result = compare_workbooks(f1, f2, sheet_name="Sheet1")
    assert result["identical"] is True
    assert result["total_differences"] == 0
    assert result["differences"] == []


def test_compare_workbooks_multiple_differences(tmp_path: Path) -> None:
    """Workbooks differing in multiple cells report each difference with details."""
    f1 = _create_file(str(tmp_path / "md1.xlsx"), ["X", "Y"], [[10, 20], [30, 40]])
    f2 = _create_file(str(tmp_path / "md2.xlsx"), ["X", "Y"], [[10, 99], [30, 99]])

    result = compare_workbooks(f1, f2, sheet_name="Sheet1")
    assert result["identical"] is False
    assert result["total_differences"] == 2  # Y2 and Y3 differ
    for diff in result["differences"]:
        assert "cell" in diff
        assert "value_a" in diff
        assert "value_b" in diff
