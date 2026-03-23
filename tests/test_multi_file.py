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
    with pytest.raises(ValueError, match="Key column.*not found"):
        validate_data_consistency(consistency_files, key_column="NonExistent")
