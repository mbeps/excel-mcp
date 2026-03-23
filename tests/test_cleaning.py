"""Tests for cleaning.py — data cleaning pipeline."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from mcp_server.tools.cleaning import data_cleaner


@pytest.fixture()
def dirty_xlsx(tmp_path: Path) -> str:
    """Create a workbook with messy data for cleaning tests."""
    path = str(tmp_path / "dirty.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Age", "City", "Salary"])
    ws.append(["  Alice  ", 30, "New York", "70000"])
    ws.append(["Bob", 25, "  Chicago  ", "$55,000"])
    ws.append(["  alice  ", 30, "new york", "70000"])  # duplicate (after normalize)
    ws.append([None, None, None, None])  # empty row
    ws.append(["Charlie", 35, "Boston", "€90,000"])
    ws.append(["  Diana  ", "28", "Chicago", "62000"])
    wb.save(path)
    return path


def test_data_cleaner_trim_whitespace(dirty_xlsx: str) -> None:
    result = data_cleaner(dirty_xlsx, operations=["trim_whitespace"])
    assert result["preview_only"] is False
    assert "trim_whitespace" in result["operations_applied"]
    assert result["changes"]["trim_whitespace"] > 0
    assert result["rows_before"] == result["rows_after"]


def test_data_cleaner_remove_empty_rows(dirty_xlsx: str) -> None:
    result = data_cleaner(dirty_xlsx, operations=["remove_empty_rows"])
    assert result["rows_after"] < result["rows_before"]
    assert result["changes"]["remove_empty_rows"] >= 1


def test_data_cleaner_fix_numbers(dirty_xlsx: str) -> None:
    result = data_cleaner(dirty_xlsx, operations=["fix_numbers"])
    assert result["changes"]["fix_numbers"] > 0


def test_data_cleaner_remove_duplicates(dirty_xlsx: str) -> None:
    result = data_cleaner(dirty_xlsx, operations=["normalize_text", "remove_duplicates"])
    assert result["changes"]["remove_duplicates"] >= 1
    assert result["rows_after"] < result["rows_before"]


def test_data_cleaner_fill_missing(dirty_xlsx: str) -> None:
    result = data_cleaner(dirty_xlsx, operations=["fill_missing"])
    assert result["changes"]["fill_missing"] >= 4  # the empty row has 4 empty cells


def test_data_cleaner_preview_mode(dirty_xlsx: str) -> None:
    result = data_cleaner(dirty_xlsx, operations=["trim_whitespace"], preview=True)
    assert result["preview_only"] is True
    assert "sample_changes" in result
    assert isinstance(result["sample_changes"], list)


def test_data_cleaner_output_file(dirty_xlsx: str, tmp_path: Path) -> None:
    output = str(tmp_path / "cleaned.xlsx")
    result = data_cleaner(dirty_xlsx, operations=["trim_whitespace", "remove_empty_rows"], output_file=output)
    assert result["preview_only"] is False
    assert Path(output).exists()


def test_data_cleaner_column_filter(dirty_xlsx: str) -> None:
    result = data_cleaner(dirty_xlsx, operations=["trim_whitespace"], columns=["A"])
    assert result["changes"]["trim_whitespace"] > 0


def test_data_cleaner_all_operations(dirty_xlsx: str) -> None:
    result = data_cleaner(dirty_xlsx)
    assert len(result["operations_applied"]) == 7
    assert result["rows_after"] <= result["rows_before"]


def test_data_cleaner_invalid_operation(dirty_xlsx: str) -> None:
    with pytest.raises(ValueError, match="Unknown operation"):
        data_cleaner(dirty_xlsx, operations=["nonexistent_op"])
