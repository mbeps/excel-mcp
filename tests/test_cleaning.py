"""Tests for cleaning.py — data cleaning pipeline."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from mcp_server.tools.cleaning import combine_columns, data_cleaner, detect_outliers, split_column


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


# ---------------------------------------------------------------------------
# Fixtures for new function tests
# ---------------------------------------------------------------------------


@pytest.fixture()
def split_xlsx(tmp_path: Path) -> str:
    """Workbook with a delimited text column suitable for split_column tests."""
    path = str(tmp_path / "split.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["FullName", "Score"])
    ws.append(["Alice,Smith", 80])
    ws.append(["Bob,Jones", 90])
    ws.append(["Charlie,Brown", 75])
    wb.save(path)
    return path


@pytest.fixture()
def numeric_xlsx(tmp_path: Path) -> str:
    """Workbook with numeric data including one extreme outlier."""
    path = str(tmp_path / "numeric.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Value", "Category"])
    for v in [10, 12, 11, 13, 9, 11, 100]:
        ws.append([v, "A"])
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# split_column tests
# ---------------------------------------------------------------------------


def test_split_column_basic(split_xlsx: str) -> None:
    result = split_column(split_xlsx, column="A", delimiter=",")
    assert result["rows_affected"] == 3
    assert result["new_columns"] == ["FullName_1", "FullName_2"]


def test_split_column_custom_names(split_xlsx: str) -> None:
    result = split_column(split_xlsx, column="FullName", delimiter=",", new_column_names=["First", "Last"])
    assert result["new_columns"] == ["First", "Last"]


def test_split_column_keep_original(split_xlsx: str, tmp_path: Path) -> None:
    out = str(tmp_path / "out.xlsx")
    result = split_column(split_xlsx, column="A", delimiter=",", drop_original=False, output_file=out)
    assert Path(out).exists()
    assert len(result["new_columns"]) == 2


def test_split_column_output_file(split_xlsx: str, tmp_path: Path) -> None:
    out = str(tmp_path / "split_out.xlsx")
    result = split_column(split_xlsx, column="A", delimiter=",", output_file=out)
    assert result["output_file"] == out
    assert Path(out).exists()


# ---------------------------------------------------------------------------
# combine_columns tests
# ---------------------------------------------------------------------------


def test_combine_columns_basic(dirty_xlsx: str) -> None:
    result = combine_columns(dirty_xlsx, columns=["A", "C"], new_column_name="NameCity", separator="-")
    assert result["new_column"] == "NameCity"
    assert result["rows_affected"] > 0


def test_combine_columns_by_name(dirty_xlsx: str, tmp_path: Path) -> None:
    out = str(tmp_path / "combined.xlsx")
    result = combine_columns(
        dirty_xlsx,
        columns=["Name", "City"],
        new_column_name="Full",
        separator=" | ",
        drop_originals=True,
        output_file=out,
    )
    assert result["rows_affected"] > 0
    assert Path(out).exists()


def test_combine_columns_no_columns_raises(dirty_xlsx: str) -> None:
    with pytest.raises(ValueError, match="non-empty"):
        combine_columns(dirty_xlsx, columns=[], new_column_name="X")


# ---------------------------------------------------------------------------
# detect_outliers tests
# ---------------------------------------------------------------------------


def test_detect_outliers_iqr_flag(numeric_xlsx: str, tmp_path: Path) -> None:
    out = str(tmp_path / "flagged.xlsx")
    result = detect_outliers(numeric_xlsx, column="A", method="iqr", threshold=1.5, action="flag", output_file=out)
    assert result["method"] == "iqr"
    assert result["outliers_found"] >= 1
    assert isinstance(result["outlier_rows"], list)
    assert Path(out).exists()


def test_detect_outliers_zscore_flag(numeric_xlsx: str, tmp_path: Path) -> None:
    out = str(tmp_path / "zscore.xlsx")
    result = detect_outliers(
        numeric_xlsx, column="Value", method="zscore", threshold=2.0, action="flag", output_file=out
    )
    assert result["method"] == "zscore"
    assert result["outliers_found"] >= 1


def test_detect_outliers_remove(numeric_xlsx: str, tmp_path: Path) -> None:
    out = str(tmp_path / "removed.xlsx")
    result = detect_outliers(numeric_xlsx, column="A", method="iqr", threshold=1.5, action="remove", output_file=out)
    assert result["action"] == "remove"
    assert result["outliers_found"] >= 1
    assert Path(out).exists()


def test_detect_outliers_invalid_method(numeric_xlsx: str) -> None:
    with pytest.raises(ValueError, match="method"):
        detect_outliers(numeric_xlsx, column="A", method="invalid")


def test_detect_outliers_custom_flag_name(numeric_xlsx: str, tmp_path: Path) -> None:
    out = str(tmp_path / "custom_flag.xlsx")
    result = detect_outliers(
        numeric_xlsx, column="Value", flag_column_name="is_outlier", action="flag", output_file=out
    )
    assert result["action"] == "flag"
    assert result["outliers_found"] >= 1
