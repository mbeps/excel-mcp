"""Tests for miscellaneous bug fixes: parse_date_column, list_charts position,
combo chart index validation, multi_file filter operators and duplicate columns."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest
from openpyxl import Workbook

from mcp_server.tools.charts import create_chart, create_combo_chart, list_charts
from mcp_server.tools.cleaning import parse_date_column
from mcp_server.tools.multi_file import bulk_filter_multi_files

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_date_workbook(tmp_path: Path, name: str, dates: list[object]) -> str:
    """Create a workbook with a single 'Date' column."""
    path = str(tmp_path / name)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Date"])
    for d in dates:
        ws.append([d])
    wb.save(path)
    wb.close()
    return path


def _make_chart_workbook(tmp_path: Path, name: str = "chart.xlsx") -> str:
    path = str(tmp_path / name)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Month", "Sales", "Costs", "Profit"])
    ws.append(["Jan", 100, 60, 40])
    ws.append(["Feb", 150, 80, 70])
    ws.append(["Mar", 200, 90, 110])
    ws.append(["Apr", 180, 85, 95])
    ws.append(["May", 220, 100, 120])
    wb.save(path)
    wb.close()
    return path


def _create_file(path: str, headers: list, rows: list) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)
    wb.close()
    return path


# ===========================================================================
# parse_date_column tests
# ===========================================================================


class TestParseDateColumn:
    """Tests for parse_date_column bug fixes and behaviour."""

    def test_parse_date_iso_format(self, tmp_path: Path) -> None:
        """ISO dates like '2026-03-01' and '2026-06-15' parse correctly with month preserved."""
        path = _make_date_workbook(tmp_path, "iso.xlsx", ["2026-03-01", "2026-06-15"])
        result = parse_date_column(path, "Sheet1", column="Date")
        assert result["parsed_count"] == 2
        assert result["failed_count"] == 0

        wb = openpyxl.load_workbook(path)
        ws = wb.active
        v1 = ws.cell(row=2, column=1).value
        v2 = ws.cell(row=3, column=1).value
        wb.close()
        assert v1 == "2026-03-01"
        assert v2 == "2026-06-15"

    def test_parse_date_mixed_formats(self, tmp_path: Path) -> None:
        """Various date formats all parse successfully."""
        dates = ["2026-03-01", "03/15/2026", "4 Mar 2026", "March 20, 2026"]
        path = _make_date_workbook(tmp_path, "mixed.xlsx", dates)
        result = parse_date_column(path, "Sheet1", column="Date")
        assert result["parsed_count"] == 4
        assert result["failed_count"] == 0

    def test_parse_date_dayfirst_false_default(self, tmp_path: Path) -> None:
        """Default dayfirst=False: '2026-03-01' is March 1st, not Jan 3rd."""
        path = _make_date_workbook(tmp_path, "dayfirst_false.xlsx", ["2026-03-01"])
        result = parse_date_column(path, "Sheet1", column="Date", dayfirst=False)
        assert result["parsed_count"] == 1

        wb = openpyxl.load_workbook(path)
        ws = wb.active
        val = ws.cell(row=2, column=1).value
        wb.close()
        assert val is not None
        assert val.startswith("2026-03-01")

    def test_parse_date_dayfirst_true_override(self, tmp_path: Path) -> None:
        """With dayfirst=True, '01/03/2026' is interpreted as March 1st (day-first)."""
        path = _make_date_workbook(tmp_path, "dayfirst_true.xlsx", ["01/03/2026"])
        result = parse_date_column(path, "Sheet1", column="Date", dayfirst=True)
        assert result["parsed_count"] == 1

        wb = openpyxl.load_workbook(path)
        ws = wb.active
        val = ws.cell(row=2, column=1).value
        wb.close()
        # dayfirst=True: 01/03/2026 → day=01, month=03 → March 1st
        assert val is not None
        assert val.startswith("2026-03-01")

    def test_parse_date_already_datetime(self, tmp_path: Path) -> None:
        """Cells already containing datetime objects pass through unchanged."""
        dt1 = datetime(2026, 5, 10, 14, 30)
        dt2 = datetime(2026, 12, 25)
        path = _make_date_workbook(tmp_path, "datetime.xlsx", [dt1, dt2])
        result = parse_date_column(path, "Sheet1", column="Date")
        assert result["parsed_count"] == 2
        assert result["failed_count"] == 0

        wb = openpyxl.load_workbook(path)
        ws = wb.active
        v1 = ws.cell(row=2, column=1).value
        v2 = ws.cell(row=3, column=1).value
        wb.close()
        assert v1 == "2026-05-10"
        assert v2 == "2026-12-25"

    def test_parse_date_null_and_empty(self, tmp_path: Path) -> None:
        """NaN/None/empty string values produce None, not a crash."""
        path = _make_date_workbook(tmp_path, "nulls.xlsx", [None, "", "2026-01-15"])
        result = parse_date_column(path, "Sheet1", column="Date")
        assert result["parsed_count"] == 1
        assert result["failed_count"] == 2

        wb = openpyxl.load_workbook(path)
        ws = wb.active
        assert ws.cell(row=2, column=1).value is None
        assert ws.cell(row=3, column=1).value is None
        assert ws.cell(row=4, column=1).value == "2026-01-15"
        wb.close()

    def test_parse_date_output_column(self, tmp_path: Path) -> None:
        """output_column writes to a new column instead of overwriting the original."""
        path = _make_date_workbook(tmp_path, "outcol.xlsx", ["2026-03-01", "2026-06-15"])
        result = parse_date_column(path, "Sheet1", column="Date", output_column="ParsedDate")
        assert result["output_column"] == "ParsedDate"

        wb = openpyxl.load_workbook(path)
        ws = wb.active
        # Original column should still contain raw values
        assert ws.cell(row=1, column=1).value == "Date"
        # New column header
        assert ws.cell(row=1, column=2).value == "ParsedDate"
        assert ws.cell(row=2, column=2).value == "2026-03-01"
        assert ws.cell(row=3, column=2).value == "2026-06-15"
        wb.close()

    def test_parse_date_output_format(self, tmp_path: Path) -> None:
        """Custom output_format like '%d/%m/%Y' works correctly."""
        path = _make_date_workbook(tmp_path, "fmt.xlsx", ["2026-03-01", "2026-12-25"])
        result = parse_date_column(path, "Sheet1", column="Date", output_format="%d/%m/%Y")
        assert result["parsed_count"] == 2

        wb = openpyxl.load_workbook(path)
        ws = wb.active
        v1 = ws.cell(row=2, column=1).value
        v2 = ws.cell(row=3, column=1).value
        wb.close()
        assert v1 == "01/03/2026"
        assert v2 == "25/12/2026"


# ===========================================================================
# list_charts position tests
# ===========================================================================


class TestListChartsPosition:
    """Verify list_charts returns human-readable cell references (not AnchorMarker repr)."""

    def test_list_charts_position_is_cell_reference(self, tmp_path: Path) -> None:
        """Position should be a cell reference like 'E1', not an AnchorMarker repr."""
        path = _make_chart_workbook(tmp_path)
        create_chart(path, "Sheet1", "A1:D6", chart_type="column", target_cell="E1", title="Pos Test")
        charts = list_charts(path, "Sheet1")
        assert len(charts) == 1
        pos = charts[0]["position"]
        # Must be a clean cell reference, not contain angle brackets or 'object at'
        assert "<" not in str(pos)
        assert "object" not in str(pos).lower()
        assert "AnchorMarker" not in str(pos)

    def test_list_charts_position_format(self, tmp_path: Path) -> None:
        """Position matches a cell reference pattern like 'A1', 'E2', etc."""
        path = _make_chart_workbook(tmp_path)
        create_chart(path, "Sheet1", "A1:D6", chart_type="column", target_cell="F10", title="Fmt Test")
        charts = list_charts(path, "Sheet1")
        assert len(charts) == 1
        pos = str(charts[0]["position"])
        # Cell references match: one or more letters followed by one or more digits
        assert re.match(r"^[A-Z]+\d+$", pos), f"Position '{pos}' is not a valid cell reference"


# ===========================================================================
# create_combo_chart index validation tests
# ===========================================================================


class TestComboChartIndexValidation:
    """Verify create_combo_chart rejects 0-based indices with clear errors."""

    def test_combo_chart_zero_index_bar_error(self, tmp_path: Path) -> None:
        """bar_columns=[0] should raise ValueError about 1-based indices."""
        path = _make_chart_workbook(tmp_path)
        with pytest.raises(ValueError, match="1-based"):
            create_combo_chart(path, "Sheet1", "A1:D6", bar_columns=[0], line_columns=[2], title="Bad Bar")

    def test_combo_chart_zero_index_line_error(self, tmp_path: Path) -> None:
        """line_columns=[0] should raise ValueError about 1-based indices."""
        path = _make_chart_workbook(tmp_path)
        with pytest.raises(ValueError, match="1-based"):
            create_combo_chart(path, "Sheet1", "A1:D6", bar_columns=[1], line_columns=[0], title="Bad Line")

    def test_combo_chart_valid_indices(self, tmp_path: Path) -> None:
        """bar_columns=[1], line_columns=[2] should succeed without error."""
        path = _make_chart_workbook(tmp_path)
        result = create_combo_chart(path, "Sheet1", "A1:D6", bar_columns=[1], line_columns=[2], title="Valid Combo")
        assert "combo" in result.lower()


# ===========================================================================
# multi_file filter: operator aliases and duplicate columns
# ===========================================================================


class TestMultiFileFilterOperators:
    """Verify symbolic operator aliases and no duplicate columns in results."""

    @pytest.fixture()
    def filter_files(self, tmp_path: Path) -> list[str]:
        """Two files with 5 columns for operator and column-duplication tests."""
        f1 = _create_file(
            str(tmp_path / "f1.xlsx"),
            ["ID", "Name", "Score", "Region", "Active"],
            [[1, "Alice", 90, "East", "Yes"], [2, "Bob", 70, "West", "No"], [3, "Carol", 85, "East", "Yes"]],
        )
        f2 = _create_file(
            str(tmp_path / "f2.xlsx"),
            ["ID", "Name", "Score", "Region", "Active"],
            [[4, "Dave", 60, "North", "Yes"], [5, "Eve", 95, "South", "No"]],
        )
        return [f1, f2]

    def test_filter_no_duplicate_columns(self, filter_files: list[str]) -> None:
        """Result data rows should have the same number of columns as the original (not doubled)."""
        result = bulk_filter_multi_files(filter_files, column="Name", operator="equals", value="Alice")
        assert result["total_matched"] >= 1
        for pf in result["per_file"]:
            for row in pf["data"]:
                # Original has 5 columns; data rows must not exceed 5
                assert len(row) == 5, f"Expected 5 columns, got {len(row)}: {row}"

    def test_filter_symbolic_operator_equals(self, filter_files: list[str]) -> None:
        """Symbolic '==' normalises to 'equals' internally."""
        result = bulk_filter_multi_files(filter_files, column="Name", operator="==", value="Alice")
        assert result["operator"] == "equals"
        assert result["total_matched"] >= 1

    def test_filter_symbolic_operator_greater_than(self, filter_files: list[str]) -> None:
        """Symbolic '>' normalises to 'greater_than'."""
        result = bulk_filter_multi_files(filter_files, column="Score", operator=">", value=80)
        assert result["operator"] == "greater_than"
        assert result["total_matched"] >= 2  # 90, 85, 95

    def test_filter_symbolic_operator_less_than(self, filter_files: list[str]) -> None:
        """Symbolic '<' normalises to 'less_than'."""
        result = bulk_filter_multi_files(filter_files, column="Score", operator="<", value=80)
        assert result["operator"] == "less_than"
        assert result["total_matched"] >= 2  # 70, 60

    def test_filter_word_operator_still_works(self, filter_files: list[str]) -> None:
        """Word-form 'equals' still works as before."""
        result = bulk_filter_multi_files(filter_files, column="Name", operator="equals", value="Bob")
        assert result["operator"] == "equals"
        assert result["total_matched"] >= 1

    def test_filter_greater_than_or_equal(self, filter_files: list[str]) -> None:
        """Symbolic '>=' normalises to 'greater_than_or_equal'."""
        result = bulk_filter_multi_files(filter_files, column="Score", operator=">=", value=90)
        assert result["operator"] == "greater_than_or_equal"
        assert result["total_matched"] >= 2  # 90, 95

    def test_filter_invalid_operator_raises(self, filter_files: list[str]) -> None:
        """Invalid operator raises ValueError."""
        with pytest.raises(ValueError, match="Unsupported operator"):
            bulk_filter_multi_files(filter_files, column="Name", operator="INVALID", value="x")

    def test_filter_multiple_files_no_duplicate_columns(self, filter_files: list[str]) -> None:
        """Filtering across 2 files with same schema must not have duplicated columns."""
        result = bulk_filter_multi_files(filter_files, column="Score", operator=">", value=50)
        assert result["total_matched"] >= 4  # most rows have Score > 50
        for pf in result["per_file"]:
            for row in pf["data"]:
                assert len(row) == 5, f"Expected 5 columns, got {len(row)}: {row}"
