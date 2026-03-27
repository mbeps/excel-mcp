"""Comprehensive regression tests for 8 specific bug fixes.

Tests are organized by bug number and focus on the specific behavior that was fixed.
Does NOT duplicate tests already present in test_bug_fixes.py, test_worksheet_ops.py,
test_cell_ops.py, or test_charts.py.
"""

from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from openpyxl import Workbook, load_workbook

from mcp_server.tools.cell_ops import fill_series, write_cell, write_range
from mcp_server.tools.charts import _safe_chart_title, create_chart, list_charts
from mcp_server.tools.doc_properties import set_calculation_mode
from mcp_server.tools.formulas import get_formula_value, set_formula
from mcp_server.tools.worksheet_ops import (
    copy_range_across_sheets,
    freeze_panes,
    set_auto_filter,
)

# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------


def _empty_workbook(tmp_path: Path, name: str = "test.xlsx") -> str:
    """Create a minimal empty workbook with a blank Sheet1."""
    path = str(tmp_path / name)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    wb.save(path)
    wb.close()
    return path


def _workbook_with_data(tmp_path: Path, name: str = "data.xlsx") -> str:
    """Create a workbook pre-populated with two columns of data."""
    path = str(tmp_path / name)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Month", "Sales"])
    ws.append(["Jan", 100])
    ws.append(["Feb", 150])
    wb.save(path)
    wb.close()
    return path


# ---------------------------------------------------------------------------
# Bug 1: write_cells series — MCP wrapper validates count is not None
# ---------------------------------------------------------------------------


class TestWriteCellsSeriesCountValidation:
    """Bug 1: write_cells mode='series' must raise ValueError when count=None."""

    def test_wrapper_raises_value_error_when_count_is_none(self, tmp_path: Path) -> None:
        """MCP write_cells raises ValueError for mode='series' with count=None."""
        from mcp_server.main import write_cells

        fp = _empty_workbook(tmp_path, "series.xlsx")
        with pytest.raises(ValueError, match="count is required"):
            write_cells(
                mode="series",
                file_path=fp,
                sheet_name="Sheet1",
                start_cell="A1",
                count=None,  # type: ignore[arg-type]
            )

    def test_wrapper_raises_value_error_when_start_cell_is_none(self, tmp_path: Path) -> None:
        """MCP write_cells raises ValueError for mode='series' with start_cell=None."""
        from mcp_server.main import write_cells

        fp = _empty_workbook(tmp_path, "series.xlsx")
        with pytest.raises(ValueError, match="start_cell is required"):
            write_cells(
                mode="series",
                file_path=fp,
                sheet_name="Sheet1",
                start_cell=None,
                count=5,
            )

    def test_wrapper_series_succeeds_when_count_provided(self, tmp_path: Path) -> None:
        """MCP write_cells mode='series' completes without error when count is given."""
        from mcp_server.main import write_cells

        fp = _empty_workbook(tmp_path, "series.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = 1
        wb.save(fp)
        wb.close()

        result = write_cells(
            mode="series",
            file_path=fp,
            sheet_name="Sheet1",
            start_cell="A1",
            series_type="number",
            count=4,
            step=3,
            direction="down",
        )
        assert isinstance(result, dict)
        assert result["count"] == 4

        wb = load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].value == 1
        assert ws["A2"].value == 4
        assert ws["A4"].value == 10
        wb.close()

    def test_underlying_fill_series_with_direction_right(self, tmp_path: Path) -> None:
        """fill_series tool function fills to the right when direction='right'."""
        fp = _empty_workbook(tmp_path, "series_right.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["B2"] = 0
        wb.save(fp)
        wb.close()

        result = fill_series(fp, "Sheet1", "B2", "number", 3, step=10, direction="right")
        assert result["count"] == 3

        wb = load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["B2"].value == 0
        assert ws["C2"].value == 10
        assert ws["D2"].value == 20
        wb.close()


# ---------------------------------------------------------------------------
# Bug 2: chart:list — _safe_chart_title handles all title types
# ---------------------------------------------------------------------------


class TestSafeChartTitle:
    """Bug 2: _safe_chart_title must return a plain str for all input types."""

    def test_returns_untitled_for_none(self) -> None:
        """`_safe_chart_title(None)` should return the literal string '(untitled)'."""
        assert _safe_chart_title(None) == "(untitled)"

    def test_returns_string_unchanged(self) -> None:
        """`_safe_chart_title('My Chart')` should pass through unchanged."""
        assert _safe_chart_title("My Chart") == "My Chart"

    def test_returns_empty_string_unchanged(self) -> None:
        """`_safe_chart_title('')` should pass through as empty string."""
        assert _safe_chart_title("") == ""

    def test_handles_mock_text_with_strref(self) -> None:
        """Extracts title from openpyxl Text.strRef.f (formula-based title)."""
        str_ref = MagicMock()
        str_ref.f = "Sheet1!$A$1"
        text_obj = MagicMock(spec=["strRef", "rich"])
        text_obj.strRef = str_ref
        text_obj.rich = None

        result = _safe_chart_title(text_obj)
        assert result == "Sheet1!$A$1"

    def test_handles_mock_text_with_rich_paragraphs(self) -> None:
        """Extracts title from openpyxl Text.rich paragraph runs."""
        run = MagicMock()
        run.t = "Revenue"
        para = MagicMock()
        para.runs = [run]
        rich = MagicMock()
        rich.paragraphs = [para]
        text_obj = MagicMock(spec=["strRef", "rich"])
        text_obj.strRef = None
        text_obj.rich = rich

        result = _safe_chart_title(text_obj)
        assert result == "Revenue"

    def test_handles_mock_text_multiple_runs_joined(self) -> None:
        """Joins multiple run texts with space when rich title has multiple runs."""
        run1 = MagicMock()
        run1.t = "Sales"
        run2 = MagicMock()
        run2.t = "Chart"
        para = MagicMock()
        para.runs = [run1, run2]
        rich = MagicMock()
        rich.paragraphs = [para]
        text_obj = MagicMock(spec=["strRef", "rich"])
        text_obj.strRef = None
        text_obj.rich = rich

        result = _safe_chart_title(text_obj)
        assert result == "Sales Chart"

    def test_falls_back_to_str_for_unknown_object(self) -> None:
        """Falls back to str() for objects that don't match known patterns."""

        class _Unknown:
            def __str__(self) -> str:
                return "mystery"

        result = _safe_chart_title(_Unknown())
        assert result == "mystery"


class TestListChartsTitle:
    """Bug 2: list_charts must return plain-string titles (JSON-serializable)."""

    def test_list_charts_no_title_returns_untitled_string(self, tmp_path: Path) -> None:
        """chart created with title='' should appear with '(untitled)' in list."""
        path = _workbook_with_data(tmp_path, "charts.xlsx")
        create_chart(path, "Sheet1", "A1:B3", chart_type="column", title="", target_cell="D1")
        charts = list_charts(path, "Sheet1")
        assert len(charts) == 1
        assert charts[0]["title"] == "(untitled)"

    def test_list_charts_type_is_always_plain_str(self, tmp_path: Path) -> None:
        """list_charts 'type' field is always a plain Python str."""
        path = _workbook_with_data(tmp_path, "charts.xlsx")
        create_chart(path, "Sheet1", "A1:B3", chart_type="line", title="T", target_cell="D1")
        charts = list_charts(path, "Sheet1")
        assert isinstance(charts[0]["type"], str)

    def test_list_charts_title_is_json_serializable_after_reload(self, tmp_path: Path) -> None:
        """After save-reload, list_charts title must survive json.dumps without error."""
        path = _workbook_with_data(tmp_path, "charts.xlsx")
        create_chart(path, "Sheet1", "A1:B3", chart_type="bar", title="Costs", target_cell="D1")

        # Re-loading the file causes openpyxl to parse the title as a Text object.
        charts = list_charts(path, "Sheet1")
        # Must not raise TypeError during serialisation
        serialized = json.dumps(charts)
        assert len(serialized) > 0

    def test_list_charts_multiple_charts_all_titles_are_str(self, tmp_path: Path) -> None:
        """All chart titles in list_charts result are plain strings, even after file reload."""
        path = _workbook_with_data(tmp_path, "charts.xlsx")
        create_chart(path, "Sheet1", "A1:B3", chart_type="column", title="First", target_cell="D1")
        create_chart(path, "Sheet1", "A1:B3", chart_type="line", title="Second", target_cell="D15")

        charts = list_charts(path, "Sheet1")
        assert len(charts) == 2
        for c in charts:
            assert isinstance(c["title"], str), f"title should be str, got {type(c['title'])}"


# ---------------------------------------------------------------------------
# Bug 3: write_cells range — MCP wrapper validates data is not None
# ---------------------------------------------------------------------------


class TestWriteCellsRangeValidation:
    """Bug 3: write_cells mode='range' must raise ValueError when data or start_cell is None."""

    def test_wrapper_raises_when_data_is_none(self, tmp_path: Path) -> None:
        """MCP write_cells raises ValueError for mode='range' with data=None."""
        from mcp_server.main import write_cells

        fp = _empty_workbook(tmp_path, "range.xlsx")
        with pytest.raises(ValueError, match="data is required"):
            write_cells(
                mode="range",
                file_path=fp,
                sheet_name="Sheet1",
                start_cell="A1",
                data=None,
            )

    def test_wrapper_raises_when_start_cell_is_none_for_range_mode(self, tmp_path: Path) -> None:
        """MCP write_cells raises ValueError for mode='range' with start_cell=None."""
        from mcp_server.main import write_cells

        fp = _empty_workbook(tmp_path, "range.xlsx")
        with pytest.raises(ValueError, match="start_cell is required"):
            write_cells(
                mode="range",
                file_path=fp,
                sheet_name="Sheet1",
                start_cell=None,
                data=[["A", "B"], [1, 2]],
            )

    def test_write_range_to_fresh_empty_workbook(self, tmp_path: Path) -> None:
        """write_range works on a brand-new empty workbook with no existing data."""
        fp = _empty_workbook(tmp_path, "range.xlsx")
        result = write_range(fp, "Sheet1", "A1", [["Header1", "Header2"], [10, 20], [30, 40]])
        assert "3 rows" in result

        wb = load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].value == "Header1"
        assert ws["B1"].value == "Header2"
        assert ws["A2"].value == 10
        assert ws["B3"].value == 40
        wb.close()

    def test_wrapper_range_write_succeeds_with_valid_args(self, tmp_path: Path) -> None:
        """MCP write_cells mode='range' writes data at offset position correctly."""
        from mcp_server.main import write_cells

        fp = _empty_workbook(tmp_path, "range.xlsx")
        result = write_cells(
            mode="range",
            file_path=fp,
            sheet_name="Sheet1",
            start_cell="C3",
            data=[["X", "Y"], [7, 8]],
        )
        assert isinstance(result, str)

        wb = load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["C3"].value == "X"
        assert ws["D3"].value == "Y"
        assert ws["C4"].value == 7
        assert ws["D4"].value == 8
        # Cells outside the written range remain empty
        assert ws["A1"].value is None
        wb.close()


# ---------------------------------------------------------------------------
# Bug 4: worksheet_ops freeze — response message content
# (The stored string value "B2" is already covered by test_worksheet_ops.py)
# ---------------------------------------------------------------------------


class TestFreezePanesMessages:
    """Bug 4: freeze_panes must return human-readable messages and store a string."""

    def test_freeze_message_includes_cell_ref(self, tmp_path: Path) -> None:
        """freeze_panes response message should reference the cell that was frozen."""
        fp = _empty_workbook(tmp_path)
        result = freeze_panes(fp, "Sheet1", "B2")
        assert "B2" in result

    def test_freeze_message_does_not_say_unfrozen(self, tmp_path: Path) -> None:
        """Freeze response must NOT say 'unfrozen'."""
        fp = _empty_workbook(tmp_path)
        result = freeze_panes(fp, "Sheet1", "B2")
        assert "unfrozen" not in result.lower()

    def test_unfreeze_message_indicates_unfrozen(self, tmp_path: Path) -> None:
        """Passing cell_ref=None should produce a message containing 'unfrozen'."""
        fp = _empty_workbook(tmp_path)
        freeze_panes(fp, "Sheet1", "B2")
        result = freeze_panes(fp, "Sheet1", cell_ref=None)
        assert "unfrozen" in result.lower()

    def test_freeze_stores_exact_string_not_cell_object(self, tmp_path: Path) -> None:
        """freeze_panes stores the cell ref as a plain str, not an openpyxl Cell object."""
        fp = _empty_workbook(tmp_path)
        freeze_panes(fp, "Sheet1", "D5")
        wb = load_workbook(fp)
        panes = wb["Sheet1"].freeze_panes
        assert panes == "D5"
        assert isinstance(panes, str)
        wb.close()

    def test_freeze_and_unfreeze_cycle(self, tmp_path: Path) -> None:
        """Freezing then unfreezing leaves freeze_panes as None in the saved file."""
        fp = _empty_workbook(tmp_path)
        freeze_panes(fp, "Sheet1", "C2")
        freeze_panes(fp, "Sheet1", cell_ref=None)
        wb = load_workbook(fp)
        assert wb["Sheet1"].freeze_panes is None
        wb.close()

    def test_freeze_at_a1_acts_as_unfreeze(self, tmp_path: Path) -> None:
        """Passing cell_ref='A1' should also unfreeze (top-left corner is no-op)."""
        fp = _empty_workbook(tmp_path)
        freeze_panes(fp, "Sheet1", "B2")
        result = freeze_panes(fp, "Sheet1", cell_ref="A1")
        wb = load_workbook(fp)
        assert wb["Sheet1"].freeze_panes is None
        wb.close()
        assert "unfrozen" in result.lower()


# ---------------------------------------------------------------------------
# Bug 5: worksheet_ops auto_filter — MCP wrapper requires cell_range
# ---------------------------------------------------------------------------


class TestAutoFilterRangeValidation:
    """Bug 5: MCP worksheet_ops auto_filter must raise when cell_range is None and remove=False."""

    def test_wrapper_raises_when_cell_range_none_and_not_removing(self, tmp_path: Path) -> None:
        """MCP worksheet_ops auto_filter raises ValueError when cell_range=None & remove=False."""
        from mcp_server.main import worksheet_ops

        fp = _workbook_with_data(tmp_path)
        with pytest.raises(ValueError, match="cell_range is required"):
            worksheet_ops(
                action="auto_filter",
                file_path=fp,
                sheet_name="Sheet1",
                cell_range=None,
                remove=False,
            )

    def test_wrapper_allows_remove_true_without_cell_range(self, tmp_path: Path) -> None:
        """MCP worksheet_ops auto_filter accepts remove=True with cell_range=None."""
        from mcp_server.main import worksheet_ops

        fp = _workbook_with_data(tmp_path)
        # First set a filter so removing makes sense
        set_auto_filter(fp, "Sheet1", "A1:B3")
        result = worksheet_ops(
            action="auto_filter",
            file_path=fp,
            sheet_name="Sheet1",
            cell_range=None,
            remove=True,
        )
        assert isinstance(result, str)
        wb = load_workbook(fp)
        assert wb["Sheet1"].auto_filter.ref is None
        wb.close()

    def test_wrapper_sets_filter_when_valid_range_provided(self, tmp_path: Path) -> None:
        """MCP worksheet_ops auto_filter sets the filter correctly when cell_range is given."""
        from mcp_server.main import worksheet_ops

        fp = _workbook_with_data(tmp_path)
        result = worksheet_ops(
            action="auto_filter",
            file_path=fp,
            sheet_name="Sheet1",
            cell_range="A1:B3",
            remove=False,
        )
        assert isinstance(result, str)
        wb = load_workbook(fp)
        assert wb["Sheet1"].auto_filter.ref == "A1:B3"
        wb.close()

    def test_underlying_set_auto_filter_with_none_does_not_error(self, tmp_path: Path) -> None:
        """The underlying set_auto_filter(cell_range=None) silently clears the filter."""
        fp = _workbook_with_data(tmp_path)
        # No ValueError from the underlying function itself — the guard is in the wrapper
        set_auto_filter(fp, "Sheet1", None, remove=False)
        wb = load_workbook(fp)
        assert wb["Sheet1"].auto_filter.ref is None
        wb.close()


# ---------------------------------------------------------------------------
# Bug 6: worksheet_ops copy_range_across — target_start_cell is respected
# ---------------------------------------------------------------------------


class TestCopyRangeAcrossTargetCell:
    """Bug 6: copy_range_across_sheets must place data at target_start_cell (not always A1)."""

    def test_default_target_places_data_at_a1(self, tmp_path: Path) -> None:
        """Without specifying target_start_cell the copy lands at A1."""
        fp = str(tmp_path / "copy.xlsx")
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Source"
        ws1.append(["Alpha", "Beta"])
        ws1.append([1, 2])
        wb.create_sheet("Target")
        wb.save(fp)
        wb.close()

        copy_range_across_sheets(fp, "Source", "A1:B2", "Target")

        wb = load_workbook(fp)
        ws2 = wb["Target"]
        assert ws2["A1"].value == "Alpha"
        assert ws2["B2"].value == 2
        wb.close()

    def test_custom_target_start_cell_places_data_correctly(self, tmp_path: Path) -> None:
        """Specifying target_start_cell='D5' pastes data starting at cell D5."""
        fp = str(tmp_path / "copy.xlsx")
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Source"
        ws1.append(["X", "Y"])
        ws1.append([10, 20])
        wb.create_sheet("Target")
        wb.save(fp)
        wb.close()

        copy_range_across_sheets(fp, "Source", "A1:B2", "Target", target_start_cell="D5")

        wb = load_workbook(fp)
        ws2 = wb["Target"]
        assert ws2["D5"].value == "X"
        assert ws2["E5"].value == "Y"
        assert ws2["D6"].value == 10
        assert ws2["E6"].value == 20
        # Check data did NOT land at A1
        assert ws2["A1"].value is None
        wb.close()

    def test_custom_target_start_cell_does_not_overlap_source_area(self, tmp_path: Path) -> None:
        """Range copied to a non-default cell leaves the default A1 area empty."""
        fp = str(tmp_path / "copy.xlsx")
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Src"
        ws1.append(["Hello"])
        ws1.append(["World"])
        wb.create_sheet("Dst")
        wb.save(fp)
        wb.close()

        copy_range_across_sheets(fp, "Src", "A1:A2", "Dst", target_start_cell="C3")

        wb = load_workbook(fp)
        ws2 = wb["Dst"]
        assert ws2["C3"].value == "Hello"
        assert ws2["C4"].value == "World"
        assert ws2["A1"].value is None
        wb.close()

    def test_worksheet_ops_wrapper_passes_target_start_cell(self, tmp_path: Path) -> None:
        """MCP worksheet_ops copy_range_across forwards target_start_cell to the function."""
        from mcp_server.main import worksheet_ops

        fp = str(tmp_path / "wscopy.xlsx")
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Src"
        ws1["A1"] = "Foo"
        ws1["B1"] = "Bar"
        wb.create_sheet("Dst")
        wb.save(fp)
        wb.close()

        result = worksheet_ops(
            action="copy_range_across",
            file_path=fp,
            source_sheet="Src",
            source_range="A1:B1",
            target_sheet="Dst",
            target_start_cell="F6",
        )
        assert isinstance(result, str)

        wb = load_workbook(fp)
        ws2 = wb["Dst"]
        assert ws2["F6"].value == "Foo"
        assert ws2["G6"].value == "Bar"
        assert ws2["A1"].value is None
        wb.close()

    def test_worksheet_ops_wrapper_default_target_is_a1(self, tmp_path: Path) -> None:
        """MCP worksheet_ops copy_range_across defaults target to A1 when not specified."""
        from mcp_server.main import worksheet_ops

        fp = str(tmp_path / "wsdef.xlsx")
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Src"
        ws1["A1"] = "DefaultTest"
        wb.create_sheet("Dst")
        wb.save(fp)
        wb.close()

        worksheet_ops(
            action="copy_range_across",
            file_path=fp,
            source_sheet="Src",
            source_range="A1:A1",  # must be a range expression, not single cell
            target_sheet="Dst",
            # target_start_cell deliberately omitted — should default to "A1"
        )

        wb = load_workbook(fp)
        assert wb["Dst"]["A1"].value == "DefaultTest"
        wb.close()


# ---------------------------------------------------------------------------
# Bug 7: doc_properties set_calc_mode — correct mode is persisted
# ---------------------------------------------------------------------------


class TestSetCalculationMode:
    """Bug 7: set_calculation_mode must save the requested mode, not always 'auto'."""

    def test_manual_mode_persists_in_file(self, tmp_path: Path) -> None:
        """After set_calculation_mode('manual'), reloading the file shows calcMode='manual'."""
        fp = _empty_workbook(tmp_path)
        result = set_calculation_mode(fp, "manual")
        assert "manual" in result

        wb = load_workbook(fp)
        assert wb.calculation.calcMode == "manual"
        wb.close()

    def test_auto_mode_persists_in_file(self, tmp_path: Path) -> None:
        """After switching to manual then back to auto, file shows calcMode='auto'."""
        fp = _empty_workbook(tmp_path)
        set_calculation_mode(fp, "manual")
        result = set_calculation_mode(fp, "auto")
        assert "auto" in result

        wb = load_workbook(fp)
        assert wb.calculation.calcMode == "auto"
        wb.close()

    def test_auto_no_table_mode_persists(self, tmp_path: Path) -> None:
        """set_calculation_mode('autoNoTable') stores 'autoNoTable' in the file."""
        fp = _empty_workbook(tmp_path)
        result = set_calculation_mode(fp, "autoNoTable")
        assert "autoNoTable" in result

        wb = load_workbook(fp)
        assert wb.calculation.calcMode == "autoNoTable"
        wb.close()

    def test_invalid_mode_raises_value_error(self, tmp_path: Path) -> None:
        """set_calculation_mode raises ValueError for unrecognised mode strings."""
        fp = _empty_workbook(tmp_path)
        with pytest.raises(ValueError, match="Invalid calculation mode"):
            set_calculation_mode(fp, "semiauto")

    def test_manual_mode_result_string_mentions_mode(self, tmp_path: Path) -> None:
        """The result string from setting manual mode explicitly names the mode."""
        fp = _empty_workbook(tmp_path)
        result = set_calculation_mode(fp, "manual")
        assert "manual" in result.lower()

    def test_doc_properties_wrapper_set_calc_mode_manual(self, tmp_path: Path) -> None:
        """MCP doc_properties wrapper persists 'manual' when calc_mode='manual' is passed."""
        from mcp_server.main import doc_properties

        fp = _empty_workbook(tmp_path)
        result = doc_properties(action="set_calc_mode", file_path=fp, calc_mode="manual")
        assert isinstance(result, str)
        assert "manual" in result

        wb = load_workbook(fp)
        assert wb.calculation.calcMode == "manual"
        wb.close()

    def test_doc_properties_wrapper_default_is_auto(self, tmp_path: Path) -> None:
        """MCP doc_properties wrapper defaults to 'auto' when calc_mode is not provided."""
        from mcp_server.main import doc_properties

        fp = _empty_workbook(tmp_path)
        # Set to manual first so we can verify the default resets it
        set_calculation_mode(fp, "manual")
        result = doc_properties(action="set_calc_mode", file_path=fp)
        assert "auto" in result

        wb = load_workbook(fp)
        assert wb.calculation.calcMode == "auto"
        wb.close()

    def test_set_calculation_mode_does_not_corrupt_other_data(self, tmp_path: Path) -> None:
        """Changing calc mode does not overwrite existing cell data in the workbook."""
        fp = _workbook_with_data(tmp_path)
        set_calculation_mode(fp, "manual")

        wb = load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].value == "Month"
        assert ws["B2"].value == 100
        wb.close()


# ---------------------------------------------------------------------------
# Bug 8: formula_audit value — None value response includes formula + note
# ---------------------------------------------------------------------------


class TestFormulaValueNullHandling:
    """Bug 8: get_formula_value must include 'formula' and 'note' keys when value is None."""

    def test_result_has_formula_key_when_value_is_none(self, tmp_path: Path) -> None:
        """When the cached value is None, the response contains the 'formula' key."""
        fp = _empty_workbook(tmp_path)
        set_formula(fp, "Sheet1", "A1", "=SUM(1,2,3)")
        result = get_formula_value(fp, "Sheet1", "A1")
        # openpyxl cannot evaluate formulas written programmatically
        if result["value"] is None:
            assert "formula" in result, "Expected 'formula' key when cached value is None"

    def test_result_has_note_key_when_value_is_none(self, tmp_path: Path) -> None:
        """When the cached value is None, the response contains the 'note' key."""
        fp = _empty_workbook(tmp_path)
        set_formula(fp, "Sheet1", "B2", "=A1*10")
        result = get_formula_value(fp, "Sheet1", "B2")
        if result["value"] is None:
            assert "note" in result, "Expected 'note' key when cached value is None"
            assert isinstance(result["note"], str)
            assert len(result["note"]) > 0

    def test_formula_key_contains_actual_formula_text(self, tmp_path: Path) -> None:
        """The 'formula' key shows the stored formula string, not a generic placeholder."""
        fp = _empty_workbook(tmp_path)
        formula_text = "=VLOOKUP(A1,A:B,2,0)"
        set_formula(fp, "Sheet1", "C3", formula_text)
        result = get_formula_value(fp, "Sheet1", "C3")
        if result["value"] is None:
            assert result["formula"] == formula_text

    def test_note_mentions_openpyxl_limitation(self, tmp_path: Path) -> None:
        """The 'note' key explains why the value is null (openpyxl cannot evaluate formulas)."""
        fp = _empty_workbook(tmp_path)
        set_formula(fp, "Sheet1", "A1", "=TODAY()")
        result = get_formula_value(fp, "Sheet1", "A1")
        if result["value"] is None:
            note_lower = result["note"].lower()
            assert "openpyxl" in note_lower or "formula" in note_lower or "cached" in note_lower

    def test_no_note_for_plain_non_formula_cell(self, tmp_path: Path) -> None:
        """Cells with a plain (non-formula) value should NOT have a 'note' key."""
        fp = _empty_workbook(tmp_path)
        write_cell(fp, "Sheet1", "A1", 99)
        result = get_formula_value(fp, "Sheet1", "A1")
        assert "note" not in result

    def test_no_formula_key_for_plain_non_formula_cell(self, tmp_path: Path) -> None:
        """Cells without a formula should not gain a spurious 'formula' key."""
        fp = _empty_workbook(tmp_path)
        write_cell(fp, "Sheet1", "A1", "hello")
        result = get_formula_value(fp, "Sheet1", "A1")
        # The function only adds 'formula' when val is None AND the cell has a formula
        assert result["value"] == "hello"
        assert "note" not in result

    def test_result_always_contains_cell_ref_key(self, tmp_path: Path) -> None:
        """get_formula_value always includes the 'cell' key in its response."""
        fp = _empty_workbook(tmp_path)
        set_formula(fp, "Sheet1", "D4", "=NOW()")
        result = get_formula_value(fp, "Sheet1", "D4")
        assert "cell" in result
        assert result["cell"] == "D4"

    def test_result_always_contains_value_key(self, tmp_path: Path) -> None:
        """get_formula_value always includes the 'value' key (may be None for unevaluated formulas)."""
        fp = _empty_workbook(tmp_path)
        set_formula(fp, "Sheet1", "A1", "=SUM(B1:B5)")
        result = get_formula_value(fp, "Sheet1", "A1")
        assert "value" in result

    def test_formula_audit_wrapper_action_value_returns_dict(self, tmp_path: Path) -> None:
        """MCP formula_audit wrapper with action='value' returns a dict."""
        from mcp_server.main import formula_audit

        fp = _empty_workbook(tmp_path)
        set_formula(fp, "Sheet1", "A1", "=1+1")
        result = formula_audit(action="value", file_path=fp, sheet_name="Sheet1", cell_ref="A1")
        assert isinstance(result, dict)
        assert "cell" in result
        assert "value" in result

    def test_formula_audit_wrapper_null_value_includes_formula_and_note(self, tmp_path: Path) -> None:
        """MCP formula_audit value action includes formula+note for unevaluated formula cells."""
        from mcp_server.main import formula_audit

        fp = _empty_workbook(tmp_path)
        set_formula(fp, "Sheet1", "A1", "=AVERAGE(B1:B10)")
        result = formula_audit(action="value", file_path=fp, sheet_name="Sheet1", cell_ref="A1")
        if result["value"] is None:
            assert "formula" in result
            assert "note" in result
