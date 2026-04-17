"""WF-19 · Route-layer integration tests.

Tests the actual MCP tool dispatch functions (routes) rather than the
underlying tool functions directly.  Verifies dispatch, parameter
validation, and return-type correctness for consolidated/mode-based tools.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from mcp_server.routes.cell_ops import read_cells, write_cells
from mcp_server.routes.charts import chart
from mcp_server.routes.formulas import formula_audit, formula_write
from mcp_server.routes.governance import (
    conditional_format,
    data_validation,
    protection,
)
from mcp_server.routes.metadata import comment, named_range, table
from mcp_server.routes.worksheet_ops import (
    worksheet_print,
    worksheet_structure,
    worksheet_view,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_wb(tmp_path: Path, name: str = "rt.xlsx", rows: list[list] | None = None) -> str:
    """Create a small workbook and return its path."""
    path = str(tmp_path / name)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    if rows:
        for r in rows:
            ws.append(r)
    else:
        ws.append(["Name", "Value", "Category"])
        ws.append(["A", 10, "X"])
        ws.append(["B", 20, "Y"])
        ws.append(["C", 30, "X"])
        ws.append(["D", 40, "Y"])
        ws.append(["E", 50, "X"])
    wb.save(path)
    return path


# ===================================================================
# Worksheet view route
# ===================================================================


class TestWorksheetView:
    def test_freeze(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = worksheet_view(action="freeze", file_path=fp, sheet_name="Sheet1", cell_ref="B2")
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        assert wb["Sheet1"].freeze_panes == "B2"
        wb.close()

    def test_auto_filter(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = worksheet_view(
            action="auto_filter",
            file_path=fp,
            sheet_name="Sheet1",
            cell_range="A1:C6",
        )
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        assert wb["Sheet1"].auto_filter.ref is not None
        wb.close()

    def test_set_gridlines(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = worksheet_view(action="set_gridlines", file_path=fp, sheet_name="Sheet1", show=False)
        # BUG-03 fix: must return dict, not Pydantic model
        assert isinstance(result, (str, dict))


# ===================================================================
# Worksheet structure route
# ===================================================================


class TestWorksheetStructure:
    def test_insert_rows(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = worksheet_structure(action="insert_rows", file_path=fp, sheet_name="Sheet1", row=2, count=2)
        # BUG-01 fix: must return dict
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        # Original had 7 rows (header + 5 data); inserting 2 → 9
        assert wb["Sheet1"].max_row >= 8
        wb.close()

    def test_delete_rows(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = worksheet_structure(action="delete_rows", file_path=fp, sheet_name="Sheet1", row=3, count=1)
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        assert wb["Sheet1"].max_row <= 6
        wb.close()

    def test_set_row_height(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = worksheet_structure(
            action="set_row_height",
            file_path=fp,
            sheet_name="Sheet1",
            rows=[1, 2],
            height=30.0,
        )
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        assert wb["Sheet1"].row_dimensions[1].height == 30.0
        wb.close()

    def test_set_col_width(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = worksheet_structure(
            action="set_col_width",
            file_path=fp,
            sheet_name="Sheet1",
            cols_list=["A", "B"],
            width=25.0,
        )
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        assert wb["Sheet1"].column_dimensions["A"].width == 25.0
        wb.close()

    def test_group_rows(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = worksheet_structure(
            action="group_rows",
            file_path=fp,
            sheet_name="Sheet1",
            start_row=2,
            end_row=4,
        )
        assert isinstance(result, (str, dict))


# ===================================================================
# Worksheet print route
# ===================================================================


class TestWorksheetPrint:
    def test_set_print_area(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = worksheet_print(
            action="set_print_area",
            file_path=fp,
            sheet_name="Sheet1",
            print_area="A1:C6",
        )
        # BUG-02 fix: must return dict
        assert isinstance(result, (str, dict))

    def test_set_page_setup(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = worksheet_print(
            action="set_page_setup",
            file_path=fp,
            sheet_name="Sheet1",
            orientation="landscape",
        )
        assert isinstance(result, (str, dict))

    def test_set_print_titles(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = worksheet_print(
            action="set_print_titles",
            file_path=fp,
            sheet_name="Sheet1",
            title_rows="1:1",
        )
        assert isinstance(result, (str, dict))


# ===================================================================
# Chart route
# ===================================================================


class TestChartRoute:
    @staticmethod
    def _chart_wb(tmp_path: Path) -> str:
        fp = str(tmp_path / "chart.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["Month", "Revenue", "Cost"])
        for i, m in enumerate(["Jan", "Feb", "Mar", "Apr"], 1):
            ws.append([m, i * 1000, i * 600])
        wb.save(fp)
        return fp

    def test_create(self, tmp_path: Path) -> None:
        fp = self._chart_wb(tmp_path)
        result = chart(
            action="create",
            file_path=fp,
            sheet_name="Sheet1",
            data_range="A1:C5",
            chart_type="column",
            title="Revenue",
        )
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        assert len(wb["Sheet1"]._charts) >= 1
        wb.close()

    def test_data_labels(self, tmp_path: Path) -> None:
        fp = self._chart_wb(tmp_path)
        chart(
            action="create", file_path=fp, sheet_name="Sheet1", data_range="A1:C5", chart_type="column", title="Sales"
        )
        # BUG-04 fix: data_labels must return dict
        result = chart(
            action="data_labels",
            file_path=fp,
            sheet_name="Sheet1",
            chart_title="Sales",
            show_value=True,
        )
        assert isinstance(result, (str, dict))

    def test_legend(self, tmp_path: Path) -> None:
        fp = self._chart_wb(tmp_path)
        chart(
            action="create",
            file_path=fp,
            sheet_name="Sheet1",
            data_range="A1:C5",
            chart_type="column",
            title="Legend Test",
        )
        # BUG-05 fix: legend must return dict
        result = chart(
            action="legend",
            file_path=fp,
            sheet_name="Sheet1",
            chart_title="Legend Test",
            show_legend=True,
            legend_position="bottom",
        )
        assert isinstance(result, (str, dict))

    def test_delete_by_title(self, tmp_path: Path) -> None:
        fp = self._chart_wb(tmp_path)
        chart(
            action="create",
            file_path=fp,
            sheet_name="Sheet1",
            data_range="A1:C5",
            chart_type="column",
            title="ToDelete",
        )
        # BUG-09 fix: delete by chart_title
        result = chart(
            action="delete",
            file_path=fp,
            sheet_name="Sheet1",
            chart_title="ToDelete",
        )
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        assert len(wb["Sheet1"]._charts) == 0
        wb.close()


# ===================================================================
# Read / write cells (mode-based dispatch)
# ===================================================================


class TestCellOpsRoute:
    def test_read_cells_single(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = read_cells(mode="single", file_path=fp, sheet_name="Sheet1", cell_ref="A1")
        assert isinstance(result, dict)
        assert result.get("value") == "Name"

    def test_read_cells_range(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = read_cells(mode="range", file_path=fp, sheet_name="Sheet1", start_cell="A1", end_cell="C2")
        assert isinstance(result, dict)

    def test_write_cells_single(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = write_cells(mode="single", file_path=fp, sheet_name="Sheet1", cell_ref="D1", value="New")
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        assert wb["Sheet1"]["D1"].value == "New"
        wb.close()

    def test_write_cells_range(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = write_cells(
            mode="range",
            file_path=fp,
            sheet_name="Sheet1",
            start_cell="E1",
            data=[["X", "Y"], [1, 2]],
        )
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        assert wb["Sheet1"]["E1"].value == "X"
        assert wb["Sheet1"]["F2"].value == 2
        wb.close()


# ===================================================================
# Formula routes
# ===================================================================


class TestFormulaRoute:
    def test_formula_write_set(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = formula_write(
            action="set",
            file_path=fp,
            sheet_name="Sheet1",
            cell_ref="D2",
            formula="=B2+B3",
        )
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        assert wb["Sheet1"]["D2"].value == "=B2+B3"
        wb.close()

    def test_formula_write_batch(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = formula_write(
            action="batch",
            file_path=fp,
            sheet_name="Sheet1",
            formulas={"D2": "=B2*2", "D3": "=B3*2"},
        )
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        assert wb["Sheet1"]["D2"].value == "=B2*2"
        assert wb["Sheet1"]["D3"].value == "=B3*2"
        wb.close()

    def test_formula_audit_value(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        # Write a formula first, then audit
        formula_write(action="set", file_path=fp, sheet_name="Sheet1", cell_ref="D2", formula="=B2+B3")
        result = formula_audit(action="value", file_path=fp, sheet_name="Sheet1", cell_ref="D2")
        assert isinstance(result, dict)

    def test_formula_audit_errors(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = formula_audit(action="errors", file_path=fp, sheet_name="Sheet1")
        assert isinstance(result, dict)
        assert "errors" in result

    def test_formula_audit_precedents(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        formula_write(action="set", file_path=fp, sheet_name="Sheet1", cell_ref="D2", formula="=B2+B3")
        result = formula_audit(action="precedents", file_path=fp, sheet_name="Sheet1", cell_ref="D2")
        assert isinstance(result, dict)
        # Precedents should include B2 and B3
        assert "B2" in str(result["precedents"])


# ===================================================================
# Table route
# ===================================================================


class TestTableRoute:
    def test_create(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = table(
            action="create",
            file_path=fp,
            sheet_name="Sheet1",
            table_name="Sales",
            data_range="A1:C6",
        )
        assert isinstance(result, (str, dict))

    def test_list(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        table(action="create", file_path=fp, sheet_name="Sheet1", table_name="T1", data_range="A1:C6")
        result = table(action="list", file_path=fp, sheet_name="Sheet1")
        assert isinstance(result, list)
        assert len(result) >= 1


# ===================================================================
# Conditional format route
# ===================================================================


class TestConditionalFormatRoute:
    def test_apply(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = conditional_format(
            action="apply",
            file_path=fp,
            sheet_name="Sheet1",
            cell_range="B2:B6",
            format_type="color_scale",
        )
        assert isinstance(result, (str, dict))


# ===================================================================
# Protection route
# ===================================================================


class TestProtectionRoute:
    def test_protect_sheet(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = protection(
            action="protect_sheet",
            file_path=fp,
            sheet_name="Sheet1",
        )
        assert isinstance(result, str)
        wb = load_workbook(fp)
        assert wb["Sheet1"].protection.sheet is True
        wb.close()


# ===================================================================
# Data validation route
# ===================================================================


class TestDataValidationRoute:
    def test_dropdown(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        result = data_validation(
            action="dropdown",
            file_path=fp,
            sheet_name="Sheet1",
            cell_range="D2:D6",
            options=["Yes", "No", "Maybe"],
        )
        assert isinstance(result, (str, dict))
        wb = load_workbook(fp)
        validations = wb["Sheet1"].data_validations.dataValidation
        assert len(validations) >= 1
        wb.close()


# ===================================================================
# Comment route
# ===================================================================


class TestCommentRoute:
    def test_add_and_read(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        add_result = comment(
            action="add",
            file_path=fp,
            sheet_name="Sheet1",
            cell_ref="A1",
            text="Test comment",
            author="Tester",
        )
        assert isinstance(add_result, str)

        read_result = comment(
            action="read",
            file_path=fp,
            sheet_name="Sheet1",
            cell_ref="A1",
        )
        # Should be CommentInfo or dict-like
        assert read_result is not None
        if hasattr(read_result, "text"):
            assert "Test comment" in read_result.text
        elif isinstance(read_result, dict):
            assert "Test comment" in str(read_result)


# ===================================================================
# Named range route
# ===================================================================


class TestNamedRangeRoute:
    def test_create_and_list(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        create_result = named_range(
            action="create",
            file_path=fp,
            name="SalesData",
            destination="Sheet1!A1:C6",
        )
        assert isinstance(create_result, str)

        list_result = named_range(action="list", file_path=fp)
        assert isinstance(list_result, list)
        names = [r.name if hasattr(r, "name") else r.get("name") for r in list_result]
        assert "SalesData" in names


# ===================================================================
# Parameter validation — missing required params raise ValueError
# ===================================================================


class TestRouteValidationErrors:
    def test_read_cells_single_missing_cell_ref(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        with pytest.raises(ValueError, match="cell_ref"):
            read_cells(mode="single", file_path=fp, sheet_name="Sheet1")

    def test_write_cells_range_missing_data(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        with pytest.raises(ValueError, match="data"):
            write_cells(mode="range", file_path=fp, sheet_name="Sheet1", start_cell="A1")

    def test_chart_create_missing_data_range(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        with pytest.raises(ValueError, match="data_range"):
            chart(action="create", file_path=fp, sheet_name="Sheet1")

    def test_formula_write_set_missing_formula(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        with pytest.raises(ValueError, match="formula"):
            formula_write(action="set", file_path=fp, sheet_name="Sheet1", cell_ref="A1")

    def test_unknown_action_raises(self, tmp_path: Path) -> None:
        fp = _make_wb(tmp_path)
        with pytest.raises(ValueError, match="Unknown"):
            worksheet_view(action="bogus", file_path=fp, sheet_name="Sheet1")  # type: ignore[arg-type]
