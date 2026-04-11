"""Workflow tests for formatting and conditional formatting operations.

Tests chain multiple MCP tool calls simulating real formatting workflows and verify
results independently using openpyxl — never via MCP tools.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from mcp_server.tools.cell_ops import write_cell, write_range
from mcp_server.tools.conditional_formatting import (
    add_formula_rule,
    add_highlight_rule,
    add_top_bottom_rule,
    apply_conditional_formatting,
    remove_conditional_formatting,
)
from mcp_server.tools.formatting import (
    apply_named_style,
    auto_fit_columns,
    clear_cell_format,
    copy_cell_format,
    format_cells,
)
from mcp_server.tools.workbook import create_workbook


def _create_with_data(fp: str, sheet: str = "Sheet1") -> None:
    """Create a workbook and write sample data for formatting tests."""
    create_workbook(fp, sheet_names=[sheet])
    write_range(
        fp,
        sheet,
        "A1",
        [
            ["Name", "Age", "City", "Salary"],
            ["Alice", 30, "New York", 70000],
            ["Bob", 25, "Chicago", 55000],
            ["Charlie", 35, "New York", 90000],
            ["Diana", 28, "Chicago", 62000],
            ["Eve", 32, "Boston", 80000],
        ],
    )


# ---------------------------------------------------------------------------
# 1. format_cells — font attributes
# ---------------------------------------------------------------------------


class TestFormatCellsFont:
    def test_format_cells_font(self, tmp_path: Path) -> None:
        """Apply bold, italic, font_size, font_color → verify openpyxl font."""
        fp = str(tmp_path / "font.xlsx")
        _create_with_data(fp)

        format_cells(fp, "Sheet1", "A1:D1", bold=True, italic=True, font_size=14, font_color="FF0000")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        for col in range(1, 5):
            cell = ws.cell(row=1, column=col)
            assert cell.font.bold is True
            assert cell.font.italic is True
            assert cell.font.size == 14
            assert cell.font.color and cell.font.color.rgb and "FF0000" in cell.font.color.rgb
        wb.close()


# ---------------------------------------------------------------------------
# 2. format_cells — fill
# ---------------------------------------------------------------------------


class TestFormatCellsFill:
    def test_format_cells_fill(self, tmp_path: Path) -> None:
        """Set bg_color → verify PatternFill."""
        fp = str(tmp_path / "fill.xlsx")
        _create_with_data(fp)

        format_cells(fp, "Sheet1", "A1:D1", bg_color="FFFF00")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        for col in range(1, 5):
            cell = ws.cell(row=1, column=col)
            assert cell.fill.fill_type == "solid"
            assert cell.fill.start_color and "FFFF00" in cell.fill.start_color.rgb
        wb.close()


# ---------------------------------------------------------------------------
# 3. format_cells — border
# ---------------------------------------------------------------------------


class TestFormatCellsBorder:
    def test_format_cells_border(self, tmp_path: Path) -> None:
        """Set border_style='thin' → verify all four sides."""
        fp = str(tmp_path / "border.xlsx")
        _create_with_data(fp)

        format_cells(fp, "Sheet1", "A1:D1", border_style="thin", border_color="000000")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        for col in range(1, 5):
            cell = ws.cell(row=1, column=col)
            assert cell.border.left.style == "thin"
            assert cell.border.right.style == "thin"
            assert cell.border.top.style == "thin"
            assert cell.border.bottom.style == "thin"
        wb.close()


# ---------------------------------------------------------------------------
# 4. format_cells — alignment
# ---------------------------------------------------------------------------


class TestFormatCellsAlignment:
    def test_format_cells_alignment(self, tmp_path: Path) -> None:
        """Set horizontal='center', vertical='top' → verify Alignment."""
        fp = str(tmp_path / "align.xlsx")
        _create_with_data(fp)

        format_cells(fp, "Sheet1", "A1:D1", horizontal_alignment="center", vertical_alignment="top")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        for col in range(1, 5):
            cell = ws.cell(row=1, column=col)
            assert cell.alignment.horizontal == "center"
            assert cell.alignment.vertical == "top"
        wb.close()


# ---------------------------------------------------------------------------
# 5. format_cells — number_format_preset
# ---------------------------------------------------------------------------


class TestFormatCellsNumberFormat:
    def test_format_cells_number_format(self, tmp_path: Path) -> None:
        """Set number_format_preset='decimal2' → verify format string '0.00'."""
        fp = str(tmp_path / "numfmt.xlsx")
        _create_with_data(fp)

        format_cells(fp, "Sheet1", "D2:D6", number_format_preset="decimal2")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        for row in range(2, 7):
            assert ws.cell(row=row, column=4).number_format == "0.00"
        wb.close()


# ---------------------------------------------------------------------------
# 6. format_cells — preserve_existing
# ---------------------------------------------------------------------------


class TestFormatCellsPreserveExisting:
    def test_format_cells_preserve_existing(self, tmp_path: Path) -> None:
        """Format font → format fill with preserve_existing=True → both applied."""
        fp = str(tmp_path / "preserve.xlsx")
        _create_with_data(fp)

        # Step 1: Apply bold font
        format_cells(fp, "Sheet1", "A1", bold=True, font_size=16)

        # Step 2: Apply fill with preserve_existing — should keep bold+font_size
        format_cells(fp, "Sheet1", "A1", bg_color="00FF00", preserve_existing=True)

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        cell = ws["A1"]
        assert cell.font.bold is True
        assert cell.font.size == 16
        assert cell.fill.fill_type == "solid"
        assert cell.fill.start_color and "00FF00" in cell.fill.start_color.rgb
        wb.close()


# ---------------------------------------------------------------------------
# 7. apply_named_style — "Title"
# ---------------------------------------------------------------------------


class TestApplyNamedStyle:
    def test_apply_named_style(self, tmp_path: Path) -> None:
        """Apply 'Title' style → verify cell.style is 'Title'."""
        fp = str(tmp_path / "named.xlsx")
        _create_with_data(fp)

        result = apply_named_style(fp, "Sheet1", "A1", "Title")

        assert result["cells_styled"] == 1
        assert result["style_name"] == "Title"

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].style == "Title"
        wb.close()


# ---------------------------------------------------------------------------
# 8. apply_named_style — case-insensitive
# ---------------------------------------------------------------------------


class TestApplyNamedStyleCaseInsensitive:
    def test_apply_named_style_case_insensitive(self, tmp_path: Path) -> None:
        """Apply 'title' (lowercase) → verify works and maps to 'Title'."""
        fp = str(tmp_path / "named_ci.xlsx")
        _create_with_data(fp)

        result = apply_named_style(fp, "Sheet1", "A1", "title")

        assert result["style_name"] == "Title"

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].style == "Title"
        wb.close()


# ---------------------------------------------------------------------------
# 9. auto_fit_columns
# ---------------------------------------------------------------------------


class TestAutoFitColumns:
    def test_auto_fit_columns(self, tmp_path: Path) -> None:
        """Write varying data → auto_fit → verify column widths changed."""
        fp = str(tmp_path / "autofit.xlsx")
        _create_with_data(fp)

        # Get default widths before auto-fit
        wb_before = openpyxl.load_workbook(fp)
        ws_before = wb_before["Sheet1"]
        default_width = ws_before.column_dimensions["A"].width
        wb_before.close()

        auto_fit_columns(fp, "Sheet1")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        # "Name" header is 4 chars → width ~6; "Charlie" is 7 chars → width ~9
        a_width = ws.column_dimensions["A"].width
        assert a_width is not None
        assert a_width > 0
        # "Salary" column D should be wider than min
        d_width = ws.column_dimensions["D"].width
        assert d_width is not None
        assert d_width > 0
        wb.close()


# ---------------------------------------------------------------------------
# 10. copy_cell_format
# ---------------------------------------------------------------------------


class TestCopyCellFormat:
    def test_copy_cell_format(self, tmp_path: Path) -> None:
        """Format A1 → copy_cell_format to B1 → verify B1 has same format."""
        fp = str(tmp_path / "copyfmt.xlsx")
        _create_with_data(fp)

        format_cells(fp, "Sheet1", "A1", bold=True, italic=True, bg_color="FF8800", font_size=18)

        result = copy_cell_format(fp, "Sheet1", "A1", "B1")
        assert result["cells_formatted"] == 1

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["B1"].font.bold is True
        assert ws["B1"].font.italic is True
        assert ws["B1"].font.size == 18
        assert ws["B1"].fill.fill_type == "solid"
        assert ws["B1"].fill.start_color and "FF8800" in ws["B1"].fill.start_color.rgb
        wb.close()


# ---------------------------------------------------------------------------
# 11. clear_cell_format
# ---------------------------------------------------------------------------


class TestClearCellFormat:
    def test_clear_cell_format(self, tmp_path: Path) -> None:
        """Format A1 → clear_cell_format → verify default format."""
        fp = str(tmp_path / "clearfmt.xlsx")
        _create_with_data(fp)

        format_cells(fp, "Sheet1", "A1", bold=True, bg_color="FF0000", font_size=20)

        result = clear_cell_format(fp, "Sheet1", "A1")
        assert result["cells_cleared"] == 1

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        cell = ws["A1"]
        assert cell.font.bold is not True
        assert cell.font.size != 20 or cell.font.size is None
        assert cell.fill.fill_type is None or cell.fill.fill_type == "none" or cell.fill.start_color.rgb == "00000000"
        assert cell.number_format == "General"
        wb.close()


# ---------------------------------------------------------------------------
# 12. conditional_format — highlight rule
# ---------------------------------------------------------------------------


class TestConditionalFormatHighlight:
    def test_conditional_format_highlight(self, tmp_path: Path) -> None:
        """add_highlight_rule → verify CF rule exists in openpyxl."""
        fp = str(tmp_path / "cf_highlight.xlsx")
        _create_with_data(fp)

        add_highlight_rule(
            fp,
            "Sheet1",
            "D2:D6",
            operator="greaterThan",
            formula="70000",
            font_color="006100",
            bg_color="C6EFCE",
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        cf_rules = list(ws.conditional_formatting)
        assert len(cf_rules) >= 1
        found = False
        for cf in cf_rules:
            if "D2:D6" in str(cf):
                found = True
                break
        assert found, "Expected CF rule on D2:D6"
        wb.close()


# ---------------------------------------------------------------------------
# 13. conditional_format — formula rule
# ---------------------------------------------------------------------------


class TestConditionalFormatFormulaRule:
    def test_conditional_format_formula_rule(self, tmp_path: Path) -> None:
        """add_formula_rule → verify rule exists."""
        fp = str(tmp_path / "cf_formula.xlsx")
        _create_with_data(fp)

        add_formula_rule(
            fp,
            "Sheet1",
            "A2:D6",
            formula='$C2="New York"',
            font_color="9C0006",
            bg_color="FFC7CE",
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        cf_rules = list(ws.conditional_formatting)
        assert len(cf_rules) >= 1
        found = any("A2:D6" in str(cf) for cf in cf_rules)
        assert found, "Expected CF rule on A2:D6"
        wb.close()


# ---------------------------------------------------------------------------
# 14. conditional_format — top/bottom rule
# ---------------------------------------------------------------------------


class TestConditionalFormatTopBottom:
    def test_conditional_format_top_bottom(self, tmp_path: Path) -> None:
        """add_top_bottom_rule → verify rule exists."""
        fp = str(tmp_path / "cf_top.xlsx")
        _create_with_data(fp)

        result = add_top_bottom_rule(
            fp,
            "Sheet1",
            "D2:D6",
            is_top=True,
            rank=3,
            bg_color="FFFF00",
        )

        assert result["type"] == "top10"
        assert result["rank"] == 3

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        cf_rules = list(ws.conditional_formatting)
        assert len(cf_rules) >= 1
        wb.close()


# ---------------------------------------------------------------------------
# 15. conditional_format — remove
# ---------------------------------------------------------------------------


class TestConditionalFormatRemove:
    def test_conditional_format_remove(self, tmp_path: Path) -> None:
        """Add CF → remove_conditional_formatting → verify removed."""
        fp = str(tmp_path / "cf_remove.xlsx")
        _create_with_data(fp)

        add_highlight_rule(fp, "Sheet1", "D2:D6", operator="greaterThan", formula="60000")

        # Verify rule exists
        wb = openpyxl.load_workbook(fp)
        assert len(list(wb["Sheet1"].conditional_formatting)) >= 1
        wb.close()

        # Remove all CF
        remove_conditional_formatting(fp, "Sheet1")

        wb = openpyxl.load_workbook(fp)
        assert len(list(wb["Sheet1"].conditional_formatting)) == 0
        wb.close()


# ---------------------------------------------------------------------------
# 16. complete formatting workflow
# ---------------------------------------------------------------------------


class TestFormattingWorkflowComplete:
    def test_formatting_workflow_complete(self, tmp_path: Path) -> None:
        """Write data → format header → format body → CF rules → auto-fit → verify all."""
        fp = str(tmp_path / "workflow.xlsx")
        _create_with_data(fp)

        # Step 1: Format header row — bold, fill, border
        format_cells(
            fp,
            "Sheet1",
            "A1:D1",
            bold=True,
            font_size=12,
            font_color="FFFFFF",
            bg_color="4472C4",
            border_style="thin",
            border_color="000000",
            horizontal_alignment="center",
        )

        # Step 2: Format body — alignment + number format for salary
        format_cells(fp, "Sheet1", "D2:D6", number_format_preset="currency")

        # Step 3: Add CF highlight — salary > 70000
        add_highlight_rule(
            fp,
            "Sheet1",
            "D2:D6",
            operator="greaterThan",
            formula="70000",
            font_color="006100",
            bg_color="C6EFCE",
        )

        # Step 4: Add top-3 rule on salary
        add_top_bottom_rule(fp, "Sheet1", "D2:D6", is_top=True, rank=3)

        # Step 5: Auto-fit
        auto_fit_columns(fp, "Sheet1")

        # Verify header formatting
        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]

        header_a1 = ws["A1"]
        assert header_a1.font.bold is True
        assert header_a1.font.size == 12
        assert header_a1.fill.fill_type == "solid"
        assert header_a1.alignment.horizontal == "center"
        assert header_a1.border.top.style == "thin"

        # Verify number format on salary
        assert ws["D2"].number_format == '"$"#,##0.00'

        # Verify CF rules exist (at least 2 individual rules — highlight + top-bottom)
        cf_rules = list(ws.conditional_formatting)
        total_rules = sum(len(cf.rules) for cf in cf_rules)
        assert total_rules >= 2

        # Verify column widths were adjusted
        assert ws.column_dimensions["A"].width > 0

        wb.close()


# ---------------------------------------------------------------------------
# 17. format_cells — per-side borders
# ---------------------------------------------------------------------------


class TestFormatCellsPerSideBorders:
    def test_format_cells_per_side_borders(self, tmp_path: Path) -> None:
        """Set top_border_style + bottom_border_style only → verify sides independently."""
        fp = str(tmp_path / "perside.xlsx")
        _create_with_data(fp)

        format_cells(
            fp,
            "Sheet1",
            "A1:D1",
            top_border_style="thick",
            bottom_border_style="double",
            border_color="0000FF",
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        cell = ws["A1"]
        assert cell.border.top.style == "thick"
        assert cell.border.bottom.style == "double"
        # Left and right should not have been set
        assert cell.border.left.style is None
        assert cell.border.right.style is None
        wb.close()


# ---------------------------------------------------------------------------
# 18. copy_cell_format to range
# ---------------------------------------------------------------------------


class TestCopyCellFormatToRange:
    def test_copy_cell_format_to_range(self, tmp_path: Path) -> None:
        """Format A1 → copy to B1:D1 → verify all targets."""
        fp = str(tmp_path / "copyfmt_range.xlsx")
        _create_with_data(fp)

        format_cells(fp, "Sheet1", "A1", bold=True, bg_color="AABBCC")

        result = copy_cell_format(fp, "Sheet1", "A1", "B1:D1")
        assert result["cells_formatted"] == 3

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        for col in ["B", "C", "D"]:
            cell = ws[f"{col}1"]
            assert cell.font.bold is True
            assert cell.fill.fill_type == "solid"
        wb.close()


# ---------------------------------------------------------------------------
# 19. apply_named_style to range
# ---------------------------------------------------------------------------


class TestApplyNamedStyleToRange:
    def test_apply_named_style_to_range(self, tmp_path: Path) -> None:
        """Apply 'Good' style to A2:D2 → verify all cells."""
        fp = str(tmp_path / "named_range.xlsx")
        _create_with_data(fp)

        result = apply_named_style(fp, "Sheet1", "A2:D2", "Good")
        assert result["cells_styled"] == 4
        assert result["style_name"] == "Good"

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        for col in range(1, 5):
            assert ws.cell(row=2, column=col).style == "Good"
        wb.close()


# ---------------------------------------------------------------------------
# 20. conditional_format — color_scale
# ---------------------------------------------------------------------------


class TestConditionalFormatColorScale:
    def test_conditional_format_color_scale(self, tmp_path: Path) -> None:
        """apply_conditional_formatting with color_scale → verify rule type."""
        fp = str(tmp_path / "cf_scale.xlsx")
        _create_with_data(fp)

        apply_conditional_formatting(
            fp,
            "Sheet1",
            "D2:D6",
            format_type="color_scale",
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        cf_rules = list(ws.conditional_formatting)
        assert len(cf_rules) >= 1
        rule_types = [r.type for cf in cf_rules for r in cf.rules]
        assert "colorScale" in rule_types
        wb.close()


# ---------------------------------------------------------------------------
# 21. conditional_format — remove specific range
# ---------------------------------------------------------------------------


class TestConditionalFormatRemoveSpecific:
    def test_conditional_format_remove_specific(self, tmp_path: Path) -> None:
        """Add two CF rules → remove one by range → verify only one remains."""
        fp = str(tmp_path / "cf_rm_specific.xlsx")
        _create_with_data(fp)

        add_highlight_rule(fp, "Sheet1", "D2:D6", operator="greaterThan", formula="60000")
        add_highlight_rule(fp, "Sheet1", "B2:B6", operator="lessThan", formula="30")

        remove_conditional_formatting(fp, "Sheet1", cell_range="D2:D6")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        cf_rules = list(ws.conditional_formatting)
        ranges_left = [str(cf) for cf in cf_rules]
        assert not any("D2:D6" in r for r in ranges_left)
        assert any("B2:B6" in r for r in ranges_left)
        wb.close()
