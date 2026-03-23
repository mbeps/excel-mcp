"""Conditional formatting operations for Excel worksheets."""

from __future__ import annotations

from logging import Logger

from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule, IconSetRule, Rule
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.differential import DifferentialStyle

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

VALID_FORMAT_TYPES = {"color_scale", "data_bar", "icon_set"}


def apply_conditional_formatting(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    format_type: str,
    start_color: str = "FF0000",
    mid_color: str = "FFFF00",
    end_color: str = "00FF00",
    bar_color: str = "FF638EC6",
    icon_style: str = "3Arrows",
) -> str:
    """Apply conditional formatting rule to a range."""
    if format_type not in VALID_FORMAT_TYPES:
        raise ValueError(f"Invalid format_type '{format_type}'. Must be one of: {VALID_FORMAT_TYPES}")

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        if format_type == "color_scale":
            rule = ColorScaleRule(
                start_type="percentile",
                start_value=10,
                start_color=start_color,
                mid_type="percentile",
                mid_value=50,
                mid_color=mid_color,
                end_type="percentile",
                end_value=90,
                end_color=end_color,
            )
        elif format_type == "data_bar":
            rule = DataBarRule(
                start_type="percentile",
                start_value=10,
                end_type="percentile",
                end_value=90,
                color=bar_color,
            )
        else:
            rule = IconSetRule(
                icon_style=icon_style,
                type="percent",
                values=[33, 67],
            )

        ws.conditional_formatting.add(cell_range, rule)
        save_workbook_safe(wb, file_path)
        logger.info("Applied %s conditional formatting to %s in %s", format_type, cell_range, file_path)
        return f"Applied '{format_type}' conditional formatting to '{cell_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def add_highlight_rule(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    operator: str,
    formula: str,
    font_color: str = "9C0006",
    bg_color: str = "FFC7CE",
) -> str:
    """Add a cell highlight conditional format rule."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        rule = CellIsRule(
            operator=operator,
            formula=[formula],
            font=Font(color=font_color),
            fill=PatternFill(bgColor=bg_color),
        )
        ws.conditional_formatting.add(cell_range, rule)
        save_workbook_safe(wb, file_path)
        logger.info("Added highlight rule (%s) to %s in %s", operator, cell_range, file_path)
        return f"Added highlight rule (operator='{operator}') to '{cell_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def remove_conditional_formatting(file_path: str, sheet_name: str, cell_range: str | None = None) -> str:
    """Remove conditional formatting from a sheet.

    If cell_range is None, removes all rules from the entire sheet.
    If cell_range is provided, removes only the rules applied to that specific range.
    """
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        if cell_range is None:
            ws.conditional_formatting._cf_rules.clear()  # pyright: ignore[reportAttributeAccessIssue]
            logger.info("Removed all conditional formatting from sheet '%s' in %s", sheet_name, file_path)
            msg = f"Removed all conditional formatting from sheet '{sheet_name}'."
        else:
            cf_rules = ws.conditional_formatting._cf_rules  # pyright: ignore[reportAttributeAccessIssue]
            key = cell_range.upper()
            matched = next((k for k in cf_rules if str(k.sqref).upper() == key), None)
            if matched is not None:
                del cf_rules[matched]
            logger.info(
                "Removed conditional formatting for range '%s' from sheet '%s' in %s", cell_range, sheet_name, file_path
            )
            msg = f"Removed conditional formatting for range '{cell_range}' from sheet '{sheet_name}'."
        save_workbook_safe(wb, file_path)
        return msg
    finally:
        wb.close()


def add_formula_rule(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    formula: str,
    font_color: str = "9C0006",
    bg_color: str = "FFC7CE",
) -> str:
    """Add a formula-based conditional formatting rule."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        rule = FormulaRule(
            formula=[formula],
            font=Font(color=font_color),
            fill=PatternFill(bgColor=bg_color),
        )
        ws.conditional_formatting.add(cell_range, rule)
        save_workbook_safe(wb, file_path)
        logger.info("Added formula rule to %s in %s", cell_range, file_path)
        return f"Added formula-based conditional formatting to '{cell_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def add_top_bottom_rule(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    rank: int = 10,
    bottom: bool = False,
    percent: bool = False,
    font_color: str = "9C0006",
    bg_color: str = "FFC7CE",
) -> str:
    """Add a top/bottom N conditional formatting rule."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        dxf = DifferentialStyle(
            font=Font(color=font_color),
            fill=PatternFill(bgColor=bg_color),
        )
        rule = Rule(
            type="top10",
            rank=rank,
            bottom=bottom,
            percent=percent,
            dxf=dxf,
        )
        ws.conditional_formatting.add(cell_range, rule)
        save_workbook_safe(wb, file_path)
        direction = "bottom" if bottom else "top"
        pct = "%" if percent else ""
        logger.info("Added %s %d%s rule to %s in %s", direction, rank, pct, cell_range, file_path)
        return f"Added {direction} {rank}{pct} conditional formatting to '{cell_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def add_above_average_rule(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    above: bool = True,
    font_color: str = "9C0006",
    bg_color: str = "FFC7CE",
) -> str:
    """Add an above/below average conditional formatting rule."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        dxf = DifferentialStyle(
            font=Font(color=font_color),
            fill=PatternFill(bgColor=bg_color),
        )
        rule = Rule(
            type="aboveAverage",
            aboveAverage=above,
            dxf=dxf,
        )
        ws.conditional_formatting.add(cell_range, rule)
        save_workbook_safe(wb, file_path)
        direction = "above" if above else "below"
        logger.info("Added %s average rule to %s in %s", direction, cell_range, file_path)
        return f"Added {direction}-average conditional formatting to '{cell_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def add_duplicate_rule(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    font_color: str = "9C0006",
    bg_color: str = "FFC7CE",
) -> str:
    """Add duplicate values highlighting conditional formatting rule."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        dxf = DifferentialStyle(
            font=Font(color=font_color),
            fill=PatternFill(bgColor=bg_color),
        )
        rule = Rule(
            type="duplicateValues",
            dxf=dxf,
        )
        ws.conditional_formatting.add(cell_range, rule)
        save_workbook_safe(wb, file_path)
        logger.info("Added duplicate values rule to %s in %s", cell_range, file_path)
        return f"Added duplicate-values conditional formatting to '{cell_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def list_conditional_formats(file_path: str, sheet_name: str) -> list[dict]:
    """List all conditional formatting rules on a sheet."""
    wb = load_workbook_safe(file_path, read_only=False)
    try:
        ws = get_sheet(wb, sheet_name)
        result: list[dict] = []
        for cf in ws.conditional_formatting:
            for rule in cf.rules:
                entry: dict = {
                    "range": str(cf.sqref),
                    "type": rule.type,
                    "priority": rule.priority,
                }
                if hasattr(rule, "formula") and rule.formula:
                    entry["formula"] = list(rule.formula)
                result.append(entry)
        return result
    finally:
        wb.close()
