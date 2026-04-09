from __future__ import annotations

from typing import Literal

import mcp_server.tools.protection as _protection
import mcp_server.tools.data_validation as _data_val
import mcp_server.tools.doc_properties as _doc_props
import mcp_server.tools.conditional_formatting as _cond_fmt
from mcp_server.models.common import ValidationOperator

__all__ = [
    "protection",
    "data_validation",
    "doc_properties",
    "conditional_format",
]


def protection(
    action: Literal["protect_sheet", "unprotect_sheet", "protect_cells", "protect_workbook", "unprotect_workbook"],
    file_path: str,
    sheet_name: str | None = None,
    password: str | None = None,
    locked_range: str | None = None,
    unlocked_ranges: list[str] | None = None,
    allow_formatting_cells: bool = False,
    allow_formatting_columns: bool = False,
    allow_formatting_rows: bool = False,
    allow_insert_columns: bool = False,
    allow_insert_rows: bool = False,
    allow_delete_columns: bool = False,
    allow_delete_rows: bool = False,
    allow_sort: bool = False,
    allow_filter: bool = False,
    lock_structure: bool = True,
    lock_windows: bool = False,
) -> str:
    """Protection operations for sheets, cells, and workbooks.

    action="protect_sheet": Enable sheet protection. Requires: sheet_name. Optional: password, allow_* flags.
    action="unprotect_sheet": DESTRUCTIVE. Remove sheet protection. Requires: sheet_name.
    action="protect_cells": Lock cells. Requires: sheet_name, locked_range. Optional: unlocked_ranges.
    action="protect_workbook": Protect workbook structure. Optional: password, lock_structure, lock_windows.
    action="unprotect_workbook": DESTRUCTIVE. Remove workbook protection.
    """
    if action == "protect_sheet":
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='protect_sheet'.")
        return _protection.protect_sheet(
            file_path,
            sheet_name,
            password,
            allow_formatting_cells,
            allow_formatting_columns,
            allow_formatting_rows,
            allow_insert_columns,
            allow_insert_rows,
            allow_delete_columns,
            allow_delete_rows,
            allow_sort,
            allow_filter,
        )
    if action == "unprotect_sheet":
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='unprotect_sheet'.")
        return _protection.unprotect_sheet(file_path, sheet_name, password)
    if action == "protect_cells":
        if not locked_range:
            raise ValueError("locked_range is required for action='protect_cells'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='protect_cells'.")
        return _protection.protect_cells(file_path, sheet_name, locked_range, unlocked_ranges)
    if action == "protect_workbook":
        return _doc_props.protect_workbook(file_path, password, lock_structure, lock_windows)
    if action == "unprotect_workbook":
        return _doc_props.unprotect_workbook(file_path)
    raise ValueError(f"Unknown action: {action}")


def data_validation(
    action: Literal["dropdown", "numeric", "date", "remove", "formula"],
    file_path: str,
    sheet_name: str,
    cell_range: str,
    options: list[str] | None = None,
    source_range: str | None = None,
    operator: ValidationOperator | None = None,
    value1: float | None = None,
    value2: float | None = None,
    date1: str = "",
    date2: str | None = None,
    allow_blank: bool = True,
    error_style: str = "stop",
    error_title: str | None = None,
    error_message: str | None = None,
    prompt_title: str | None = None,
    prompt_message: str | None = None,
    formula: str | None = None,
    show_error: bool = True,
) -> str | dict:
    """Data validation operations.

    action="dropdown": Add dropdown list. Requires: options or source_range.
    action="numeric": Numeric validation. Requires: operator, value1. Optional: value2 for between.
    action="date": Date validation. Optional: operator, date1, date2.
    action="remove": DESTRUCTIVE. Remove validations from range.
    action="formula": Custom formula validation. Requires: formula. Optional: error_title, error_message, show_error.
    """
    if action == "dropdown":
        if options is None and source_range is None:
            raise ValueError("options or source_range is required for action='dropdown'.")
        return _data_val.add_dropdown_validation(
            file_path,
            sheet_name,
            cell_range,
            options if options is not None else [],
            allow_blank,
            source_range,
            error_style,
            error_title,
            error_message,
            prompt_title,
            prompt_message,
        )
    if action == "numeric":
        if operator is None:
            raise ValueError("operator is required for action='numeric'.")
        if value1 is None:
            raise ValueError("value1 is required for action='numeric'.")
        return _data_val.add_numeric_validation(
            file_path,
            sheet_name,
            cell_range,
            operator,
            value1,
            value2,
            allow_blank,
            error_style,
            error_title,
            error_message,
            prompt_title,
            prompt_message,
        )
    if action == "date":
        return _data_val.add_date_validation(
            file_path,
            sheet_name,
            cell_range,
            operator if operator is not None else "greaterThan",
            date1,
            date2,
            allow_blank,
            error_style,
            error_title,
            error_message,
            prompt_title,
            prompt_message,
        )
    if action == "remove":
        return _data_val.remove_validation(file_path, sheet_name, cell_range)
    if action == "formula":
        if formula is None:
            raise ValueError("formula is required for action='formula'.")
        return _data_val.add_formula_validation(
            file_path,
            sheet_name,
            cell_range,
            formula,
            error_title or "Invalid",
            error_message or "Value does not meet the criteria.",
            show_error,
        )
    raise ValueError(f"Unknown action: {action}")


def doc_properties(
    action: Literal["get", "set_calc_mode"],
    file_path: str,
    calc_mode: str = "auto",
) -> dict | str:
    """Document property and calculation operations.

    action="get": Read workbook properties. Read-only.
    action="set_calc_mode": Set calculation mode. Requires: calc_mode ('auto', 'manual', or 'autoNoTable').
    """
    if action == "get":
        return _doc_props.get_document_properties(file_path)
    if action == "set_calc_mode":
        return _doc_props.set_calculation_mode(file_path, calc_mode)
    raise ValueError(f"Unknown action: {action}")


def conditional_format(
    action: Literal["apply", "highlight", "formula_rule", "remove", "top_bottom", "above_below_average", "list"],
    file_path: str,
    sheet_name: str,
    cell_range: str | None = None,
    format_type: str | None = None,
    start_color: str = "FF0000",
    mid_color: str = "FFFF00",
    end_color: str = "00FF00",
    bar_color: str = "FF638EC6",
    icon_style: str = "3Arrows",
    stop_if_true: bool = False,
    operator: str | None = None,
    formula: str | None = None,
    font_color: str = "9C0006",
    bg_color: str = "FFC7CE",
    is_top: bool = True,
    rank: int = 10,
    percent: bool = False,
    is_above: bool = True,
    equal_average: bool = False,
) -> str | dict | list:
    """Conditional formatting operations.

    action="apply": Apply color_scale/2_color_scale/data_bar/icon_set. Requires: cell_range, format_type.
    action="highlight": Highlight rule based on operator. Requires: cell_range, operator, formula.
    action="formula_rule": Custom formula-based rule. Requires: cell_range, formula.
    action="remove": DESTRUCTIVE. Remove rules. Optional: cell_range (omit to clear all).
    action="top_bottom": Top/bottom N rule. Requires: cell_range. Optional: is_top, rank, percent, bg_color, font_color.
    action="above_below_average": Above/below average rule. Requires: cell_range. Optional: is_above, equal_average, bg_color, font_color.
    action="list": List all conditional formatting rules on the sheet.
    """
    if action == "apply":
        if cell_range is None:
            raise ValueError("cell_range is required for action='apply'.")
        if format_type is None:
            raise ValueError("format_type is required for action='apply'.")
        return _cond_fmt.apply_conditional_formatting(
            file_path,
            sheet_name,
            cell_range,
            format_type,
            start_color,
            mid_color,
            end_color,
            bar_color,
            icon_style,
            stop_if_true,
        )
    if action == "highlight":
        if cell_range is None:
            raise ValueError("cell_range is required for action='highlight'.")
        if operator is None:
            raise ValueError("operator is required for action='highlight'.")
        if formula is None:
            raise ValueError("formula is required for action='highlight'.")
        return _cond_fmt.add_highlight_rule(
            file_path,
            sheet_name,
            cell_range,
            operator,
            formula,
            font_color,
            bg_color,
            stop_if_true,
        )
    if action == "formula_rule":
        if cell_range is None:
            raise ValueError("cell_range is required for action='formula_rule'.")
        if formula is None:
            raise ValueError("formula is required for action='formula_rule'.")
        return _cond_fmt.add_formula_rule(file_path, sheet_name, cell_range, formula, font_color, bg_color)
    if action == "remove":
        return _cond_fmt.remove_conditional_formatting(file_path, sheet_name, cell_range)
    if action == "top_bottom":
        if cell_range is None:
            raise ValueError("cell_range is required for action='top_bottom'.")
        return _cond_fmt.add_top_bottom_rule(
            file_path,
            sheet_name,
            cell_range,
            is_top,
            rank,
            percent,
            bg_color,
            font_color,
        )
    if action == "above_below_average":
        if cell_range is None:
            raise ValueError("cell_range is required for action='above_below_average'.")
        return _cond_fmt.add_above_below_average_rule(
            file_path,
            sheet_name,
            cell_range,
            is_above,
            equal_average,
            bg_color,
            font_color,
        )
    if action == "list":
        return _cond_fmt.list_conditional_formatting(file_path, sheet_name)
    raise ValueError(f"Unknown action: {action}")


def register(mcp) -> None:
    """Register tools on *mcp*."""
    mcp.tool()(protection)
    mcp.tool()(data_validation)
    mcp.tool()(doc_properties)
    mcp.tool()(conditional_format)
