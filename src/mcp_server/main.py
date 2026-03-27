from __future__ import annotations

import json
from logging import Logger
from typing import Any, Literal

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

# Module imports (clean, one per module)
import mcp_server.tools.analysis as _analysis
import mcp_server.tools.cell_ops as _cell_ops
import mcp_server.tools.charts as _charts
import mcp_server.tools.cleaning as _cleaning
import mcp_server.tools.comments as _comments
import mcp_server.tools.conditional_formatting as _cond_fmt
import mcp_server.tools.csv_ops as _csv_ops
import mcp_server.tools.data_validation as _data_val
import mcp_server.tools.doc_properties as _doc_props
import mcp_server.tools.financial as _financial
import mcp_server.tools.formatting as _formatting
import mcp_server.tools.formulas as _formulas
import mcp_server.tools.hyperlinks as _hyperlinks
import mcp_server.tools.multi_file as _multi_file
import mcp_server.tools.named_ranges as _named_ranges
import mcp_server.tools.pivot_etl as _pivot_etl
import mcp_server.tools.protection as _protection
import mcp_server.tools.scenarios as _scenarios
import mcp_server.tools.solver as _solver
import mcp_server.tools.statistical as _statistical
import mcp_server.tools.tables as _tables
import mcp_server.tools.workbook as _workbook
import mcp_server.tools.worksheet_ops as _ws_ops
from mcp_server.utils.logger import configure_logging

mcp: FastMCP = FastMCP("excel-mcp-server")
logger: Logger = configure_logging("mcp_server.main")


# ---------------------------------------------------------------------------
# Resources
# ---------------------------------------------------------------------------


@mcp.resource("excel://workbook/{file_path}/sheets")
def resource_list_sheets(file_path: str) -> str:
    """List all sheets in a workbook as JSON."""
    return json.dumps(_workbook.get_workbook_metadata(file_path)["sheets"], indent=2)


@mcp.resource("excel://workbook/{file_path}/sheet/{sheet_name}/preview")
def resource_sheet_preview(file_path: str, sheet_name: str) -> str:
    """Preview the first 20 rows of a sheet as JSON."""
    data = _cell_ops.read_range(file_path, sheet_name, "A1", "Z20")
    return json.dumps(data, indent=2, default=str)


# ---------------------------------------------------------------------------
# Workbook Management (4 standalone tools — genuinely distinct)
# ---------------------------------------------------------------------------


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def get_workbook_metadata(file_path: str) -> dict:
    """Get workbook metadata including sheet names, dimensions, active sheet, and named ranges."""
    return _workbook.get_workbook_metadata(file_path)


@mcp.tool()
def create_workbook(file_path: str, sheet_names: list[str] | None = None, sheet_name: str | None = None) -> dict:
    """Create a new .xlsx workbook.

    Optionally specify initial sheet names via sheet_names (list) or sheet_name (single).
    """
    return _workbook.create_workbook(file_path, sheet_names, sheet_name)


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def get_sheet_summary(file_path: str, sheet_name: str) -> dict:
    """Get sheet summary: name, row/col counts, headers, and used range."""
    return _workbook.get_sheet_summary(file_path, sheet_name)


@mcp.tool()
def write_multi_sheet(file_path: str, sheets: list[dict]) -> dict:
    """Create a new workbook with multiple named sheets, headers, data, and column widths in one call."""
    return _workbook.write_multi_sheet(file_path, sheets)


# ---------------------------------------------------------------------------
# Sheet Management (consolidated — rename/delete/copy)
# ---------------------------------------------------------------------------


@mcp.tool()
def sheet_management(
    action: Literal["rename", "delete", "copy"],
    file_path: str,
    sheet_name: str,
    new_name: str | None = None,
) -> str:
    """Manage worksheets within a workbook.

    action="rename": Rename a sheet. Requires: new_name.
    action="delete": DESTRUCTIVE. Delete a sheet. Raises if only sheet.
    action="copy": Copy a sheet. Requires: new_name for the copy.
    """
    if action == "rename":
        return _workbook.rename_sheet(file_path, sheet_name, new_name)
    if action == "delete":
        return _workbook.delete_sheet(file_path, sheet_name)
    if action == "copy":
        return _workbook.copy_sheet(file_path, sheet_name, new_name)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Cell Read Operations (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def read_cells(
    mode: Literal["single", "range", "chunked"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    start_cell: str | None = None,
    end_cell: str | None = None,
    include_formula: bool = False,
    include_metadata: bool = False,
    show_formula: bool = False,
    show_style: bool = False,
    output_format: str = "json",
    max_cells: int | None = None,
    start_row: int = 0,
    chunk_size: int = 1000,
) -> dict:
    """Read cell data from a worksheet.

    mode="single": Read one cell. Requires: cell_ref. Optional: include_formula, include_metadata.
    mode="range": Read rectangular range. Requires: start_cell, end_cell.
        Optional: show_formula, show_style, output_format, max_cells.
    mode="chunked": Read large sheets in chunks. Optional: start_row, chunk_size.
    """
    if mode == "single":
        return _cell_ops.read_cell(file_path, sheet_name, cell_ref, include_formula, include_metadata)
    if mode == "range":
        return _cell_ops.read_range(
            file_path,
            sheet_name,
            start_cell,
            end_cell,
            show_formula,
            show_style,
            output_format,
            max_cells,
        )
    if mode == "chunked":
        return _cell_ops.read_file_chunked(file_path, sheet_name, start_row, chunk_size)
    raise ValueError(f"Unknown mode: {mode}")


# ---------------------------------------------------------------------------
# Cell Write Operations (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def write_cells(
    mode: Literal["single", "range", "series"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    value: str | int | float | bool | None = None,
    start_cell: str | None = None,
    data: list[list] | None = None,
    series_type: str = "number",
    count: int | None = None,
    step: float | str = 1,
    direction: str = "down",
    start_value: float | int | None = None,
) -> str | dict:
    """Write data to cells.

    mode="single": Write one cell. Requires: cell_ref, value.
    mode="range": Write 2D array. Requires: start_cell, data.
    mode="series": Fill series. Requires: start_cell, count. Optional: series_type, step, direction, start_value.
    """
    if mode == "single":
        return _cell_ops.write_cell(file_path, sheet_name, cell_ref, value)
    if mode == "range":
        if start_cell is None:
            raise ValueError("start_cell is required for mode='range'.")
        if data is None:
            raise ValueError("data is required for mode='range'.")
        return _cell_ops.write_range(file_path, sheet_name, start_cell, data)
    if mode == "series":
        if start_cell is None:
            raise ValueError("start_cell is required for mode='series'.")
        if count is None:
            raise ValueError("count is required for mode='series'.")
        return _cell_ops.fill_series(
            file_path, sheet_name, start_cell, series_type, count, step, direction, start_value
        )
    raise ValueError(f"Unknown mode: {mode}")


# ---------------------------------------------------------------------------
# Cell Operations (standalone — distinct operations)
# ---------------------------------------------------------------------------


@mcp.tool(annotations=ToolAnnotations(destructiveHint=True))
def clear_range(file_path: str, sheet_name: str, start_cell: str, end_cell: str) -> str:
    """Clear all values in a rectangular cell range."""
    return _cell_ops.clear_range(file_path, sheet_name, start_cell, end_cell)


@mcp.tool()
def copy_range(
    file_path: str,
    source_sheet: str,
    source_range: str,
    dest_sheet: str,
    dest_range: str,
    copy_values: bool = True,
    copy_styles: bool = True,
) -> str:
    """Copy cells from source range to destination range within same or across sheets."""
    return _cell_ops.copy_range(file_path, source_sheet, source_range, dest_sheet, dest_range, copy_values, copy_styles)


# ---------------------------------------------------------------------------
# Formatting (2 standalone tools — distinct operations)
# ---------------------------------------------------------------------------


@mcp.tool()
def format_cells(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    bold: bool = False,
    italic: bool = False,
    font_size: int | None = None,
    font_color: str | None = None,
    bg_color: str | None = None,
    number_format: str | None = None,
    horizontal_alignment: str | None = None,
    vertical_alignment: str | None = None,
    wrap_text: bool = False,
    border_style: str | None = None,
    border_color: str | None = None,
    font_name: str | None = None,
    underline: str | None = None,
    strikethrough: bool = False,
    text_rotation: int | None = None,
    indent: int | None = None,
    shrink_to_fit: bool = False,
    top_border_style: str | None = None,
    bottom_border_style: str | None = None,
    left_border_style: str | None = None,
    right_border_style: str | None = None,
) -> str:
    """Apply formatting (font, fill, alignment, borders, number format) to a cell range."""
    return _formatting.format_cells(
        file_path,
        sheet_name,
        cell_range,
        bold,
        italic,
        font_size,
        font_color,
        bg_color,
        number_format,
        horizontal_alignment,
        vertical_alignment,
        wrap_text,
        border_style,
        border_color,
        font_name,
        underline,
        strikethrough,
        text_rotation,
        indent,
        shrink_to_fit,
        top_border_style,
        bottom_border_style,
        left_border_style,
        right_border_style,
    )


@mcp.tool()
def auto_fit_columns(file_path: str, sheet_name: str) -> str:
    """Auto-fit all column widths based on content length."""
    return _formatting.auto_fit_columns(file_path, sheet_name)


# ---------------------------------------------------------------------------
# Formula Write (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def formula_write(
    action: Literal["set", "batch"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    formula: str | None = None,
    is_array: bool = False,
    target_range: str | None = None,
    formulas: dict[str, str] | None = None,
) -> str:
    """Write formulas to cells.

    action="set": Set a formula on a cell. Requires: cell_ref, formula. Optional: is_array, target_range.
    action="batch": Set multiple formulas. Requires: formulas (dict of cell_ref -> formula).
    """
    if action == "set":
        return _formulas.set_formula(file_path, sheet_name, cell_ref, formula, is_array, target_range)
    if action == "batch":
        return _formulas.set_formulas_batch(file_path, sheet_name, formulas)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Formula Audit (consolidated — all read-only)
# ---------------------------------------------------------------------------


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def formula_audit(
    action: Literal["value", "errors", "precedents", "dependents", "list"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    cell_range: str | None = None,
) -> dict | list[dict] | list[str]:
    """Audit and inspect formulas. All actions are read-only.

    action="value": Get cached display value. Requires: cell_ref.
    action="errors": Find error cells (#VALUE!, #REF!, etc.). Optional: cell_range to limit scope.
    action="precedents": Cells feeding into cell_ref. Requires: cell_ref.
    action="dependents": Cells referencing cell_ref. Requires: cell_ref.
    action="list": List all formulas in the sheet.
    """
    if action == "value":
        return _formulas.get_formula_value(file_path, sheet_name, cell_ref)
    if action == "errors":
        return _formulas.get_formula_errors(file_path, sheet_name, cell_range)
    if action == "precedents":
        return _formulas.get_formula_precedents(file_path, sheet_name, cell_ref)
    if action == "dependents":
        return _formulas.get_formula_dependents(file_path, sheet_name, cell_ref)
    if action == "list":
        return _formulas.list_formulas(file_path, sheet_name)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# CSV Operations (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def csv_ops(
    action: Literal["preview", "to_xlsx", "to_csv"],
    file_path: str | None = None,
    csv_path: str | None = None,
    xlsx_path: str | None = None,
    sheet_name: str = "Sheet1",
    output_path: str | None = None,
    rows: int = 10,
    delimiter: str = ",",
    encoding: str = "utf-8",
) -> dict | str:
    """CSV operations: preview, convert CSV to XLSX, export XLSX to CSV.

    action="preview": Preview CSV. Requires: file_path. Optional: rows, delimiter, encoding.
    action="to_xlsx": Convert CSV to XLSX. Requires: csv_path, xlsx_path. Optional: sheet_name, delimiter, encoding.
    action="to_csv": Export sheet to CSV. Requires: file_path, sheet_name, output_path. Optional: delimiter, encoding.
    """
    if action == "preview":
        return _csv_ops.read_csv_preview(file_path, rows, delimiter, encoding)
    if action == "to_xlsx":
        effective_csv = csv_path or file_path
        effective_xlsx = xlsx_path or output_path
        if not effective_csv:
            raise ValueError("csv_path (or file_path) is required for action='to_xlsx'.")
        if not effective_xlsx:
            raise ValueError("xlsx_path (or output_path) is required for action='to_xlsx'.")
        return _csv_ops.csv_to_xlsx(effective_csv, effective_xlsx, sheet_name, delimiter, encoding)
    if action == "to_csv":
        return _csv_ops.xlsx_to_csv(file_path, sheet_name, output_path, delimiter, encoding)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Conditional Formatting (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def conditional_format(
    action: Literal["apply", "highlight", "formula_rule", "remove"],
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
) -> str:
    """Conditional formatting operations.

    action="apply": Apply color_scale/2_color_scale/data_bar/icon_set. Requires: cell_range, format_type.
    action="highlight": Highlight rule based on operator. Requires: cell_range, operator, formula.
    action="formula_rule": Custom formula-based rule. Requires: cell_range, formula.
    action="remove": DESTRUCTIVE. Remove rules. Optional: cell_range (omit to clear all).
    """
    if action == "apply":
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
        return _cond_fmt.add_formula_rule(file_path, sheet_name, cell_range, formula, font_color, bg_color)
    if action == "remove":
        return _cond_fmt.remove_conditional_formatting(file_path, sheet_name, cell_range)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Tables (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def table(
    action: Literal["create", "list", "resize", "totals", "data"],
    file_path: str,
    sheet_name: str,
    table_name: str | None = None,
    data_range: str | None = None,
    style_name: str = "TableStyleMedium9",
    new_range: str | None = None,
    show_totals: bool | None = None,
    column_totals: dict[str, str] | None = None,
) -> str | list[dict] | dict:
    """Excel table (ListObject) operations.

    action="create": Create a table. Requires: data_range, table_name. Optional: style_name.
    action="list": List all tables. Read-only.
    action="resize": Resize a table. Requires: table_name, new_range.
    action="totals": Toggle totals row. Requires: table_name, show_totals. Optional: column_totals.
    action="data": Read table data. Read-only. Requires: table_name.
    """
    if action == "create":
        if not data_range:
            raise ValueError("data_range is required for action='create'.")
        if not table_name:
            raise ValueError("table_name is required for action='create'.")
        return _tables.create_table(file_path, sheet_name, data_range, table_name, style_name)
    if action == "list":
        return _tables.list_tables(file_path, sheet_name)
    if action == "resize":
        return _tables.resize_table(file_path, sheet_name, table_name, new_range)
    if action == "totals":
        return _tables.set_table_totals_row(file_path, sheet_name, table_name, show_totals, column_totals)
    if action == "data":
        return _tables.get_table_data(file_path, sheet_name, table_name)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Data Validation (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def data_validation(
    action: Literal["dropdown", "numeric", "date", "remove"],
    file_path: str,
    sheet_name: str,
    cell_range: str,
    options: list[str] | None = None,
    source_range: str | None = None,
    operator: str | None = None,
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
) -> str:
    """Data validation operations.

    action="dropdown": Add dropdown list. Requires: options or source_range.
    action="numeric": Numeric validation. Requires: operator, value1. Optional: value2 for between.
    action="date": Date validation. Optional: operator, date1, date2.
    action="remove": DESTRUCTIVE. Remove validations from range.
    """
    if action == "dropdown":
        return _data_val.add_dropdown_validation(
            file_path,
            sheet_name,
            cell_range,
            options,
            allow_blank,
            source_range,
            error_style,
            error_title,
            error_message,
            prompt_title,
            prompt_message,
        )
    if action == "numeric":
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
            operator,
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
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Protection (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
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
        return _protection.unprotect_sheet(file_path, sheet_name, password)
    if action == "protect_cells":
        if not locked_range:
            raise ValueError("locked_range is required for action='protect_cells'.")
        return _protection.protect_cells(file_path, sheet_name, locked_range, unlocked_ranges)
    if action == "protect_workbook":
        return _doc_props.protect_workbook(file_path, password, lock_structure, lock_windows)
    if action == "unprotect_workbook":
        return _doc_props.unprotect_workbook(file_path)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Charts (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def chart(
    action: Literal["create", "delete", "list", "add_series", "set_axes", "trendline", "combo"],
    file_path: str,
    sheet_name: str,
    chart_index: int | None = None,
    data_range: str | None = None,
    chart_type: str = "column",
    target_cell: str = "E1",
    title: str | None = None,
    x_axis_title: str = "",
    y_axis_title: str = "",
    style: int = 10,
    width: float = 15,
    height: float = 10,
    title_from_data: bool = True,
    x_title: str | None = None,
    y_title: str | None = None,
    x_min: float | None = None,
    x_max: float | None = None,
    y_min: float | None = None,
    y_max: float | None = None,
    y_number_format: str | None = None,
    log_scale_y: bool = False,
    series_index: int = 0,
    trendline_type: str = "linear",
    name: str | None = None,
    periods_forward: int = 0,
    periods_backward: int = 0,
    bar_columns: list[int] | None = None,
    line_columns: list[int] | None = None,
    anchor_cell: str = "F1",
    x_axis_column: int = 0,
) -> str | list[dict]:
    """Chart operations.

    action="create": Create chart. Requires: data_range. Optional: chart_type, target_cell, title, axes, style, size.
    action="delete": DESTRUCTIVE. Delete chart. Optional: chart_index (default 0).
    action="list": List all charts. Read-only.
    action="add_series": Add data series. Requires: chart_index, data_range.
    action="set_axes": Configure axes. Optional: chart_index, x/y titles, min/max, number_format, log_scale.
    action="trendline": Add trendline. Optional: chart_index, series_index, trendline_type, periods.
    action="combo": Create combo chart. Requires: data_range, bar_columns, line_columns.
    """
    if action == "create":
        return _charts.create_chart(
            file_path,
            sheet_name,
            data_range,
            chart_type,
            target_cell,
            title or "",
            x_axis_title,
            y_axis_title,
            style,
            width,
            height,
        )
    if action == "delete":
        return _charts.delete_chart(file_path, sheet_name, chart_index or 0)
    if action == "list":
        return _charts.list_charts(file_path, sheet_name)
    if action == "add_series":
        return _charts.add_chart_series(file_path, sheet_name, chart_index, data_range, title_from_data)
    if action == "set_axes":
        return _charts.set_chart_axes(
            file_path,
            sheet_name,
            chart_index or 0,
            x_title,
            y_title,
            x_min,
            x_max,
            y_min,
            y_max,
            y_number_format,
            log_scale_y,
        )
    if action == "trendline":
        return _charts.add_chart_trendline(
            file_path,
            sheet_name,
            chart_index or 0,
            series_index,
            trendline_type,
            name,
            periods_forward,
            periods_backward,
        )
    if action == "combo":
        if not data_range:
            raise ValueError("data_range is required for action='combo'.")
        if not bar_columns:
            raise ValueError("bar_columns is required for action='combo'.")
        if not line_columns:
            raise ValueError("line_columns is required for action='combo'.")
        return _charts.create_combo_chart(
            file_path,
            sheet_name,
            data_range,
            bar_columns,
            line_columns,
            title,
            anchor_cell,
            x_axis_column,
            width,
            height,
        )
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Document Properties (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
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


# ---------------------------------------------------------------------------
# Named Ranges (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def named_range(
    action: Literal["list", "create", "delete", "update"],
    file_path: str,
    name: str | None = None,
    destination: str | None = None,
    scope: str = "workbook",
    new_destination: str | None = None,
) -> list[dict] | str:
    """Manage named ranges.

    action="list": List all named ranges. Read-only. Requires: file_path only.
    action="create": Create a named range. Requires: name, destination. Optional: scope.
    action="delete": DESTRUCTIVE. Delete a named range. Requires: name.
    action="update": Update destination. Requires: name, new_destination.
    """
    if action == "list":
        return _named_ranges.list_named_ranges(file_path)
    if action == "create":
        if not name:
            raise ValueError("name is required for action='create'.")
        if not destination:
            raise ValueError("destination is required for action='create'.")
        return _named_ranges.create_named_range(file_path, name, destination, scope)
    if action == "delete":
        return _named_ranges.delete_named_range(file_path, name)
    if action == "update":
        if not name:
            raise ValueError("name is required for action='update'.")
        if not new_destination:
            raise ValueError("new_destination is required for action='update'.")
        return _named_ranges.update_named_range(file_path, name, new_destination)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Comments (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def comment(
    action: Literal["add", "read", "delete", "list"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    text: str | None = None,
    author: str = "Excel MCP",
) -> str | dict | list[dict] | None:
    """Comment operations on cells.

    action="add": Add a comment. Requires: cell_ref, text. Optional: author.
    action="read": Read a comment. Read-only. Requires: cell_ref. Returns None if no comment.
    action="delete": DESTRUCTIVE. Delete a comment. Requires: cell_ref.
    action="list": List all comments in sheet. Read-only.
    """
    if action == "add":
        return _comments.add_comment(file_path, sheet_name, cell_ref, text, author)
    if action == "read":
        return _comments.read_comment(file_path, sheet_name, cell_ref)
    if action == "delete":
        return _comments.delete_comment(file_path, sheet_name, cell_ref)
    if action == "list":
        return _comments.list_comments(file_path, sheet_name)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Hyperlinks (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def hyperlink(
    action: Literal["add", "read", "delete", "list"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    url: str | None = None,
    display_text: str | None = None,
    tooltip: str | None = None,
) -> str | dict | list[dict] | None:
    """Hyperlink operations.

    action="add": Add hyperlink. Requires: cell_ref, url. Optional: display_text, tooltip.
    action="read": Read hyperlink. Read-only. Requires: cell_ref. Returns None if none.
    action="delete": DESTRUCTIVE. Delete hyperlink. Requires: cell_ref.
    action="list": List all hyperlinks. Read-only.
    """
    if action == "add":
        return _hyperlinks.add_hyperlink(file_path, sheet_name, cell_ref, url, display_text, tooltip)
    if action == "read":
        return _hyperlinks.read_hyperlink(file_path, sheet_name, cell_ref)
    if action == "delete":
        return _hyperlinks.delete_hyperlink(file_path, sheet_name, cell_ref)
    if action == "list":
        return _hyperlinks.list_hyperlinks(file_path, sheet_name)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Scenario Management (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def scenario(
    action: Literal["add", "list", "apply"],
    file_path: str,
    name: str | None = None,
    cell_values: dict[str, dict[str, Any]] | None = None,
    description: str = "",
) -> str | list[dict] | dict:
    """Scenario management for what-if analysis.

    action="add": Save a scenario. Requires: name, cell_values ({sheet: {cell: value}}). Optional: description.
    action="list": List all scenarios. Read-only.
    action="apply": DESTRUCTIVE. Apply scenario values to sheet. Requires: name.
    """
    if action == "add":
        return _scenarios.add_scenario(file_path, name, cell_values, description)
    if action == "list":
        return _scenarios.list_scenarios(file_path)
    if action == "apply":
        return _scenarios.apply_scenario(file_path, name)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Multi-File Operations (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def multi_file(
    action: Literal["aggregate", "filter", "validate", "compare"],
    file_paths: list[str] | None = None,
    file_a: str | None = None,
    file_b: str | None = None,
    column: str | None = None,
    operation: str = "sum",
    operator: str | None = None,
    value: str | float | None = None,
    key_column: str | None = None,
    check_columns: list[str] | None = None,
    sheet_name: str = "Sheet1",
    header_row: int = 1,
    output_file: str | None = None,
    compare_values: bool = True,
    compare_formulas: bool = False,
) -> dict:
    """Cross-file operations.

    action="aggregate": Aggregate a column across files. Requires: file_paths, column. Optional: operation.
    action="filter": Filter rows across files. Requires: file_paths, column, operator, value.
    action="validate": Cross-file consistency check. Requires: file_paths, key_column.
    action="compare": Compare two workbooks. Requires: file_a, file_b. Optional: sheet_name.
    """
    if action == "aggregate":
        return _multi_file.bulk_aggregate_multi_files(
            file_paths,
            column,
            operation,
            sheet_name,
            header_row,
            output_file,
        )
    if action == "filter":
        return _multi_file.bulk_filter_multi_files(
            file_paths,
            column,
            operator,
            value,
            sheet_name,
            header_row,
            output_file,
        )
    if action == "validate":
        return _multi_file.validate_data_consistency(file_paths, key_column, check_columns, sheet_name, header_row)
    if action == "compare":
        if not file_a:
            raise ValueError("file_a is required for action='compare'.")
        if not file_b:
            raise ValueError("file_b is required for action='compare'.")
        return _multi_file.compare_workbooks(file_a, file_b, sheet_name, output_file)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Worksheet Operations (consolidated: view + cross-workbook)
# ---------------------------------------------------------------------------


@mcp.tool()
def worksheet_ops(
    action: Literal["freeze", "auto_filter", "copy_range_across", "copy_sheet_across", "merge_workbooks"],
    file_path: str | None = None,
    sheet_name: str | None = None,
    cell_ref: str | None = None,
    cell_range: str | None = None,
    remove: bool = False,
    source_sheet: str | None = None,
    source_range: str | None = None,
    target_sheet: str | None = None,
    target_start_cell: str = "A1",
    copy_values: bool = True,
    copy_styles: bool = False,
    source_file: str | None = None,
    dest_file: str | None = None,
    dest_sheet_name: str | None = None,
    source_files: list[str] | None = None,
    output_file: str | None = None,
    conflict_strategy: str = "rename",
) -> str | dict:
    """Worksheet operations: freeze panes, auto filter, cross-sheet/workbook operations.

    action="freeze": Freeze panes at cell_ref. Requires: file_path, sheet_name. Optional: cell_ref (None to unfreeze).
    action="auto_filter": Toggle auto-filter. Requires: file_path, sheet_name. Optional: cell_range, remove.
    action="copy_range_across": Copy range between sheets.
        Requires: file_path, source_sheet, source_range, target_sheet.
        Optional: target_start_cell (default "A1") — top-left cell in target_sheet to paste into.
    action="copy_sheet_across": Copy sheet between workbooks. Requires: source_file, source_sheet, dest_file.
    action="merge_workbooks": Merge multiple workbooks. Requires: source_files, output_file.
    """
    if action == "freeze":
        return _ws_ops.freeze_panes(file_path, sheet_name, cell_ref)
    if action == "auto_filter":
        if not remove and cell_range is None:
            raise ValueError("cell_range is required for action='auto_filter' when remove=False.")
        return _ws_ops.set_auto_filter(file_path, sheet_name, cell_range, remove)
    if action == "copy_range_across":
        if not source_sheet:
            raise ValueError("source_sheet is required for action='copy_range_across'.")
        if not source_range:
            raise ValueError("source_range is required for action='copy_range_across'.")
        if not target_sheet:
            raise ValueError("target_sheet is required for action='copy_range_across'.")
        return _ws_ops.copy_range_across_sheets(
            file_path,
            source_sheet,
            source_range,
            target_sheet,
            target_start_cell,
            copy_values,
            copy_styles,
        )
    if action == "copy_sheet_across":
        if not source_file:
            raise ValueError("source_file is required for action='copy_sheet_across'.")
        if not source_sheet:
            raise ValueError("source_sheet is required for action='copy_sheet_across'.")
        if not dest_file:
            raise ValueError("dest_file is required for action='copy_sheet_across'.")
        return _ws_ops.copy_sheet_across_workbooks(source_file, source_sheet, dest_file, dest_sheet_name)
    if action == "merge_workbooks":
        if not source_files:
            raise ValueError("source_files is required for action='merge_workbooks'.")
        if not output_file:
            raise ValueError("output_file is required for action='merge_workbooks'.")
        return _ws_ops.merge_workbooks(source_files, output_file, conflict_strategy)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Data Analysis (6 standalone tools — genuinely distinct param shapes)
# ---------------------------------------------------------------------------


@mcp.tool()
def sort_data(
    file_path: str,
    sheet_name: str,
    sort_by: list[dict] | None = None,
    column: str | None = None,
    ascending: bool = True,
    has_header: bool = True,
) -> str:
    """Sort sheet data by one or more columns and write back."""
    return _analysis.sort_data(file_path, sheet_name, sort_by, column, ascending, has_header)


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def column_statistics(file_path: str, sheet_name: str, column: str, has_header: bool = True) -> dict:
    """Compute descriptive statistics (mean, median, std, min, max, sum) for a numeric column."""
    return _analysis.column_statistics(file_path, sheet_name, column, has_header)


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def aggregate_data(
    file_path: str,
    sheet_name: str,
    group_by: str | list[str],
    value_column: str,
    operation: str = "sum",
    has_header: bool = True,
    aggfunc: str | dict | None = None,
) -> dict:
    """Group by one or more columns and aggregate (sum, mean, count, min, max, median, std)."""
    return _analysis.aggregate_data(file_path, sheet_name, group_by, value_column, operation, has_header, aggfunc)


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def find_duplicates(file_path: str, sheet_name: str, columns: list[str], has_header: bool = True) -> dict:
    """Find duplicate rows based on specified columns."""
    return _analysis.find_duplicates(file_path, sheet_name, columns, has_header)


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def vlookup_helper(
    lookup_file: str,
    data_file: str,
    lookup_column: str,
    data_key_column: str,
    data_return_columns: list[str],
    lookup_sheet: str = "Sheet1",
    data_sheet: str = "Sheet1",
    fuzzy: bool = False,
    fuzzy_threshold: float = 0.8,
    output_file: str | None = None,
    header_row: int = 1,
) -> dict:
    """Cross-file VLOOKUP with optional fuzzy string matching."""
    return _analysis.vlookup_helper(
        lookup_file,
        data_file,
        lookup_column,
        data_key_column,
        data_return_columns,
        lookup_sheet,
        data_sheet,
        fuzzy,
        fuzzy_threshold,
        output_file,
        header_row,
    )


@mcp.tool(annotations=ToolAnnotations(destructiveHint=True))
def filter_data_advanced(
    file_path: str,
    sheet_name: str,
    conditions: list[dict],
    logic: str = "AND",
    output_sheet: str | None = None,
    header_row: int = 1,
) -> dict:
    """Multi-condition AND/OR filtering. Each condition: {column, operator, value}."""
    return _analysis.filter_data_advanced(file_path, sheet_name, conditions, logic, output_sheet, header_row)


# ---------------------------------------------------------------------------
# Pivot & ETL (5 standalone tools)
# ---------------------------------------------------------------------------


@mcp.tool()
def create_pivot_table(
    file_path: str,
    sheet_name: str,
    index_cols: list[str],
    value_cols: list[str],
    aggfunc: str | dict = "sum",
    output_sheet: str | None = None,
    output_file: str | None = None,
    column_field: str | None = None,
) -> dict:
    """Create a pivot table and optionally write results to a sheet or file."""
    return _pivot_etl.create_pivot_table(
        file_path,
        sheet_name,
        index_cols,
        value_cols,
        aggfunc,
        output_sheet,
        output_file,
        False,
        column_field,
    )


@mcp.tool()
def unpivot_data(
    file_path: str,
    sheet_name: str,
    id_vars: list[str],
    value_vars: list[str],
    var_name: str = "Variable",
    value_name: str = "Value",
) -> dict:
    """Unpivot (melt) data from wide to long format."""
    return _pivot_etl.unpivot_data(file_path, sheet_name, id_vars, value_vars, var_name, value_name)


@mcp.tool()
def merge_datasets(
    file_path: str,
    sheet1: str,
    sheet2: str,
    join_key: str | list[str],
    how: str = "left",
    output_sheet: str | None = None,
) -> dict:
    """Merge two sheets like a SQL join (left, right, inner, outer)."""
    return _pivot_etl.merge_datasets(file_path, sheet1, sheet2, join_key, how, output_sheet)


@mcp.tool()
def add_computed_column(
    file_path: str,
    sheet_name: str,
    new_column_name: str,
    expression: str,
    has_header: bool = True,
) -> str:
    """Add a computed column using a pandas-eval expression (e.g. 'Revenue - Cost')."""
    return _pivot_etl.add_computed_column(file_path, sheet_name, new_column_name, expression, has_header)


@mcp.tool(annotations=ToolAnnotations(destructiveHint=True))
def deduplicate_data(
    file_path: str,
    sheet_name: str,
    columns: list[str] | None = None,
    keep: str = "first",
) -> str:
    """Remove duplicate rows from a sheet. keep: 'first', 'last', or False."""
    return _pivot_etl.deduplicate_data(file_path, sheet_name, columns, keep)


# ---------------------------------------------------------------------------
# Financial (7 standalone + 1 consolidated TVM tool)
# ---------------------------------------------------------------------------


@mcp.tool()
def goal_seek(
    file_path: str,
    sheet_name: str,
    variable_cell: str,
    expression: str,
    target_value: float,
    initial_value: float = 0.0,
    tolerance: float = 1e-6,
    max_iterations: int = 1000,
) -> dict:
    """Find the variable_cell value that makes expression equal target_value, then write it to the workbook."""
    return _financial.goal_seek(
        file_path,
        sheet_name,
        variable_cell,
        expression,
        target_value,
        initial_value,
        tolerance,
        max_iterations,
    )


@mcp.tool()
def loan_amortization(
    principal: float,
    annual_rate: float,
    years: int,
    payments_per_year: int = 12,
) -> dict:
    """Generate a loan amortization schedule with payment breakdown."""
    return _financial.loan_amortization(principal, annual_rate, years, payments_per_year)


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def dcf_analysis(
    cash_flows: list[float],
    discount_rate: float,
    terminal_growth_rate: float = 0.02,
    initial_investment: float = 0.0,
) -> dict:
    """Discounted Cash Flow valuation with Gordon Growth Model terminal value."""
    return _financial.dcf_analysis(cash_flows, discount_rate, terminal_growth_rate, initial_investment)


@mcp.tool()
def budget_variance_analysis(
    file_path: str,
    sheet_name: str = "Sheet1",
    category_column: str = "A",
    budget_column: str = "B",
    actual_column: str = "C",
    header_row: int = 1,
    output_file: str | None = None,
) -> dict:
    """Analyze budget vs actual spending. Returns variance per category with status."""
    return _financial.budget_variance_analysis(
        file_path,
        sheet_name,
        category_column,
        budget_column,
        actual_column,
        header_row,
        output_file,
    )


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def financial_ratio_analysis(financial_data: dict, industry_benchmarks: dict | None = None) -> dict:
    """Compute financial ratios from raw financial metric values with optional benchmark comparison.

    ``financial_data`` is a dict of raw financial metric values (NOT computed ratio names).
    Valid keys: current_assets, current_liabilities, total_debt, total_equity, net_income,
    total_assets, revenue, gross_profit, ebitda, interest_expense.

    Example: {"current_assets": 500000, "current_liabilities": 250000, "net_income": 100000,
              "revenue": 1000000, "total_equity": 400000, "gross_profit": 600000}
    """
    return _financial.financial_ratio_analysis(financial_data, industry_benchmarks)


@mcp.tool()
def break_even_analysis(fixed_costs: float, price_per_unit: float, variable_cost_per_unit: float) -> dict:
    """Calculate break-even point in units and revenue."""
    return _financial.break_even_analysis(fixed_costs, price_per_unit, variable_cost_per_unit)


@mcp.tool()
def create_sensitivity_table(
    file_path: str,
    sheet_name: str,
    output_cell: str,
    expression: str,
    var1_name: str,
    var1_values: list[float],
    var2_name: str | None = None,
    var2_values: list[float] | None = None,
) -> dict:
    """Create a one- or two-variable sensitivity/what-if table in the workbook."""
    return _financial.create_sensitivity_table(
        file_path,
        sheet_name,
        output_cell,
        expression,
        var1_name,
        var1_values,
        var2_name,
        var2_values,
    )


# ---------------------------------------------------------------------------
# Time Value of Money (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def time_value_calc(
    operation: Literal["fv", "pv", "nper", "rate", "depreciation"],
    rate: float | None = None,
    nper: int | None = None,
    pmt: float | None = None,
    pv: float = 0.0,
    fv: float = 0.0,
    when: str = "end",
    guess: float = 0.1,
    cost: float | None = None,
    salvage: float | None = None,
    life: int | None = None,
    method: str = "sln",
    period: int | None = None,
) -> dict:
    """Time value of money and depreciation calculations.

    operation="fv": Future value. Requires: rate, nper, pmt. Optional: pv, when.
    operation="pv": Present value. Requires: rate, nper, pmt. Optional: fv, when.
    operation="nper": Number of periods. Requires: rate, pmt, pv. Optional: fv, when.
    operation="rate": Interest rate. Requires: nper, pmt, pv. Optional: fv, when, guess.
    operation="depreciation": Asset depreciation. Requires: cost, salvage, life. Optional: method, period.
    """
    if operation == "fv":
        return _financial.calculate_fv(rate, nper, pmt, pv, when)
    if operation == "pv":
        return _financial.calculate_pv(rate, nper, pmt, fv, when)
    if operation == "nper":
        return _financial.calculate_nper(rate, pmt, pv, fv, when)
    if operation == "rate":
        return _financial.calculate_rate(nper, pmt, pv, fv, when, guess)
    if operation == "depreciation":
        return _financial.calculate_depreciation(cost, salvage, life, method, period)
    raise ValueError(f"Unknown operation: {operation}")


# ---------------------------------------------------------------------------
# Data Cleaning (3 standalone tools)
# ---------------------------------------------------------------------------


@mcp.tool()
def split_column(
    file_path: str,
    sheet_name: str,
    column: str,
    delimiter: str = ",",
    new_columns: list[str] | None = None,
    drop_original: bool = True,
    output_file: str | None = None,
    header_row: int = 1,
) -> dict:
    """Split a text column into multiple columns by delimiter."""
    return _cleaning.split_column(
        file_path,
        sheet_name,
        column,
        delimiter,
        new_columns,
        drop_original,
        output_file,
        header_row,
    )


@mcp.tool()
def data_cleaner(
    file_path: str,
    sheet_name: str = "Sheet1",
    operations: list[str] | None = None,
    columns: list[str] | None = None,
    preview: bool = False,
    output_file: str | None = None,
    header_row: int = 1,
    fill_missing_strategy: str = "value",
) -> dict:
    """Batch data cleaning pipeline.

    Operations: trim_whitespace, remove_empty_rows, remove_empty_columns,
    normalize_text, fix_numbers, remove_duplicates, fill_missing.
    """
    return _cleaning.data_cleaner(
        file_path,
        sheet_name,
        operations,
        columns,
        preview,
        output_file,
        header_row,
        fill_missing_strategy,
    )


@mcp.tool(annotations=ToolAnnotations(destructiveHint=True))
def parse_date_column(
    file_path: str,
    sheet_name: str,
    column: str,
    output_column: str | None = None,
    output_format: str = "%Y-%m-%d",
    dayfirst: bool = True,
    header_row: int = 1,
) -> dict:
    """Parse mixed date formats in a column and normalise to a standard output format."""
    return _cleaning.parse_date_column(
        file_path,
        sheet_name,
        column,
        output_column,
        output_format,
        dayfirst,
        header_row,
    )


# ---------------------------------------------------------------------------
# Statistical Analysis (2 standalone tools)
# ---------------------------------------------------------------------------


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def run_regression(
    file_path: str,
    sheet_name: str,
    y_column: str,
    x_columns: list[str],
    header_row: int = 1,
    output_file: str | None = None,
) -> dict:
    """Run OLS linear regression and return coefficients, R-squared, and residuals."""
    return _statistical.run_regression(
        file_path,
        sheet_name,
        y_column,
        x_columns,
        output_sheet=output_file or "Regression Output",
        header_row=header_row,
    )


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def run_exponential_smoothing(
    file_path: str,
    sheet_name: str,
    column: str,
    alpha: float = 0.3,
    new_column_name: str | None = None,
    header_row: int = 1,
    output_file: str | None = None,
) -> dict:
    """Apply exponential smoothing to a time series and write the result to a new column."""
    return _statistical.run_exponential_smoothing(
        file_path,
        sheet_name,
        column,
        alpha,
        new_column_name,
        header_row,
    )


# ---------------------------------------------------------------------------
# Solver / Optimization (1 standalone tool)
# ---------------------------------------------------------------------------


@mcp.tool()
def run_solver(
    file_path: str,
    sheet_name: str,
    objective_expression: str,
    variable_cells: dict[str, tuple[float, float]],
    constraints: list[dict] | None = None,
    maximize: bool = False,
    tolerance: float = 1e-6,
    max_iterations: int = 1000,
) -> dict:
    """Multi-variable constrained optimization using scipy.

    objective_expression: arithmetic expression using cell refs (e.g. "B2 * B3 - B4").
    variable_cells: {cell_ref: (lower_bound, upper_bound)} dict.
    constraints: list of {"expression": str, "type": "ineq"|"eq"} dicts.
    maximize: True to maximise instead of minimise.
    """
    return _solver.run_solver(
        file_path,
        sheet_name,
        objective_expression,
        variable_cells,
        constraints,
        maximize,
        tolerance,
        max_iterations,
    )


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------


def run() -> None:
    """Launch the MCP server."""
    try:
        mcp.run()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception:
        logger.critical("Fatal server error", exc_info=True)
        raise


if __name__ == "__main__":
    run()
