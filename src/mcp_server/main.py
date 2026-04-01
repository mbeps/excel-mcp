from __future__ import annotations

import json
from datetime import date, datetime
from logging import Logger
from typing import Literal

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
import mcp_server.tools.custom_code as _custom_code
import mcp_server.tools.data_validation as _data_val
import mcp_server.tools.doc_properties as _doc_props
import mcp_server.tools.financial as _financial
import mcp_server.tools.formatting as _formatting
import mcp_server.tools.formulas as _formulas
import mcp_server.tools.hyperlinks as _hyperlinks
import mcp_server.tools.images as _images
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
from mcp_server.models.analysis import ColumnStats
from mcp_server.models.charts import ChartInfo
from mcp_server.models.comments import CommentInfo
from mcp_server.models.common import (
    BorderStyle,
    HorizontalAlignment,
    ScenarioCellValue,
    ValidationOperator,
    VerticalAlignment,
)
from mcp_server.models.hyperlinks import HyperlinkInfo, HyperlinkReadResult
from mcp_server.models.named_ranges import FormulaErrorInfo, FormulaInfo, NamedRangeInfo
from mcp_server.models.pivot_etl import ChunkReadResult
from mcp_server.models.scenarios import ScenarioInfo
from mcp_server.models.solver import SolverResult
from mcp_server.models.statistics import RegressionResult
from mcp_server.models.tables import TableInfo
from mcp_server.models.workbook import (
    SheetDefinition,
    SheetSummary,
    WorkbookCreatedResult,
    WorkbookMetadata,
    WriteMultiSheetResult,
)
from mcp_server.tools.scenarios import ScenarioApplyResult
from mcp_server.utils.logger import configure_logging

mcp: FastMCP = FastMCP("excel-mcp-server")
logger: Logger = configure_logging("mcp_server.main")


# ---------------------------------------------------------------------------
# Resources
# ---------------------------------------------------------------------------


@mcp.resource("excel://workbook/{file_path}/sheets")
def resource_list_sheets(file_path: str) -> str:
    """List all sheets in a workbook as JSON."""
    metadata = _workbook.get_workbook_metadata(file_path)
    return json.dumps([s.model_dump() for s in metadata.sheets], indent=2)


@mcp.resource("excel://workbook/{file_path}/sheet/{sheet_name}/preview")
def resource_sheet_preview(file_path: str, sheet_name: str) -> str:
    """Preview the first 20 rows of a sheet as JSON."""
    data = _cell_ops.read_range(file_path, sheet_name, "A1", "Z20")
    return json.dumps(data, indent=2, default=str)


# ---------------------------------------------------------------------------
# Workbook Management (4 standalone tools — genuinely distinct)
# ---------------------------------------------------------------------------


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def get_workbook_metadata(file_path: str) -> WorkbookMetadata:
    """Get workbook metadata including sheet names, dimensions, active sheet, and named ranges."""
    return _workbook.get_workbook_metadata(file_path)


@mcp.tool()
def create_workbook(
    file_path: str, sheet_names: list[str] | None = None, sheet_name: str | None = None
) -> WorkbookCreatedResult:
    """Create a new .xlsx workbook.

    Optionally specify initial sheet names via sheet_names (list) or sheet_name (single).
    """
    return _workbook.create_workbook(file_path, sheet_names, sheet_name)


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def get_sheet_summary(file_path: str, sheet_name: str) -> SheetSummary:
    """Get sheet summary: name, row/col counts, headers, and used range."""
    return _workbook.get_sheet_summary(file_path, sheet_name)


@mcp.tool()
def write_multi_sheet(file_path: str, sheets: list[SheetDefinition]) -> WriteMultiSheetResult:
    """Create a new workbook with multiple named sheets, headers, data, and column widths in one call."""
    return _workbook.write_multi_sheet(file_path, sheets)


# ---------------------------------------------------------------------------
# Sheet Management (consolidated — rename/delete/copy)
# ---------------------------------------------------------------------------


@mcp.tool()
def sheet_management(
    action: Literal["rename", "delete", "copy", "hide", "unhide", "tab_color", "move"],
    file_path: str,
    sheet_name: str,
    new_name: str | None = None,
    color: str | None = None,
    offset: int | None = None,
) -> str | dict:
    """Manage worksheets within a workbook.

    action="rename": Rename a sheet. Requires: new_name.
    action="delete": DESTRUCTIVE. Delete a sheet. Raises if only sheet.
    action="copy": Copy a sheet. Requires: new_name for the copy.
    action="hide": Hide a sheet. Raises if it's the last visible sheet.
    action="unhide": Unhide a hidden sheet.
    action="tab_color": Set the tab colour. Requires: color (6-char hex, e.g. "FF0000"). Pass "000000" to clear.
    action="move": Reorder the sheet tab. Requires: offset (positive=right, negative=left).
    """
    if action == "rename":
        if not new_name:
            raise ValueError("new_name is required for action='rename'.")
        return _workbook.rename_sheet(file_path, sheet_name, new_name)
    if action == "delete":
        return _workbook.delete_sheet(file_path, sheet_name)
    if action == "copy":
        if not new_name:
            raise ValueError("new_name is required for action='copy'.")
        return _workbook.copy_sheet(file_path, sheet_name, new_name)
    if action == "hide":
        return _workbook.hide_sheet(file_path, sheet_name)
    if action == "unhide":
        return _workbook.unhide_sheet(file_path, sheet_name)
    if action == "tab_color":
        if not color:
            raise ValueError("color is required for action='tab_color'.")
        return _workbook.set_tab_color(file_path, sheet_name, color)
    if action == "move":
        if offset is None:
            raise ValueError("offset is required for action='move'.")
        return _workbook.move_sheet(file_path, sheet_name, offset)
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
) -> dict | ChunkReadResult:
    """Read cell data from a worksheet.

    mode="single": Read one cell. Requires: cell_ref. Optional: include_formula, include_metadata.
    mode="range": Read rectangular range. Requires: start_cell, end_cell.
        Optional: show_formula, show_style, output_format, max_cells.
    mode="chunked": Read large sheets in chunks. Optional: start_row, chunk_size.
    """
    if mode == "single":
        if cell_ref is None:
            raise ValueError("cell_ref is required for mode='single'.")
        return _cell_ops.read_cell(file_path, sheet_name, cell_ref, include_formula, include_metadata)
    if mode == "range":
        if start_cell is None:
            raise ValueError("start_cell is required for mode='range'.")
        if end_cell is None:
            raise ValueError("end_cell is required for mode='range'.")
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
    mode: Literal["single", "range", "series", "merge", "unmerge"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    value: str | int | float | bool | None = None,
    start_cell: str | None = None,
    data: list[list[str | int | float | bool | datetime | date | None]] | None = None,
    series_type: str = "number",
    count: int | None = None,
    step: float | str = 1,
    direction: str = "down",
    start_value: float | int | None = None,
    range_string: str | None = None,
) -> str | dict:
    """Write data to cells.

    mode="single": Write one cell. Requires: cell_ref, value.
    mode="range": Write 2D array. Requires: start_cell, data.
    mode="series": Fill series. Requires: start_cell, count. Optional: series_type, step, direction, start_value.
    mode="merge": Merge cells. Requires: range_string (e.g. 'A1:D1').
    mode="unmerge": Unmerge cells. Requires: range_string (e.g. 'A1:D1').
    """
    if mode == "single":
        if cell_ref is None:
            raise ValueError("cell_ref is required for mode='single'.")
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
    if mode == "merge":
        if not range_string:
            raise ValueError("range_string is required for mode='merge'.")
        return _cell_ops.merge_cells(file_path, sheet_name, range_string)
    if mode == "unmerge":
        if not range_string:
            raise ValueError("range_string is required for mode='unmerge'.")
        return _cell_ops.unmerge_cells(file_path, sheet_name, range_string)
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
    paste_values_only: bool = False,
) -> str:
    """Copy cells from source range to destination range within same or across sheets."""
    return _cell_ops.copy_range(
        file_path, source_sheet, source_range, dest_sheet, dest_range, copy_values, copy_styles, paste_values_only
    )


@mcp.tool()
def find_replace(
    file_path: str,
    sheet_name: str,
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    match_entire_cell: bool = False,
    search_formulas: bool = False,
) -> str:
    """Find and replace text across a worksheet. Set match_case for case-sensitive search."""
    return json.dumps(
        _cell_ops.find_replace(
            file_path,
            sheet_name,
            find_text,
            replace_text,
            match_case,
            match_entire_cell,
            search_formulas,
        )
    )


@mcp.tool()
def transpose_range(
    file_path: str,
    sheet_name: str,
    source_range: str,
    target_cell: str,
    source_sheet: str | None = None,
    paste_values_only: bool = False,
) -> str:
    """Transpose (swap rows and columns) of source_range and write starting at target_cell."""
    return json.dumps(
        _cell_ops.transpose_range(
            file_path,
            sheet_name,
            source_range,
            target_cell,
            source_sheet,
            paste_values_only,
        )
    )


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
    horizontal_alignment: HorizontalAlignment | None = None,
    vertical_alignment: VerticalAlignment | None = None,
    wrap_text: bool = False,
    border_style: BorderStyle | None = None,
    border_color: str | None = None,
    font_name: str | None = None,
    underline: str | None = None,
    strikethrough: bool = False,
    text_rotation: int | None = None,
    indent: int | None = None,
    shrink_to_fit: bool = False,
    top_border_style: BorderStyle | None = None,
    bottom_border_style: BorderStyle | None = None,
    left_border_style: BorderStyle | None = None,
    right_border_style: BorderStyle | None = None,
    number_format_preset: str | None = None,
    preserve_existing: bool = False,
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
        number_format_preset,
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
        preserve_existing,
    )


@mcp.tool()
def auto_fit_columns(file_path: str, sheet_name: str) -> str:
    """Auto-fit all column widths based on content length."""
    return _formatting.auto_fit_columns(file_path, sheet_name)


@mcp.tool()
def copy_cell_format(
    file_path: str,
    sheet_name: str,
    source_cell: str,
    target_range: str,
) -> str:
    """Copy all formatting from source_cell and apply it to every cell in target_range."""
    return json.dumps(_formatting.copy_cell_format(file_path, sheet_name, source_cell, target_range))


@mcp.tool()
def clear_cell_format(
    file_path: str,
    sheet_name: str,
    range_str: str,
) -> str:
    """Reset all formatting on cells in range_str without touching values."""
    return json.dumps(_formatting.clear_cell_format(file_path, sheet_name, range_str))


@mcp.tool()
def apply_named_style(
    file_path: str,
    sheet_name: str,
    range_str: str,
    style_name: str,
) -> str:
    """Apply a named built-in Excel style (e.g. 'Good', 'Bad', 'Neutral', 'Heading 1') to a cell range."""
    return json.dumps(_formatting.apply_named_style(file_path, sheet_name, range_str, style_name))


# ---------------------------------------------------------------------------
# Formula Write (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def formula_write(
    action: Literal["set", "batch", "fill", "auto_sum"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    formula: str | None = None,
    is_array: bool = False,
    target_range: str | None = None,
    formulas: dict[str, str] | None = None,
    source_range: str | None = None,
) -> str:
    """Write formulas to cells.

    action="set": Set a formula on a cell. Requires: cell_ref, formula. Optional: is_array, target_range.
    action="batch": Set multiple formulas. Requires: formulas (dict of cell_ref -> formula).
    action="fill": Drag-fill a formula. Requires: cell_ref (source), target_range.
    action="auto_sum": Insert =SUM() formula. Requires: cell_ref (destination). Optional: source_range.
    """
    if action == "set":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='set'.")
        if formula is None:
            raise ValueError("formula is required for action='set'.")
        return _formulas.set_formula(file_path, sheet_name, cell_ref, formula, is_array, target_range)
    if action == "batch":
        if formulas is None:
            raise ValueError("formulas is required for action='batch'.")
        return _formulas.set_formulas_batch(file_path, sheet_name, formulas)
    if action == "fill":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='fill'.")
        if target_range is None:
            raise ValueError("target_range is required for action='fill'.")
        return json.dumps(_cell_ops.fill_formula(file_path, sheet_name, cell_ref, target_range))
    if action == "auto_sum":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='auto_sum'.")
        return json.dumps(_cell_ops.auto_sum(file_path, sheet_name, cell_ref, source_range))
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
) -> dict | list[FormulaErrorInfo] | list[str] | list[FormulaInfo]:
    """Audit and inspect formulas. All actions are read-only.

    action="value": Get cached display value. Requires: cell_ref.
    action="errors": Find error cells (#VALUE!, #REF!, etc.). Optional: cell_range to limit scope.
    action="precedents": Cells feeding into cell_ref. Requires: cell_ref.
    action="dependents": Cells referencing cell_ref. Requires: cell_ref.
    action="list": List all formulas in the sheet.
    """
    if action == "value":
        if not cell_ref:
            raise ValueError("cell_ref is required for action='value'.")
        return _formulas.get_formula_value(file_path, sheet_name, cell_ref)
    if action == "errors":
        return _formulas.get_formula_errors(file_path, sheet_name, cell_range)
    if action == "precedents":
        if not cell_ref:
            raise ValueError("cell_ref is required for action='precedents'.")
        return _formulas.get_formula_precedents(file_path, sheet_name, cell_ref)
    if action == "dependents":
        if not cell_ref:
            raise ValueError("cell_ref is required for action='dependents'.")
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
        if file_path is None:
            raise ValueError("file_path is required for action='preview'.")
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
        if file_path is None:
            raise ValueError("file_path is required for action='to_csv'.")
        if output_path is None:
            raise ValueError("output_path is required for action='to_csv'.")
        return _csv_ops.xlsx_to_csv(file_path, sheet_name, output_path, delimiter, encoding)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Conditional Formatting (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def conditional_format(
    action: Literal["apply", "highlight", "formula_rule", "remove", "top_bottom", "above_below_average"],
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
) -> str:
    """Conditional formatting operations.

    action="apply": Apply color_scale/2_color_scale/data_bar/icon_set. Requires: cell_range, format_type.
    action="highlight": Highlight rule based on operator. Requires: cell_range, operator, formula.
    action="formula_rule": Custom formula-based rule. Requires: cell_range, formula.
    action="remove": DESTRUCTIVE. Remove rules. Optional: cell_range (omit to clear all).
    action="top_bottom": Top/bottom N rule. Requires: cell_range. Optional: is_top, rank, percent, bg_color, font_color.
    action="above_below_average": Above/below average rule. Requires: cell_range. Optional: is_above, equal_average, bg_color, font_color.
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
        return json.dumps(
            _cond_fmt.add_top_bottom_rule(
                file_path,
                sheet_name,
                cell_range,
                is_top,
                rank,
                percent,
                bg_color,
                font_color,
            )
        )
    if action == "above_below_average":
        if cell_range is None:
            raise ValueError("cell_range is required for action='above_below_average'.")
        return json.dumps(
            _cond_fmt.add_above_below_average_rule(
                file_path,
                sheet_name,
                cell_range,
                is_above,
                equal_average,
                bg_color,
                font_color,
            )
        )
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Tables (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def table(
    action: Literal["create", "list", "resize", "totals", "data", "convert_to_range"],
    file_path: str,
    sheet_name: str,
    table_name: str | None = None,
    data_range: str | None = None,
    style_name: str = "TableStyleMedium9",
    new_range: str | None = None,
    show_totals: bool | None = None,
    column_totals: dict[str, str] | None = None,
) -> str | list[TableInfo] | dict:
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
        if table_name is None:
            raise ValueError("table_name is required for action='resize'.")
        if new_range is None:
            raise ValueError("new_range is required for action='resize'.")
        return _tables.resize_table(file_path, sheet_name, table_name, new_range)
    if action == "totals":
        if table_name is None:
            raise ValueError("table_name is required for action='totals'.")
        if show_totals is None:
            raise ValueError("show_totals is required for action='totals'.")
        return _tables.set_table_totals_row(file_path, sheet_name, table_name, show_totals, column_totals)
    if action == "data":
        if table_name is None:
            raise ValueError("table_name is required for action='data'.")
        return _tables.get_table_data(file_path, sheet_name, table_name)
    if action == "convert_to_range":
        if table_name is None:
            raise ValueError("table_name is required for action='convert_to_range'.")
        return _tables.convert_table_to_range(file_path, sheet_name, table_name)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Data Validation (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
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
) -> str:
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
        return json.dumps(
            _data_val.add_formula_validation(
                file_path,
                sheet_name,
                cell_range,
                formula,
                error_title or "Invalid",
                error_message or "Value does not meet the criteria.",
                show_error,
            )
        )
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


# ---------------------------------------------------------------------------
# Charts (consolidated)
# ---------------------------------------------------------------------------


@mcp.tool()
def chart(
    action: Literal[
        "create", "delete", "list", "add_series", "set_axes", "trendline", "combo", "data_labels", "legend", "update"
    ],
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
    chart_title: str | None = None,
    show_value: bool = True,
    show_category: bool = False,
    show_series_name: bool = False,
    show_percentage: bool = False,
    label_position: str | None = None,
    show_legend: bool = True,
    legend_position: str | None = None,
) -> str | list[ChartInfo]:
    """Chart operations.

    action="create": Create chart. Requires: data_range. Optional: chart_type, target_cell, title, axes, style, size.
    action="delete": DESTRUCTIVE. Delete chart. Optional: chart_index (default 0).
    action="list": List all charts. Read-only.
    action="add_series": Add data series. Requires: chart_index, data_range.
    action="set_axes": Configure axes. Optional: chart_index, x/y titles, min/max, number_format, log_scale.
    action="trendline": Add trendline. Optional: chart_index, series_index, trendline_type, periods.
    action="combo": Create combo chart. Requires: data_range, bar_columns, line_columns.
    action="update": Update chart title, size, or anchor. Optional: chart_index, title, width, height, anchor_cell.
    """
    if action == "create":
        if data_range is None:
            raise ValueError("data_range is required for action='create'.")
        return _charts.create_chart(
            file_path,
            sheet_name,
            data_range,
            chart_type,
            target_cell,
            name or title or "",
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
        if chart_index is None:
            raise ValueError("chart_index is required for action='add_series'.")
        if data_range is None:
            raise ValueError("data_range is required for action='add_series'.")
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
    if action == "data_labels":
        if not chart_title:
            raise ValueError("chart_title is required for action='data_labels'.")
        return _charts.set_chart_data_labels(
            file_path,
            sheet_name,
            chart_title,
            show_value,
            show_category,
            show_series_name,
            show_percentage,
            label_position,
        )
    if action == "legend":
        if not chart_title:
            raise ValueError("chart_title is required for action='legend'.")
        return _charts.set_chart_legend(file_path, sheet_name, chart_title, show_legend, legend_position)
    if action == "update":
        return _charts.update_chart(
            file_path,
            sheet_name,
            chart_index or 0,
            title,
            width if width != 15 else None,
            height if height != 10 else None,
            anchor_cell if anchor_cell != "F1" else None,
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
) -> list[NamedRangeInfo] | str:
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
        if not name:
            raise ValueError("name is required for action='delete'.")
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
) -> str | CommentInfo | list[CommentInfo] | None:
    """Comment operations on cells.

    action="add": Add a comment. Requires: cell_ref, text. Optional: author.
    action="read": Read a comment. Read-only. Requires: cell_ref. Returns None if no comment.
    action="delete": DESTRUCTIVE. Delete a comment. Requires: cell_ref.
    action="list": List all comments in sheet. Read-only.
    """
    if action == "add":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='add'.")
        if text is None:
            raise ValueError("text is required for action='add'.")
        return _comments.add_comment(file_path, sheet_name, cell_ref, text, author)
    if action == "read":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='read'.")
        return _comments.read_comment(file_path, sheet_name, cell_ref)
    if action == "delete":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='delete'.")
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
) -> str | HyperlinkReadResult | list[HyperlinkInfo] | None:
    """Hyperlink operations.

    action="add": Add hyperlink. Requires: cell_ref, url. Optional: display_text, tooltip.
    action="read": Read hyperlink. Read-only. Requires: cell_ref. Returns None if none.
    action="delete": DESTRUCTIVE. Delete hyperlink. Requires: cell_ref.
    action="list": List all hyperlinks. Read-only.
    """
    if action == "add":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='add'.")
        if url is None:
            raise ValueError("url is required for action='add'.")
        return _hyperlinks.add_hyperlink(file_path, sheet_name, cell_ref, url, display_text, tooltip)
    if action == "read":
        if not cell_ref:
            raise ValueError("cell_ref is required for action='read'.")
        return _hyperlinks.read_hyperlink(file_path, sheet_name, cell_ref)
    if action == "delete":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='delete'.")
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
    cell_values: dict[str, dict[str, ScenarioCellValue]] | None = None,
    description: str = "",
) -> str | list[ScenarioInfo] | ScenarioApplyResult:
    """Scenario management for what-if analysis.

    action="add": Save a scenario. Requires: name, cell_values ({sheet: {cell: value}}). Optional: description.
    action="list": List all scenarios. Read-only.
    action="apply": DESTRUCTIVE. Apply scenario values to sheet. Requires: name.
    """
    if action == "add":
        if not name:
            raise ValueError("name is required for action='add'.")
        if cell_values is None:
            raise ValueError("cell_values is required for action='add'.")
        return _scenarios.add_scenario(file_path, name, cell_values, description)
    if action == "list":
        return _scenarios.list_scenarios(file_path)
    if action == "apply":
        if not name:
            raise ValueError("name is required for action='apply'.")
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
    sheet_name_b: str | None = None,
) -> dict:
    """Cross-file operations.

    action="aggregate": Aggregate a column across files. Requires: file_paths, column. Optional: operation.
    action="filter": Filter rows across files. Requires: file_paths, column, operator, value.
    action="validate": Cross-file consistency check. Requires: file_paths, key_column.
    action="compare": Compare two workbooks. Requires: file_a, file_b. Optional: sheet_name, sheet_name_b.
    """
    if action == "aggregate":
        if file_paths is None:
            raise ValueError("file_paths is required for action='aggregate'.")
        if column is None:
            raise ValueError("column is required for action='aggregate'.")
        return _multi_file.bulk_aggregate_multi_files(
            file_paths,
            column,
            operation,
            sheet_name,
            header_row,
            output_file,
        )
    if action == "filter":
        if file_paths is None:
            raise ValueError("file_paths is required for action='filter'.")
        if column is None:
            raise ValueError("column is required for action='filter'.")
        if operator is None:
            raise ValueError("operator is required for action='filter'.")
        if value is None:
            raise ValueError("value is required for action='filter'.")
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
        if file_paths is None:
            raise ValueError("file_paths is required for action='validate'.")
        if key_column is None:
            raise ValueError("key_column is required for action='validate'.")
        return _multi_file.validate_data_consistency(file_paths, key_column, check_columns, sheet_name, header_row)
    if action == "compare":
        if not file_a:
            raise ValueError("file_a is required for action='compare'.")
        if not file_b:
            raise ValueError("file_b is required for action='compare'.")
        return _multi_file.compare_workbooks(file_a, file_b, sheet_name, output_file, sheet_name_b)
    raise ValueError(f"Unknown action: {action}")


# ---------------------------------------------------------------------------
# Worksheet Operations (consolidated: view + cross-workbook)
# ---------------------------------------------------------------------------


@mcp.tool()
def worksheet_ops(
    action: Literal[
        "freeze",
        "auto_filter",
        "copy_range_across",
        "copy_sheet_across",
        "merge_workbooks",
        "insert_rows",
        "delete_rows",
        "insert_cols",
        "delete_cols",
        "set_print_area",
        "set_page_setup",
        "group_rows",
        "group_cols",
        "ungroup_rows",
        "ungroup_cols",
        "set_print_titles",
        "set_row_height",
        "set_col_width",
        "set_gridlines",
        "stack_sheets",
        "add_page_break",
        "remove_page_break",
    ],
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
    row: int | None = None,
    col: int | None = None,
    count: int | None = None,
    start_row: int | None = None,
    end_row: int | None = None,
    start_col: int | None = None,
    end_col: int | None = None,
    outline_level: int | None = None,
    hidden: bool | None = None,
    print_area: str | None = None,
    orientation: str | None = None,
    paper_size: int | None = None,
    fit_to_width: int | None = None,
    fit_to_height: int | None = None,
    title_rows: str | None = None,
    title_cols: str | None = None,
    rows: list[int] | None = None,
    height: float | None = None,
    cols_list: list[str] | None = None,
    width: float | None = None,
    show: bool = True,
    sheet_names: list[str] | None = None,
    dest_sheet: str | None = None,
    include_header: bool = True,
    output_path: str | None = None,
) -> str | dict:
    """Worksheet operations: freeze panes, auto filter, cross-sheet/workbook ops, row/col management, grouping, print.

    action="freeze": Freeze panes. Requires: file_path, sheet_name. Optional: cell_ref.
    action="auto_filter": Toggle auto-filter. Requires: file_path, sheet_name. Optional: cell_range, remove.
    action="copy_range_across": Copy range between sheets.
        Requires: file_path, source_sheet, source_range, target_sheet.
        Optional: target_start_cell (default "A1").
    action="copy_sheet_across": Copy sheet between workbooks. Requires: source_file, source_sheet, dest_file.
    action="merge_workbooks": Merge workbooks. Requires: source_files, output_file.
    action="insert_rows": Insert rows. Requires: file_path, sheet_name, row. Optional: count.
    action="delete_rows": Delete rows. Requires: file_path, sheet_name, row. Optional: count.
    action="insert_cols": Insert columns. Requires: file_path, sheet_name, col. Optional: count.
    action="delete_cols": Delete columns. Requires: file_path, sheet_name, col. Optional: count.
    action="set_print_area": Set print area. Requires: file_path, sheet_name, print_area.
    action="set_page_setup": Configure page setup. Requires: file_path, sheet_name.
        Optional: orientation, paper_size, fit_to_width, fit_to_height.
    action="group_rows": Group rows. Requires: file_path, sheet_name, start_row, end_row.
        Optional: outline_level, hidden.
    action="group_cols": Group columns. Requires: file_path, sheet_name, start_col, end_col.
        Optional: outline_level, hidden.
    action="ungroup_rows": Ungroup rows. Requires: file_path, sheet_name, start_row, end_row.
    action="ungroup_cols": Ungroup columns. Requires: file_path, sheet_name, start_col, end_col.
    action="add_page_break": Insert a manual page break. Requires: file_path, sheet_name.
      Optional: row (horizontal break before this row), col (vertical break before this column).
      At least one of row or col must be provided. row/col must be >= 2.
    action="remove_page_break": Remove a manual page break. Requires: file_path, sheet_name.
      Optional: row, col. If omitted for a type, all breaks of that type are cleared.
    """
    if action == "freeze":
        if file_path is None:
            raise ValueError("file_path is required for action='freeze'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='freeze'.")
        return _ws_ops.freeze_panes(file_path, sheet_name, cell_ref)
    if action == "auto_filter":
        if file_path is None:
            raise ValueError("file_path is required for action='auto_filter'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='auto_filter'.")
        if not remove and cell_range is None:
            raise ValueError("cell_range is required for action='auto_filter' when remove=False.")
        return _ws_ops.set_auto_filter(file_path, sheet_name, cell_range, remove)
    if action == "copy_range_across":
        if file_path is None:
            raise ValueError("file_path is required for action='copy_range_across'.")
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
    if action == "insert_rows":
        if file_path is None:
            raise ValueError("file_path is required for action='insert_rows'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='insert_rows'.")
        if row is None:
            raise ValueError("row is required for action='insert_rows'.")
        return _ws_ops.insert_rows(file_path, sheet_name, row, count or 1)
    if action == "delete_rows":
        if file_path is None:
            raise ValueError("file_path is required for action='delete_rows'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='delete_rows'.")
        if row is None:
            raise ValueError("row is required for action='delete_rows'.")
        return _ws_ops.delete_rows(file_path, sheet_name, row, count or 1)
    if action == "insert_cols":
        if file_path is None:
            raise ValueError("file_path is required for action='insert_cols'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='insert_cols'.")
        if col is None:
            raise ValueError("col is required for action='insert_cols'.")
        return _ws_ops.insert_cols(file_path, sheet_name, col, count or 1)
    if action == "delete_cols":
        if file_path is None:
            raise ValueError("file_path is required for action='delete_cols'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='delete_cols'.")
        if col is None:
            raise ValueError("col is required for action='delete_cols'.")
        return _ws_ops.delete_cols(file_path, sheet_name, col, count or 1)
    if action == "set_print_area":
        if file_path is None:
            raise ValueError("file_path is required for action='set_print_area'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='set_print_area'.")
        if not print_area:
            raise ValueError("print_area is required for action='set_print_area'.")
        return _ws_ops.set_print_area(file_path, sheet_name, print_area)
    if action == "set_page_setup":
        if file_path is None:
            raise ValueError("file_path is required for action='set_page_setup'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='set_page_setup'.")
        return _ws_ops.set_page_setup(
            file_path, sheet_name, orientation or "portrait", paper_size or 1, fit_to_width, fit_to_height
        )
    if action == "group_rows":
        if file_path is None:
            raise ValueError("file_path is required for action='group_rows'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='group_rows'.")
        if start_row is None:
            raise ValueError("start_row is required for action='group_rows'.")
        if end_row is None:
            raise ValueError("end_row is required for action='group_rows'.")
        return _ws_ops.group_rows(file_path, sheet_name, start_row, end_row, outline_level or 1, hidden or False)
    if action == "group_cols":
        if file_path is None:
            raise ValueError("file_path is required for action='group_cols'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='group_cols'.")
        if start_col is None:
            raise ValueError("start_col is required for action='group_cols'.")
        if end_col is None:
            raise ValueError("end_col is required for action='group_cols'.")
        return _ws_ops.group_cols(file_path, sheet_name, start_col, end_col, outline_level or 1, hidden or False)
    if action == "ungroup_rows":
        if file_path is None:
            raise ValueError("file_path is required for action='ungroup_rows'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='ungroup_rows'.")
        if start_row is None:
            raise ValueError("start_row is required for action='ungroup_rows'.")
        if end_row is None:
            raise ValueError("end_row is required for action='ungroup_rows'.")
        return _ws_ops.ungroup_rows(file_path, sheet_name, start_row, end_row)
    if action == "ungroup_cols":
        if file_path is None:
            raise ValueError("file_path is required for action='ungroup_cols'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='ungroup_cols'.")
        if start_col is None:
            raise ValueError("start_col is required for action='ungroup_cols'.")
        if end_col is None:
            raise ValueError("end_col is required for action='ungroup_cols'.")
        return _ws_ops.ungroup_cols(file_path, sheet_name, start_col, end_col)
    if action == "set_print_titles":
        if file_path is None:
            raise ValueError("file_path is required for action='set_print_titles'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='set_print_titles'.")
        return _ws_ops.set_print_titles(file_path, sheet_name, title_rows, title_cols)
    if action == "set_row_height":
        if file_path is None:
            raise ValueError("file_path is required for action='set_row_height'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='set_row_height'.")
        if rows is None:
            raise ValueError("rows is required for action='set_row_height'.")
        if height is None:
            raise ValueError("height is required for action='set_row_height'.")
        return _ws_ops.set_row_height(file_path, sheet_name, rows, height)
    if action == "set_col_width":
        if file_path is None:
            raise ValueError("file_path is required for action='set_col_width'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='set_col_width'.")
        if cols_list is None:
            raise ValueError("cols_list is required for action='set_col_width'.")
        if width is None:
            raise ValueError("width is required for action='set_col_width'.")
        return _ws_ops.set_col_width(file_path, sheet_name, cols_list, width)
    if action == "set_gridlines":
        if file_path is None:
            raise ValueError("file_path is required for action='set_gridlines'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='set_gridlines'.")
        return _ws_ops.set_gridlines(file_path, sheet_name, show)
    if action == "stack_sheets":
        if file_path is None:
            raise ValueError("file_path is required for action='stack_sheets'.")
        if sheet_names is None:
            raise ValueError("sheet_names is required for action='stack_sheets'.")
        if dest_sheet is None:
            raise ValueError("dest_sheet is required for action='stack_sheets'.")
        return _ws_ops.stack_sheets(file_path, sheet_names, dest_sheet, include_header, output_path)
    if action == "add_page_break":
        if file_path is None:
            raise ValueError("file_path is required for action='add_page_break'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='add_page_break'.")
        return _ws_ops.add_page_break(file_path, sheet_name, row, col)
    if action == "remove_page_break":
        if file_path is None:
            raise ValueError("file_path is required for action='remove_page_break'.")
        if sheet_name is None:
            raise ValueError("sheet_name is required for action='remove_page_break'.")
        return _ws_ops.remove_page_break(file_path, sheet_name, row, col)
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
def column_statistics(file_path: str, sheet_name: str, column: str, has_header: bool = True) -> ColumnStats:
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


@mcp.tool()
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


@mcp.tool()
def insert_subtotals(
    file_path: str,
    sheet_name: str,
    group_col: str,
    value_col: str,
    subtotal_func: int = 9,
    include_grand_total: bool = True,
) -> dict:
    """Insert SUBTOTAL formula rows after each group in a sorted sheet.

    subtotal_func: 1=AVERAGE, 2=COUNT, 3=COUNTA, 4=MAX, 5=MIN, 9=SUM.
    """
    return _analysis.insert_subtotals(file_path, sheet_name, group_col, value_col, subtotal_func, include_grand_total)


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
    date_freq: str | None = None,
) -> dict:
    """Create a pivot table and optionally write results to a sheet or file.

    date_freq: if set, groups datetime index columns by this period before pivoting.
      Common values: 'ME' (month-end), 'QE' (quarter-end), 'YE' (year-end), 'W' (weekly).
      Any valid pandas DateOffset alias is accepted.
    """
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
        date_freq,
    )


@mcp.tool()
def refresh_pivot_table(
    file_path: str,
    output_sheet: str,
    source_file_path: str | None = None,
    source_sheet: str | None = None,
) -> str:
    """Refresh a previously-created pivot table by re-running its stored definition."""
    return json.dumps(_pivot_etl.refresh_pivot_table(file_path, output_sheet, source_file_path, source_sheet))


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
    join_key: str | list[str] | None = None,
    how: str = "left",
    output_sheet: str | None = None,
    left_on: str | list[str] | None = None,
    right_on: str | list[str] | None = None,
) -> dict:
    """Merge two sheets like a SQL join (left, right, inner, outer).

    Use ``join_key`` when both sheets share the same column name(s).
    Use ``left_on`` / ``right_on`` to join on differently-named columns.
    """
    return _pivot_etl.merge_datasets(
        file_path,
        sheet1,
        sheet2,
        join_key,
        how,
        output_sheet,
        left_on=left_on,
        right_on=right_on,
    )


@mcp.tool()
def add_computed_column(
    file_path: str,
    sheet_name: str,
    new_column_name: str,
    expression: str,
    has_header: bool = True,
    column_type: str = "formula",
    source_col: str | None = None,
    window: int | None = None,
    rolling_func: str = "mean",
) -> str:
    """Add a computed column using a pandas-eval expression (e.g. 'Revenue - Cost').

    column_type='formula' (default): evaluate expression via pandas eval.
    column_type='cumsum': compute a running total of source_col. Requires: source_col.
    column_type='rolling': compute a rolling window aggregation of source_col.
      Requires: source_col, window (int). Optional: rolling_func ('mean' or 'sum', default 'mean').
    """
    return _pivot_etl.add_computed_column(
        file_path,
        sheet_name,
        new_column_name,
        expression,
        has_header,
        column_type,
        source_col,
        window,
        rolling_func,
    )


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
    operation: Literal["fv", "pv", "nper", "rate", "depreciation", "irr"],
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
    cash_flows: list[float] | None = None,
) -> dict:
    """Time value of money and depreciation calculations.

    operation="fv": Future value. Requires: rate, nper, pmt. Optional: pv, when.
    operation="pv": Present value. Requires: rate, nper, pmt. Optional: fv, when.
    operation="nper": Number of periods. Requires: rate, pmt, pv. Optional: fv, when.
    operation="rate": Interest rate. Requires: nper, pmt, pv. Optional: fv, when, guess.
    operation="depreciation": Asset depreciation. Requires: cost, salvage, life. Optional: method, period.
      method values: "sln" / "straight_line", "syd" / "sum_of_years" / "sum_of_years_digits",
                     "ddb" / "double_declining" / "double_declining_balance". Default: "sln".
    operation="irr": Internal Rate of Return. Requires: cash_flows (list of floats,
      first value typically negative as initial investment). Returns irr and irr_percent.
    """
    if operation == "fv":
        if rate is None:
            raise ValueError("rate is required for operation='fv'.")
        if nper is None:
            raise ValueError("nper is required for operation='fv'.")
        if pmt is None:
            raise ValueError("pmt is required for operation='fv'.")
        return _financial.calculate_fv(rate, nper, pmt, pv, when)
    if operation == "pv":
        if rate is None:
            raise ValueError("rate is required for operation='pv'.")
        if nper is None:
            raise ValueError("nper is required for operation='pv'.")
        if pmt is None:
            raise ValueError("pmt is required for operation='pv'.")
        return _financial.calculate_pv(rate, nper, pmt, fv, when)
    if operation == "nper":
        if rate is None:
            raise ValueError("rate is required for operation='nper'.")
        if pmt is None:
            raise ValueError("pmt is required for operation='nper'.")
        return _financial.calculate_nper(rate, pmt, pv, fv, when)
    if operation == "rate":
        if nper is None:
            raise ValueError("nper is required for operation='rate'.")
        if pmt is None:
            raise ValueError("pmt is required for operation='rate'.")
        return _financial.calculate_rate(nper, pmt, pv, fv, when, guess)
    if operation == "depreciation":
        if cost is None:
            raise ValueError("cost is required for operation='depreciation'.")
        if salvage is None:
            raise ValueError("salvage is required for operation='depreciation'.")
        if life is None:
            raise ValueError("life is required for operation='depreciation'.")
        return _financial.calculate_depreciation(cost, salvage, life, method, period)
    if operation == "irr":
        if cash_flows is None:
            raise ValueError("cash_flows is required for operation='irr'.")
        return _financial.calculate_irr(cash_flows)
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
    fill_value: str | None = None,
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
        fill_value,
    )


@mcp.tool(annotations=ToolAnnotations(destructiveHint=True))
def parse_date_column(
    file_path: str,
    sheet_name: str,
    column: str,
    output_column: str | None = None,
    output_format: str = "%Y-%m-%d",
    dayfirst: bool = False,
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
    output_sheet: str = "Regression Output",
    output_file: str | None = None,
) -> RegressionResult:
    """Run OLS linear regression and return coefficients, R-squared, and residuals.

    If output_file is provided, results are written to that file instead of file_path.
    """
    return _statistical.run_regression(
        file_path,
        sheet_name,
        y_column,
        x_columns,
        output_sheet=output_sheet,
        output_file=output_file,
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
    method: str = "simple",
    seasonal_periods: int | None = None,
    forecast_steps: int = 0,
) -> dict:
    """Apply exponential smoothing to a time series and write the result to a new column.

    method: "simple" (pandas EWM), "holt" (Holt linear trend), "holt_winters" (Holt-Winters seasonal).
    seasonal_periods: required for holt_winters (e.g. 12 for monthly data).
    forecast_steps: number of out-of-sample steps to forecast.
    """
    return _statistical.run_exponential_smoothing(
        file_path,
        sheet_name,
        column,
        alpha,
        new_column_name,
        header_row,
        method=method,
        seasonal_periods=seasonal_periods,
        forecast_steps=forecast_steps,
        output_file=output_file,
    )


# ---------------------------------------------------------------------------
# Solver / Optimization (1 standalone tool)
# ---------------------------------------------------------------------------


@mcp.tool()
def run_solver(
    file_path: str,
    sheet_name: str,
    objective_expression: str,
    variable_cells: dict[str, list[float]],
    constraints: list[dict] | None = None,
    maximize: bool = False,
    tolerance: float = 1e-6,
    max_iterations: int = 1000,
) -> SolverResult:
    """Multi-variable constrained optimization using scipy.

    objective_expression: arithmetic expression using cell refs (e.g. "B2 * B3 - B4").
    variable_cells: {cell_ref: [lower_bound, upper_bound]} dict.
    constraints: list of {"expression": str, "type": "ineq"|"eq"} dicts.
    maximize: True to maximise instead of minimise.
    """
    normalised_cells: dict[str, tuple[float, float]] = {k: (v[0], v[1]) for k, v in variable_cells.items()}
    return _solver.run_solver(
        file_path,
        sheet_name,
        objective_expression,
        normalised_cells,
        constraints,
        maximize,
        tolerance,
        max_iterations,
    )


# ---------------------------------------------------------------------------
# Data Profiling (standalone)
# ---------------------------------------------------------------------------


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def profile_data(
    file_path: str,
    sheet: str | None = None,
    data_range: str | None = None,
) -> dict:
    """Profile data in a worksheet, returning column statistics, types, null counts, and sample values."""
    return _analysis.profile_data(file_path, sheet, data_range)


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def value_counts(
    file_path: str,
    sheet_name: str,
    column: str,
    normalize: bool = False,
    top_n: int | None = None,
    dropna: bool = True,
    has_header: bool = True,
) -> dict:
    """Return a full value-frequency table for a column.

    column: header name of the column to count.
    normalize: if True, return proportions (0–1) instead of raw counts.
    top_n: if given, return only the top N most-frequent values.
    dropna: if True (default), exclude null values from counts.
    Returns: {"column", "total_rows", "normalize", "counts": [{"value", "count"}, ...]}.
    """
    return _analysis.value_counts(file_path, sheet_name, column, normalize, top_n, dropna, has_header)


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
def correlation_matrix(
    file_path: str,
    sheet_name: str,
    columns: list[str] | None = None,
    output_sheet: str | None = None,
    output_file: str | None = None,
    header_row: int = 1,
) -> dict:
    """Compute a Pearson correlation matrix for numeric columns.

    columns: list of column names to include. If None, all numeric columns are used.
    output_sheet: if given, writes the matrix to this sheet (created if absent).
    output_file: target file for output; defaults to file_path.
    Returns: {"columns": [...], "matrix": [[float, ...], ...]}.
    """
    return _statistical.correlation_matrix(file_path, sheet_name, columns, output_sheet, output_file, header_row)


# ---------------------------------------------------------------------------
# Image Insertion (standalone)
# ---------------------------------------------------------------------------


@mcp.tool()
def insert_image(
    file_path: str,
    sheet: str,
    image_path: str,
    cell: str,
    width: int | None = None,
    height: int | None = None,
) -> dict:
    """Insert an image into a worksheet at the specified cell."""
    return _images.insert_image(file_path, sheet, image_path, cell, width, height)


# ---------------------------------------------------------------------------
# Custom Code Execution (standalone)
# ---------------------------------------------------------------------------


@mcp.tool()
def execute_custom_code(
    file_path: str,
    code: str,
    sheet: str | None = None,
    output_file: str | None = None,
) -> dict:
    """Execute custom Python/pandas code against an Excel file in a sandboxed environment.

    Use this for operations not covered by other tools.
    The code has access to 'df' (the DataFrame), 'pd' (pandas), and 'np' (numpy).
    Set 'result' variable to return data.
    """
    return _custom_code.execute_custom_code(file_path, code, sheet, output_file)


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
