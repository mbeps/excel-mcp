from __future__ import annotations

import json
from logging import Logger

from mcp.server.fastmcp import FastMCP

from mcp_server.tools.analysis import (
    aggregate_data as _aggregate_data,
)
from mcp_server.tools.analysis import (
    calculate_correlation as _calculate_correlation,
)
from mcp_server.tools.analysis import (
    calculate_percentiles as _calculate_percentiles,
)
from mcp_server.tools.analysis import (
    column_statistics as _column_statistics,
)
from mcp_server.tools.analysis import (
    create_histogram as _create_histogram,
)
from mcp_server.tools.analysis import (
    extract_unique_values as _extract_unique_values,
)
from mcp_server.tools.analysis import (
    filter_data as _filter_data,
)
from mcp_server.tools.analysis import (
    find_duplicates as _find_duplicates,
)
from mcp_server.tools.analysis import (
    profile_data as _profile_data,
)
from mcp_server.tools.analysis import (
    rank_data as _rank_data,
)
from mcp_server.tools.analysis import (
    sample_data as _sample_data,
)
from mcp_server.tools.analysis import (
    search_replace as _search_replace,
)
from mcp_server.tools.analysis import (
    sort_data as _sort_data,
)
from mcp_server.tools.analysis import (
    transpose_data as _transpose_data,
)
from mcp_server.tools.cell_ops import (
    clear_range as _clear_range,
)
from mcp_server.tools.cell_ops import (
    read_cell as _read_cell,
)
from mcp_server.tools.cell_ops import (
    read_cell_detailed as _read_cell_detailed,
)
from mcp_server.tools.cell_ops import (
    read_file_chunked as _read_file_chunked,
)
from mcp_server.tools.cell_ops import (
    read_range as _read_range,
)
from mcp_server.tools.cell_ops import (
    write_cell as _write_cell,
)
from mcp_server.tools.cell_ops import (
    write_range as _write_range,
)
from mcp_server.tools.charts import (
    create_chart as _create_chart,
)
from mcp_server.tools.charts import (
    delete_chart as _delete_chart,
)
from mcp_server.tools.charts import (
    list_charts as _list_charts,
)
from mcp_server.tools.charts import (
    update_chart_properties as _update_chart_properties,
)
from mcp_server.tools.comments import (
    add_comment as _add_comment,
)
from mcp_server.tools.comments import (
    delete_comment as _delete_comment,
)
from mcp_server.tools.comments import (
    list_comments as _list_comments,
)
from mcp_server.tools.comments import (
    read_comment as _read_comment,
)
from mcp_server.tools.conditional_formatting import (
    add_above_average_rule as _add_above_average_rule,
)
from mcp_server.tools.conditional_formatting import (
    add_duplicate_rule as _add_duplicate_rule,
)
from mcp_server.tools.conditional_formatting import (
    add_formula_rule as _add_formula_rule,
)
from mcp_server.tools.conditional_formatting import (
    add_highlight_rule as _add_highlight_rule,
)
from mcp_server.tools.conditional_formatting import (
    add_top_bottom_rule as _add_top_bottom_rule,
)
from mcp_server.tools.conditional_formatting import (
    apply_conditional_formatting as _apply_conditional_formatting,
)
from mcp_server.tools.conditional_formatting import (
    list_conditional_formats as _list_conditional_formats,
)
from mcp_server.tools.conditional_formatting import (
    remove_conditional_formatting as _remove_conditional_formatting,
)
from mcp_server.tools.csv_ops import (
    csv_to_xlsx as _csv_to_xlsx,
)
from mcp_server.tools.csv_ops import (
    read_csv_preview as _read_csv_preview,
)
from mcp_server.tools.csv_ops import (
    xlsx_to_csv as _xlsx_to_csv,
)
from mcp_server.tools.data_validation import (
    add_date_validation as _add_date_validation,
)
from mcp_server.tools.data_validation import (
    add_dropdown_validation as _add_dropdown_validation,
)
from mcp_server.tools.data_validation import (
    add_formula_validation as _add_formula_validation,
)
from mcp_server.tools.data_validation import (
    add_numeric_validation as _add_numeric_validation,
)
from mcp_server.tools.data_validation import (
    add_text_length_validation as _add_text_length_validation,
)
from mcp_server.tools.data_validation import (
    list_validations as _list_validations,
)
from mcp_server.tools.data_validation import (
    remove_validation as _remove_validation,
)
from mcp_server.tools.doc_properties import (
    get_document_properties as _get_document_properties,
)
from mcp_server.tools.doc_properties import (
    protect_workbook as _protect_workbook,
)
from mcp_server.tools.doc_properties import (
    set_calculation_mode as _set_calculation_mode,
)
from mcp_server.tools.doc_properties import (
    set_document_properties as _set_document_properties,
)
from mcp_server.tools.doc_properties import (
    unprotect_workbook as _unprotect_workbook,
)
from mcp_server.tools.financial import (
    calculate_irr as _calculate_irr,
)
from mcp_server.tools.financial import (
    calculate_npv as _calculate_npv,
)
from mcp_server.tools.financial import (
    calculate_pmt as _calculate_pmt,
)
from mcp_server.tools.financial import (
    goal_seek as _goal_seek,
)
from mcp_server.tools.financial import (
    loan_amortization as _loan_amortization,
)
from mcp_server.tools.formatting import (
    auto_fit_columns as _auto_fit_columns,
)
from mcp_server.tools.formatting import (
    format_cells as _format_cells,
)
from mcp_server.tools.formatting import (
    get_cell_formatting as _get_cell_formatting,
)
from mcp_server.tools.formatting import (
    list_merged_ranges as _list_merged_ranges,
)
from mcp_server.tools.formatting import (
    merge_cells as _merge_cells,
)
from mcp_server.tools.formatting import (
    set_column_width as _set_column_width,
)
from mcp_server.tools.formatting import (
    set_row_height as _set_row_height,
)
from mcp_server.tools.formatting import (
    unmerge_cells as _unmerge_cells,
)
from mcp_server.tools.formulas import (
    set_array_formula as _set_array_formula,
)
from mcp_server.tools.formulas import (
    set_formula as _set_formula,
)
from mcp_server.tools.formulas import (
    set_formulas_batch as _set_formulas_batch,
)
from mcp_server.tools.formulas import (
    validate_formula_syntax as _validate_formula_syntax,
)
from mcp_server.tools.hyperlinks import (
    add_hyperlink as _add_hyperlink,
)
from mcp_server.tools.hyperlinks import (
    delete_hyperlink as _delete_hyperlink,
)
from mcp_server.tools.hyperlinks import (
    list_hyperlinks as _list_hyperlinks,
)
from mcp_server.tools.hyperlinks import (
    read_hyperlink as _read_hyperlink,
)
from mcp_server.tools.images import (
    delete_image as _delete_image,
)
from mcp_server.tools.images import (
    insert_image as _insert_image,
)
from mcp_server.tools.images import (
    list_images as _list_images,
)
from mcp_server.tools.named_ranges import (
    create_named_range as _create_named_range,
)
from mcp_server.tools.named_ranges import (
    delete_named_range as _delete_named_range,
)
from mcp_server.tools.named_ranges import (
    list_named_ranges as _list_named_ranges,
)
from mcp_server.tools.named_ranges import (
    update_named_range as _update_named_range,
)
from mcp_server.tools.pivot_etl import (
    add_computed_column as _add_computed_column,
)
from mcp_server.tools.pivot_etl import (
    create_pivot_table as _create_pivot_table,
)
from mcp_server.tools.pivot_etl import (
    deduplicate_data as _deduplicate_data,
)
from mcp_server.tools.pivot_etl import (
    merge_datasets as _merge_datasets,
)
from mcp_server.tools.pivot_etl import (
    unpivot_data as _unpivot_data,
)
from mcp_server.tools.protection import (
    protect_cells as _protect_cells,
)
from mcp_server.tools.protection import (
    protect_sheet as _protect_sheet,
)
from mcp_server.tools.protection import (
    unprotect_sheet as _unprotect_sheet,
)
from mcp_server.tools.row_col import (
    delete_cols as _delete_cols,
)
from mcp_server.tools.row_col import (
    delete_columns_by_letter as _delete_columns_by_letter,
)
from mcp_server.tools.row_col import (
    delete_rows as _delete_rows,
)
from mcp_server.tools.row_col import (
    insert_cols as _insert_cols,
)
from mcp_server.tools.row_col import (
    insert_columns_by_letter as _insert_columns_by_letter,
)
from mcp_server.tools.row_col import (
    insert_rows as _insert_rows,
)
from mcp_server.tools.tables import (
    create_table as _create_table,
)
from mcp_server.tools.tables import (
    list_tables as _list_tables,
)
from mcp_server.tools.workbook import (
    copy_sheet as _copy_sheet,
)
from mcp_server.tools.workbook import (
    create_workbook as _create_workbook,
)
from mcp_server.tools.workbook import (
    delete_sheet as _delete_sheet,
)
from mcp_server.tools.workbook import (
    get_sheet_summary as _get_sheet_summary,
)
from mcp_server.tools.workbook import (
    get_workbook_metadata as _get_workbook_metadata,
)
from mcp_server.tools.workbook import (
    list_sheets as _list_sheets,
)
from mcp_server.tools.workbook import (
    rename_sheet as _rename_sheet,
)
from mcp_server.tools.worksheet_ops import (
    freeze_panes as _freeze_panes,
)
from mcp_server.tools.worksheet_ops import (
    group_columns as _group_columns,
)
from mcp_server.tools.worksheet_ops import (
    group_rows as _group_rows,
)
from mcp_server.tools.worksheet_ops import (
    hide_columns as _hide_columns,
)
from mcp_server.tools.worksheet_ops import (
    hide_rows as _hide_rows,
)
from mcp_server.tools.worksheet_ops import (
    hide_sheet as _hide_sheet,
)
from mcp_server.tools.worksheet_ops import (
    move_sheet as _move_sheet,
)
from mcp_server.tools.worksheet_ops import (
    remove_auto_filter as _remove_auto_filter,
)
from mcp_server.tools.worksheet_ops import (
    set_auto_filter as _set_auto_filter,
)
from mcp_server.tools.worksheet_ops import (
    set_header_footer as _set_header_footer,
)
from mcp_server.tools.worksheet_ops import (
    set_page_margins as _set_page_margins,
)
from mcp_server.tools.worksheet_ops import (
    set_page_setup as _set_page_setup,
)
from mcp_server.tools.worksheet_ops import (
    set_print_area as _set_print_area,
)
from mcp_server.tools.worksheet_ops import (
    set_sheet_tab_color as _set_sheet_tab_color,
)
from mcp_server.tools.worksheet_ops import (
    set_zoom as _set_zoom,
)
from mcp_server.tools.worksheet_ops import (
    show_gridlines as _show_gridlines,
)
from mcp_server.tools.worksheet_ops import (
    unfreeze_panes as _unfreeze_panes,
)
from mcp_server.tools.worksheet_ops import (
    ungroup_columns as _ungroup_columns,
)
from mcp_server.tools.worksheet_ops import (
    ungroup_rows as _ungroup_rows,
)
from mcp_server.tools.worksheet_ops import (
    unhide_columns as _unhide_columns,
)
from mcp_server.tools.worksheet_ops import (
    unhide_rows as _unhide_rows,
)
from mcp_server.tools.worksheet_ops import (
    unhide_sheet as _unhide_sheet,
)
from mcp_server.utils.logger import configure_logging

mcp: FastMCP = FastMCP("excel-mcp-server")
logger: Logger = configure_logging("mcp_server.main")


# ---------------------------------------------------------------------------
# Resources
# ---------------------------------------------------------------------------


@mcp.resource("excel://workbook/{file_path}/sheets")
def resource_list_sheets(file_path: str) -> str:
    """List all sheets in a workbook as JSON."""
    return json.dumps(_list_sheets(file_path), indent=2)


@mcp.resource("excel://workbook/{file_path}/sheet/{sheet_name}/preview")
def resource_sheet_preview(file_path: str, sheet_name: str) -> str:
    """Preview the first 20 rows of a sheet as JSON."""
    data = _read_range(file_path, sheet_name, "A1", "Z20")
    return json.dumps(data, indent=2, default=str)


# ---------------------------------------------------------------------------
# --- Workbook Management ---
# ---------------------------------------------------------------------------


@mcp.tool()
def get_workbook_metadata(file_path: str) -> dict:
    """Get workbook metadata including sheet names, dimensions, active sheet, and named ranges."""
    return _get_workbook_metadata(file_path)


@mcp.tool()
def list_sheets(file_path: str) -> list[dict]:
    """List all sheets in a workbook with their names and dimensions."""
    return _list_sheets(file_path)


@mcp.tool()
def create_workbook(file_path: str, sheet_names: list[str] | None = None) -> dict:
    """Create a new .xlsx workbook. Optionally specify initial sheet names."""
    return _create_workbook(file_path, sheet_names)


@mcp.tool()
def get_sheet_summary(file_path: str, sheet_name: str) -> dict:
    """Get sheet summary: name, row/col counts, headers, and used range."""
    return _get_sheet_summary(file_path, sheet_name)


@mcp.tool()
def rename_sheet(file_path: str, old_name: str, new_name: str) -> str:
    """Rename a worksheet."""
    return _rename_sheet(file_path, old_name, new_name)


@mcp.tool()
def delete_sheet(file_path: str, sheet_name: str) -> str:
    """Delete a worksheet. Fails if it's the only sheet."""
    return _delete_sheet(file_path, sheet_name)


@mcp.tool()
def copy_sheet(file_path: str, source_sheet: str, new_name: str) -> str:
    """Copy a sheet within the same workbook under a new name."""
    return _copy_sheet(file_path, source_sheet, new_name)


# ---------------------------------------------------------------------------
# --- Named Ranges ---
# ---------------------------------------------------------------------------


@mcp.tool()
def list_named_ranges(file_path: str) -> list[dict]:
    """List all named ranges in a workbook with name, destination, and scope."""
    return _list_named_ranges(file_path)


@mcp.tool()
def create_named_range(file_path: str, name: str, destination: str, scope: str = "workbook") -> str:
    """Create a named range. Destination like 'Sheet1!$A$1:$A$10'. Scope is 'workbook' or a sheet name."""
    return _create_named_range(file_path, name, destination, scope)


@mcp.tool()
def delete_named_range(file_path: str, name: str) -> str:
    """Delete a named range by name."""
    return _delete_named_range(file_path, name)


@mcp.tool()
def update_named_range(file_path: str, name: str, new_destination: str) -> str:
    """Update the destination of an existing named range."""
    return _update_named_range(file_path, name, new_destination)


# ---------------------------------------------------------------------------
# --- Comments ---
# ---------------------------------------------------------------------------


@mcp.tool()
def add_comment(file_path: str, sheet_name: str, cell_ref: str, text: str, author: str = "Excel MCP") -> str:
    """Add a comment/note to a cell."""
    return _add_comment(file_path, sheet_name, cell_ref, text, author)


@mcp.tool()
def read_comment(file_path: str, sheet_name: str, cell_ref: str) -> dict | None:
    """Read a comment from a cell. Returns dict with text and author, or None."""
    return _read_comment(file_path, sheet_name, cell_ref)


@mcp.tool()
def delete_comment(file_path: str, sheet_name: str, cell_ref: str) -> str:
    """Delete a comment from a cell."""
    return _delete_comment(file_path, sheet_name, cell_ref)


@mcp.tool()
def list_comments(file_path: str, sheet_name: str) -> list[dict]:
    """List all comments in a sheet with cell reference, text, and author."""
    return _list_comments(file_path, sheet_name)


# ---------------------------------------------------------------------------
# --- Hyperlinks ---
# ---------------------------------------------------------------------------


@mcp.tool()
def add_hyperlink(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
    url: str,
    display_text: str | None = None,
    tooltip: str | None = None,
) -> str:
    """Add a hyperlink to a cell."""
    return _add_hyperlink(file_path, sheet_name, cell_ref, url, display_text, tooltip)


@mcp.tool()
def read_hyperlink(file_path: str, sheet_name: str, cell_ref: str) -> dict | None:
    """Read hyperlink from a cell. Returns dict with target, location, tooltip, or None."""
    return _read_hyperlink(file_path, sheet_name, cell_ref)


@mcp.tool()
def delete_hyperlink(file_path: str, sheet_name: str, cell_ref: str) -> str:
    """Delete a hyperlink from a cell."""
    return _delete_hyperlink(file_path, sheet_name, cell_ref)


@mcp.tool()
def list_hyperlinks(file_path: str, sheet_name: str) -> list[dict]:
    """List all hyperlinks in a sheet."""
    return _list_hyperlinks(file_path, sheet_name)


# ---------------------------------------------------------------------------
# --- Worksheet Operations ---
# ---------------------------------------------------------------------------


@mcp.tool()
def freeze_panes(file_path: str, sheet_name: str, cell_ref: str) -> str:
    """Freeze panes at the given cell (e.g. 'B2' freezes row 1 and column A)."""
    return _freeze_panes(file_path, sheet_name, cell_ref)


@mcp.tool()
def unfreeze_panes(file_path: str, sheet_name: str) -> str:
    """Remove freeze panes from a sheet."""
    return _unfreeze_panes(file_path, sheet_name)


@mcp.tool()
def set_auto_filter(file_path: str, sheet_name: str, cell_range: str) -> str:
    """Enable auto-filter on a range (e.g. 'A1:E1')."""
    return _set_auto_filter(file_path, sheet_name, cell_range)


@mcp.tool()
def remove_auto_filter(file_path: str, sheet_name: str) -> str:
    """Remove auto-filter from a sheet."""
    return _remove_auto_filter(file_path, sheet_name)


@mcp.tool()
def hide_rows(file_path: str, sheet_name: str, start_row: int, end_row: int) -> str:
    """Hide rows from start_row to end_row (1-based, inclusive)."""
    return _hide_rows(file_path, sheet_name, start_row, end_row)


@mcp.tool()
def unhide_rows(file_path: str, sheet_name: str, start_row: int, end_row: int) -> str:
    """Unhide rows from start_row to end_row (1-based, inclusive)."""
    return _unhide_rows(file_path, sheet_name, start_row, end_row)


@mcp.tool()
def hide_columns(file_path: str, sheet_name: str, start_col: str, end_col: str) -> str:
    """Hide columns from start_col to end_col (letters like 'A', 'D')."""
    return _hide_columns(file_path, sheet_name, start_col, end_col)


@mcp.tool()
def unhide_columns(file_path: str, sheet_name: str, start_col: str, end_col: str) -> str:
    """Unhide columns from start_col to end_col."""
    return _unhide_columns(file_path, sheet_name, start_col, end_col)


@mcp.tool()
def group_rows(
    file_path: str,
    sheet_name: str,
    start_row: int,
    end_row: int,
    outline_level: int = 1,
    hidden: bool = False,
) -> str:
    """Group rows (outline) at a specified outline level."""
    return _group_rows(file_path, sheet_name, start_row, end_row, outline_level, hidden)


@mcp.tool()
def ungroup_rows(file_path: str, sheet_name: str, start_row: int, end_row: int) -> str:
    """Ungroup rows by resetting outline level to 0."""
    return _ungroup_rows(file_path, sheet_name, start_row, end_row)


@mcp.tool()
def group_columns(
    file_path: str,
    sheet_name: str,
    start_col: str,
    end_col: str,
    outline_level: int = 1,
    hidden: bool = False,
) -> str:
    """Group columns (outline) at a specified outline level."""
    return _group_columns(file_path, sheet_name, start_col, end_col, outline_level, hidden)


@mcp.tool()
def ungroup_columns(file_path: str, sheet_name: str, start_col: str, end_col: str) -> str:
    """Ungroup columns by resetting outline level to 0."""
    return _ungroup_columns(file_path, sheet_name, start_col, end_col)


@mcp.tool()
def set_sheet_tab_color(file_path: str, sheet_name: str, color: str) -> str:
    """Set sheet tab color (hex like 'FF0000' for red)."""
    return _set_sheet_tab_color(file_path, sheet_name, color)


@mcp.tool()
def hide_sheet(file_path: str, sheet_name: str, very_hidden: bool = False) -> str:
    """Hide a sheet. Use very_hidden=True for extra hiding (not accessible from Excel UI)."""
    return _hide_sheet(file_path, sheet_name, very_hidden)


@mcp.tool()
def unhide_sheet(file_path: str, sheet_name: str) -> str:
    """Unhide a hidden sheet."""
    return _unhide_sheet(file_path, sheet_name)


@mcp.tool()
def move_sheet(file_path: str, sheet_name: str, offset: int) -> str:
    """Move sheet by offset positions. Positive=right, negative=left."""
    return _move_sheet(file_path, sheet_name, offset)


@mcp.tool()
def set_zoom(file_path: str, sheet_name: str, zoom_scale: int) -> str:
    """Set sheet zoom level (10-400)."""
    return _set_zoom(file_path, sheet_name, zoom_scale)


@mcp.tool()
def show_gridlines(file_path: str, sheet_name: str, show: bool = True) -> str:
    """Show or hide gridlines on a sheet."""
    return _show_gridlines(file_path, sheet_name, show)


@mcp.tool()
def set_print_area(file_path: str, sheet_name: str, cell_range: str) -> str:
    """Set the print area for a sheet."""
    return _set_print_area(file_path, sheet_name, cell_range)


@mcp.tool()
def set_page_setup(
    file_path: str,
    sheet_name: str,
    orientation: str = "portrait",
    paper_size: int = 1,
    fit_to_width: int | None = None,
    fit_to_height: int | None = None,
) -> str:
    """Configure page setup (orientation, paper size, fit-to-page)."""
    return _set_page_setup(file_path, sheet_name, orientation, paper_size, fit_to_width, fit_to_height)


@mcp.tool()
def set_page_margins(
    file_path: str,
    sheet_name: str,
    top: float = 0.75,
    bottom: float = 0.75,
    left: float = 0.7,
    right: float = 0.7,
    header: float = 0.3,
    footer: float = 0.3,
) -> str:
    """Set page margins in inches."""
    return _set_page_margins(file_path, sheet_name, top, bottom, left, right, header, footer)


@mcp.tool()
def set_header_footer(
    file_path: str,
    sheet_name: str,
    header_center: str | None = None,
    header_left: str | None = None,
    header_right: str | None = None,
    footer_center: str | None = None,
    footer_left: str | None = None,
    footer_right: str | None = None,
) -> str:
    """Set page header and footer text."""
    return _set_header_footer(
        file_path,
        sheet_name,
        header_center,
        header_left,
        header_right,
        footer_center,
        footer_left,
        footer_right,
    )


# ---------------------------------------------------------------------------
# --- Cell Operations ---
# ---------------------------------------------------------------------------


@mcp.tool()
def read_cell(file_path: str, sheet_name: str, cell_ref: str) -> dict:
    """Read a single cell's value and data type."""
    return _read_cell(file_path, sheet_name, cell_ref)


@mcp.tool()
def write_cell(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
    value: str | int | float | bool | None,
) -> str:
    """Write a value to a single cell."""
    return _write_cell(file_path, sheet_name, cell_ref, value)


@mcp.tool()
def read_range(file_path: str, sheet_name: str, start_cell: str, end_cell: str) -> dict:
    """Read a rectangular range and return rows as a list of lists."""
    return _read_range(file_path, sheet_name, start_cell, end_cell)


@mcp.tool()
def write_range(file_path: str, sheet_name: str, start_cell: str, data: list[list]) -> str:
    """Write a 2D array of values starting from start_cell."""
    return _write_range(file_path, sheet_name, start_cell, data)


@mcp.tool()
def clear_range(file_path: str, sheet_name: str, start_cell: str, end_cell: str) -> str:
    """Clear all values in a rectangular cell range."""
    return _clear_range(file_path, sheet_name, start_cell, end_cell)


@mcp.tool()
def read_file_chunked(
    file_path: str,
    sheet_name: str,
    start_row: int = 0,
    chunk_size: int = 1000,
) -> dict:
    """Read a sheet in chunks for large files. Returns rows, chunk info, and has_more flag."""
    return _read_file_chunked(file_path, sheet_name, start_row, chunk_size)


@mcp.tool()
def read_cell_detailed(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
    data_only: bool = False,
) -> dict:
    """Read a cell with full detail: value, type, formula, comment, hyperlink, merge status."""
    return _read_cell_detailed(file_path, sheet_name, cell_ref, data_only)


# ---------------------------------------------------------------------------
# --- Row & Column Operations ---
# ---------------------------------------------------------------------------


@mcp.tool()
def insert_rows(file_path: str, sheet_name: str, row_index: int, count: int = 1) -> str:
    """Insert empty rows at a 1-based position."""
    return _insert_rows(file_path, sheet_name, row_index, count)


@mcp.tool()
def delete_rows(file_path: str, sheet_name: str, row_index: int, count: int = 1) -> str:
    """Delete rows starting at a 1-based position."""
    return _delete_rows(file_path, sheet_name, row_index, count)


@mcp.tool()
def insert_cols(file_path: str, sheet_name: str, col_index: int, count: int = 1) -> str:
    """Insert empty columns at a 1-based position."""
    return _insert_cols(file_path, sheet_name, col_index, count)


@mcp.tool()
def delete_cols(file_path: str, sheet_name: str, col_index: int, count: int = 1) -> str:
    """Delete columns starting at a 1-based position."""
    return _delete_cols(file_path, sheet_name, col_index, count)


@mcp.tool()
def insert_columns_by_letter(file_path: str, sheet_name: str, column: str, count: int = 1) -> str:
    """Insert columns before the specified column letter (e.g. 'C')."""
    return _insert_columns_by_letter(file_path, sheet_name, column, count)


@mcp.tool()
def delete_columns_by_letter(file_path: str, sheet_name: str, column: str, count: int = 1) -> str:
    """Delete columns starting at the specified column letter (e.g. 'C')."""
    return _delete_columns_by_letter(file_path, sheet_name, column, count)


# ---------------------------------------------------------------------------
# --- Formatting ---
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
    return _format_cells(
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
def set_column_width(file_path: str, sheet_name: str, column: str, width: float) -> str:
    """Set the width of a column by its letter (e.g. 'A')."""
    return _set_column_width(file_path, sheet_name, column, width)


@mcp.tool()
def set_row_height(file_path: str, sheet_name: str, row: int, height: float) -> str:
    """Set the height of a row (1-based)."""
    return _set_row_height(file_path, sheet_name, row, height)


@mcp.tool()
def merge_cells(file_path: str, sheet_name: str, cell_range: str) -> str:
    """Merge a range of cells (e.g. 'A1:C1')."""
    return _merge_cells(file_path, sheet_name, cell_range)


@mcp.tool()
def unmerge_cells(file_path: str, sheet_name: str, cell_range: str) -> str:
    """Unmerge a previously merged cell range."""
    return _unmerge_cells(file_path, sheet_name, cell_range)


@mcp.tool()
def auto_fit_columns(file_path: str, sheet_name: str) -> str:
    """Auto-fit all column widths based on content length."""
    return _auto_fit_columns(file_path, sheet_name)


@mcp.tool()
def get_cell_formatting(file_path: str, sheet_name: str, cell_ref: str) -> dict:
    """Read all formatting properties of a cell (font, fill, border, alignment, number format, protection)."""
    return _get_cell_formatting(file_path, sheet_name, cell_ref)


@mcp.tool()
def list_merged_ranges(file_path: str, sheet_name: str) -> list[str]:
    """List all merged cell ranges in a sheet."""
    return _list_merged_ranges(file_path, sheet_name)


# ---------------------------------------------------------------------------
# --- Formulas ---
# ---------------------------------------------------------------------------


@mcp.tool()
def set_formula(file_path: str, sheet_name: str, cell_ref: str, formula: str) -> str:
    """Set an Excel formula on a cell. Prepends '=' if missing."""
    return _set_formula(file_path, sheet_name, cell_ref, formula)


@mcp.tool()
def set_array_formula(file_path: str, sheet_name: str, target_range: str, formula: str) -> str:
    """Apply an array (CSE) formula to a range."""
    return _set_array_formula(file_path, sheet_name, target_range, formula)


@mcp.tool()
def validate_formula_syntax(formula: str) -> dict:
    """Validate an Excel formula's syntax and return token info."""
    return _validate_formula_syntax(formula)


@mcp.tool()
def set_formulas_batch(file_path: str, sheet_name: str, formulas: dict[str, str]) -> str:
    """Set multiple formulas at once. Keys are cell refs, values are formulas."""
    return _set_formulas_batch(file_path, sheet_name, formulas)


# ---------------------------------------------------------------------------
# --- CSV Operations ---
# ---------------------------------------------------------------------------


@mcp.tool()
def read_csv_preview(file_path: str, rows: int = 10, delimiter: str = ",") -> dict:
    """Preview the first N rows of a CSV file."""
    return _read_csv_preview(file_path, rows, delimiter)


@mcp.tool()
def csv_to_xlsx(csv_path: str, xlsx_path: str, sheet_name: str = "Sheet1", delimiter: str = ",") -> str:
    """Convert a CSV file to .xlsx format."""
    return _csv_to_xlsx(csv_path, xlsx_path, sheet_name, delimiter)


@mcp.tool()
def xlsx_to_csv(file_path: str, sheet_name: str, output_path: str, delimiter: str = ",") -> str:
    """Export a worksheet to CSV."""
    return _xlsx_to_csv(file_path, sheet_name, output_path, delimiter)


# ---------------------------------------------------------------------------
# --- Conditional Formatting ---
# ---------------------------------------------------------------------------


@mcp.tool()
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
    """Apply conditional formatting (color_scale, data_bar, or icon_set) to a range."""
    return _apply_conditional_formatting(
        file_path,
        sheet_name,
        cell_range,
        format_type,
        start_color,
        mid_color,
        end_color,
        bar_color,
        icon_style,
    )


@mcp.tool()
def add_highlight_rule(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    operator: str,
    formula: str,
    font_color: str = "9C0006",
    bg_color: str = "FFC7CE",
) -> str:
    """Add a cell-highlight conditional formatting rule based on a comparison operator."""
    return _add_highlight_rule(
        file_path,
        sheet_name,
        cell_range,
        operator,
        formula,
        font_color,
        bg_color,
    )


@mcp.tool()
def remove_conditional_formatting(file_path: str, sheet_name: str) -> str:
    """Remove all conditional formatting rules from a sheet."""
    return _remove_conditional_formatting(file_path, sheet_name)


@mcp.tool()
def add_formula_rule(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    formula: str,
    font_color: str = "9C0006",
    bg_color: str = "FFC7CE",
) -> str:
    """Add a formula-based conditional formatting rule."""
    return _add_formula_rule(file_path, sheet_name, cell_range, formula, font_color, bg_color)


@mcp.tool()
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
    return _add_top_bottom_rule(file_path, sheet_name, cell_range, rank, bottom, percent, font_color, bg_color)


@mcp.tool()
def add_above_average_rule(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    above: bool = True,
    font_color: str = "9C0006",
    bg_color: str = "FFC7CE",
) -> str:
    """Add an above/below average conditional formatting rule."""
    return _add_above_average_rule(file_path, sheet_name, cell_range, above, font_color, bg_color)


@mcp.tool()
def add_duplicate_rule(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    font_color: str = "9C0006",
    bg_color: str = "FFC7CE",
) -> str:
    """Add duplicate values highlighting conditional formatting rule."""
    return _add_duplicate_rule(file_path, sheet_name, cell_range, font_color, bg_color)


@mcp.tool()
def list_conditional_formats(file_path: str, sheet_name: str) -> list[dict]:
    """List all conditional formatting rules on a sheet."""
    return _list_conditional_formats(file_path, sheet_name)


# ---------------------------------------------------------------------------
# --- Tables ---
# ---------------------------------------------------------------------------


@mcp.tool()
def create_table(
    file_path: str,
    sheet_name: str,
    data_range: str,
    table_name: str,
    style_name: str = "TableStyleMedium9",
) -> str:
    """Create a native Excel table (ListObject) from a data range."""
    return _create_table(file_path, sheet_name, data_range, table_name, style_name)


@mcp.tool()
def list_tables(file_path: str, sheet_name: str) -> list[dict]:
    """List all tables in a sheet with name, ref, and style."""
    return _list_tables(file_path, sheet_name)


# ---------------------------------------------------------------------------
# --- Data Validation ---
# ---------------------------------------------------------------------------


@mcp.tool()
def add_dropdown_validation(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    options: list[str],
    allow_blank: bool = True,
) -> str:
    """Add a dropdown-list data validation to a cell range."""
    return _add_dropdown_validation(file_path, sheet_name, cell_range, options, allow_blank)


@mcp.tool()
def add_numeric_validation(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    operator: str,
    value1: float,
    value2: float | None = None,
    allow_blank: bool = True,
) -> str:
    """Add numeric data validation (whole or decimal) to a cell range."""
    return _add_numeric_validation(
        file_path,
        sheet_name,
        cell_range,
        operator,
        value1,
        value2,
        allow_blank,
    )


@mcp.tool()
def add_formula_validation(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    formula: str,
    allow_blank: bool = True,
) -> str:
    """Add custom formula-based data validation to a cell range."""
    return _add_formula_validation(file_path, sheet_name, cell_range, formula, allow_blank)


@mcp.tool()
def list_validations(file_path: str, sheet_name: str) -> list[dict]:
    """List all data validation rules on a sheet."""
    return _list_validations(file_path, sheet_name)


@mcp.tool()
def add_date_validation(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    operator: str = "greaterThan",
    date1: str = "",
    date2: str | None = None,
    allow_blank: bool = True,
) -> str:
    """Add date validation to a cell range."""
    return _add_date_validation(file_path, sheet_name, cell_range, operator, date1, date2, allow_blank)


@mcp.tool()
def add_text_length_validation(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    operator: str = "lessThanOrEqual",
    length1: int = 100,
    length2: int | None = None,
    allow_blank: bool = True,
) -> str:
    """Add text length validation to a cell range."""
    return _add_text_length_validation(file_path, sheet_name, cell_range, operator, length1, length2, allow_blank)


@mcp.tool()
def remove_validation(file_path: str, sheet_name: str, cell_range: str) -> str:
    """Remove data validations from a specific range."""
    return _remove_validation(file_path, sheet_name, cell_range)


# ---------------------------------------------------------------------------
# --- Protection ---
# ---------------------------------------------------------------------------


@mcp.tool()
def protect_sheet(
    file_path: str,
    sheet_name: str,
    password: str | None = None,
    allow_formatting_cells: bool = False,
    allow_formatting_columns: bool = False,
    allow_formatting_rows: bool = False,
    allow_insert_columns: bool = False,
    allow_insert_rows: bool = False,
    allow_delete_columns: bool = False,
    allow_delete_rows: bool = False,
    allow_sort: bool = False,
    allow_filter: bool = False,
) -> str:
    """Enable sheet protection with configurable permissions."""
    return _protect_sheet(
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


@mcp.tool()
def unprotect_sheet(file_path: str, sheet_name: str, password: str | None = None) -> str:
    """Remove sheet protection."""
    return _unprotect_sheet(file_path, sheet_name, password)


@mcp.tool()
def protect_cells(
    file_path: str,
    sheet_name: str,
    locked_range: str,
    unlocked_ranges: list[str] | None = None,
) -> str:
    """Lock specific cells and optionally unlock others. Requires sheet protection to take effect."""
    return _protect_cells(file_path, sheet_name, locked_range, unlocked_ranges)


# ---------------------------------------------------------------------------
# --- Charts ---
# ---------------------------------------------------------------------------


@mcp.tool()
def create_chart(
    file_path: str,
    sheet_name: str,
    data_range: str,
    chart_type: str = "column",
    target_cell: str = "E1",
    title: str = "",
    x_axis_title: str = "",
    y_axis_title: str = "",
    style: int = 10,
    width: float = 15,
    height: float = 10,
) -> str:
    """Create a native Excel chart from a data range.

    Supported types: bar, column, line, pie, scatter, area, radar, doughnut, bubble, stock.
    """
    return _create_chart(
        file_path,
        sheet_name,
        data_range,
        chart_type,
        target_cell,
        title,
        x_axis_title,
        y_axis_title,
        style,
        width,
        height,
    )


@mcp.tool()
def delete_chart(file_path: str, sheet_name: str, chart_index: int = 0) -> str:
    """Delete a chart from a sheet by its zero-based index."""
    return _delete_chart(file_path, sheet_name, chart_index)


@mcp.tool()
def list_charts(file_path: str, sheet_name: str) -> list[dict]:
    """List all charts on a sheet with title, type, and position."""
    return _list_charts(file_path, sheet_name)


@mcp.tool()
def update_chart_properties(
    file_path: str,
    sheet_name: str,
    chart_index: int = 0,
    title: str | None = None,
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    legend_position: str | None = None,
    show_legend: bool | None = None,
    style: int | None = None,
) -> str:
    """Update properties (title, axes, legend, style) of an existing chart."""
    return _update_chart_properties(
        file_path,
        sheet_name,
        chart_index,
        title,
        x_axis_title,
        y_axis_title,
        legend_position,
        show_legend,
        style,
    )


# ---------------------------------------------------------------------------
# --- Images ---
# ---------------------------------------------------------------------------


@mcp.tool()
def insert_image(
    file_path: str,
    sheet_name: str,
    image_path: str,
    cell_ref: str,
    width: int | None = None,
    height: int | None = None,
) -> str:
    """Insert an image into a sheet at the specified cell."""
    return _insert_image(file_path, sheet_name, image_path, cell_ref, width, height)


@mcp.tool()
def list_images(file_path: str, sheet_name: str) -> list[dict]:
    """List all images in a sheet with index, dimensions, and anchor."""
    return _list_images(file_path, sheet_name)


@mcp.tool()
def delete_image(file_path: str, sheet_name: str, image_index: int = 0) -> str:
    """Delete an image by its zero-based index from a sheet."""
    return _delete_image(file_path, sheet_name, image_index)


# ---------------------------------------------------------------------------
# --- Document Properties & Workbook Settings ---
# ---------------------------------------------------------------------------


@mcp.tool()
def get_document_properties(file_path: str) -> dict:
    """Get workbook document properties (title, creator, dates, etc.)."""
    return _get_document_properties(file_path)


@mcp.tool()
def set_document_properties(
    file_path: str,
    title: str | None = None,
    creator: str | None = None,
    description: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    category: str | None = None,
    company: str | None = None,
) -> str:
    """Set workbook document properties. Only specified fields are updated."""
    return _set_document_properties(file_path, title, creator, description, subject, keywords, category, company)


@mcp.tool()
def protect_workbook(
    file_path: str,
    password: str | None = None,
    lock_structure: bool = True,
    lock_windows: bool = False,
) -> str:
    """Protect workbook structure and/or windows with an optional password."""
    return _protect_workbook(file_path, password, lock_structure, lock_windows)


@mcp.tool()
def unprotect_workbook(file_path: str) -> str:
    """Remove workbook-level protection."""
    return _unprotect_workbook(file_path)


@mcp.tool()
def set_calculation_mode(file_path: str, mode: str = "auto") -> str:
    """Set workbook calculation mode: 'auto', 'manual', or 'autoNoTable'."""
    return _set_calculation_mode(file_path, mode)


# ---------------------------------------------------------------------------
# --- Data Analysis ---
# ---------------------------------------------------------------------------


@mcp.tool()
def filter_data(
    file_path: str,
    sheet_name: str,
    column: str,
    operator: str,
    value: str | int | float,
    has_header: bool = True,
) -> dict:
    """Filter rows by a column condition. Operators: ==, !=, >, <, >=, <=, contains, startswith, endswith."""
    return _filter_data(file_path, sheet_name, column, operator, value, has_header)


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
    return _sort_data(file_path, sheet_name, sort_by, column, ascending, has_header)


@mcp.tool()
def column_statistics(file_path: str, sheet_name: str, column: str, has_header: bool = True) -> dict:
    """Compute descriptive statistics (mean, median, std, min, max, sum) for a numeric column."""
    return _column_statistics(file_path, sheet_name, column, has_header)


@mcp.tool()
def aggregate_data(
    file_path: str,
    sheet_name: str,
    group_by: str,
    value_column: str,
    operation: str,
    has_header: bool = True,
) -> dict:
    """Group by a column and aggregate (sum, mean, count, min, max, median, std)."""
    return _aggregate_data(
        file_path,
        sheet_name,
        group_by,
        value_column,
        operation,
        has_header,
    )


@mcp.tool()
def find_duplicates(file_path: str, sheet_name: str, columns: list[str], has_header: bool = True) -> dict:
    """Find duplicate rows based on specified columns."""
    return _find_duplicates(file_path, sheet_name, columns, has_header)


@mcp.tool()
def profile_data(file_path: str, sheet_name: str) -> dict:
    """Comprehensive data profiling: column types, missing values, duplicates, and summary stats."""
    return _profile_data(file_path, sheet_name)


@mcp.tool()
def search_replace(
    file_path: str,
    sheet_name: str,
    search_value: str,
    replace_value: str,
    cell_range: str | None = None,
) -> str:
    """Find and replace values in a sheet or specific range."""
    return _search_replace(file_path, sheet_name, search_value, replace_value, cell_range)


@mcp.tool()
def calculate_correlation(
    file_path: str,
    sheet_name: str,
    columns: list[str],
    has_header: bool = True,
) -> dict:
    """Calculate correlation matrix between specified numeric columns."""
    return _calculate_correlation(file_path, sheet_name, columns, has_header)


@mcp.tool()
def rank_data(
    file_path: str,
    sheet_name: str,
    column: str,
    method: str = "dense",
    ascending: bool = True,
    has_header: bool = True,
) -> dict:
    """Rank values in a column using dense, min, max, first, or average method."""
    return _rank_data(file_path, sheet_name, column, method, ascending, has_header)


@mcp.tool()
def calculate_percentiles(
    file_path: str,
    sheet_name: str,
    column: str,
    percentiles: list[float] | None = None,
    has_header: bool = True,
) -> dict:
    """Calculate percentile values for a numeric column."""
    return _calculate_percentiles(file_path, sheet_name, column, percentiles, has_header)


@mcp.tool()
def sample_data(
    file_path: str,
    sheet_name: str,
    n: int = 10,
    fraction: float | None = None,
    random_state: int | None = None,
    has_header: bool = True,
) -> dict:
    """Random sample of rows from a sheet."""
    return _sample_data(file_path, sheet_name, n, fraction, random_state, has_header)


@mcp.tool()
def create_histogram(
    file_path: str,
    sheet_name: str,
    column: str,
    bins: int = 10,
    has_header: bool = True,
) -> dict:
    """Create histogram bins for a numeric column."""
    return _create_histogram(file_path, sheet_name, column, bins, has_header)


@mcp.tool()
def transpose_data(
    file_path: str,
    sheet_name: str,
    output_sheet: str | None = None,
    has_header: bool = True,
) -> str:
    """Transpose data (rows to columns and vice versa)."""
    return _transpose_data(file_path, sheet_name, output_sheet, has_header)


@mcp.tool()
def extract_unique_values(
    file_path: str,
    sheet_name: str,
    column: str,
    has_header: bool = True,
) -> dict:
    """Extract unique values from a column."""
    return _extract_unique_values(file_path, sheet_name, column, has_header)


# ---------------------------------------------------------------------------
# --- Pivot & ETL ---
# ---------------------------------------------------------------------------


@mcp.tool()
def create_pivot_table(
    file_path: str,
    sheet_name: str,
    index_cols: list[str],
    value_cols: list[str],
    aggfunc: str = "sum",
    output_sheet: str | None = None,
    output_file: str | None = None,
) -> dict:
    """Create a pivot table and optionally write results to a sheet or file."""
    return _create_pivot_table(
        file_path,
        sheet_name,
        index_cols,
        value_cols,
        aggfunc,
        output_sheet,
        output_file,
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
    return _unpivot_data(file_path, sheet_name, id_vars, value_vars, var_name, value_name)


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
    return _merge_datasets(file_path, sheet1, sheet2, join_key, how, output_sheet)


@mcp.tool()
def add_computed_column(
    file_path: str,
    sheet_name: str,
    new_column_name: str,
    expression: str,
    has_header: bool = True,
) -> str:
    """Add a computed column using a pandas-eval expression (e.g. 'Revenue - Cost')."""
    return _add_computed_column(file_path, sheet_name, new_column_name, expression, has_header)


@mcp.tool()
def deduplicate_data(
    file_path: str,
    sheet_name: str,
    columns: list[str] | None = None,
    keep: str = "first",
) -> str:
    """Remove duplicate rows from a sheet. keep: 'first', 'last', or False."""
    return _deduplicate_data(file_path, sheet_name, columns, keep)


# ---------------------------------------------------------------------------
# --- Financial ---
# ---------------------------------------------------------------------------


@mcp.tool()
def calculate_npv(discount_rate: float, cash_flows: list[float]) -> dict:
    """Calculate Net Present Value given a discount rate and cash flows."""
    return _calculate_npv(discount_rate, cash_flows)


@mcp.tool()
def calculate_irr(cash_flows: list[float]) -> dict:
    """Calculate Internal Rate of Return for a series of cash flows."""
    return _calculate_irr(cash_flows)


@mcp.tool()
def calculate_pmt(rate: float, nper: int, pv: float, fv: float = 0) -> dict:
    """Calculate periodic payment for a loan or annuity."""
    return _calculate_pmt(rate, nper, pv, fv)


@mcp.tool()
def goal_seek(
    target_formula: str,
    target_value: float,
    initial_guess: float = 1.0,
) -> dict:
    """Find x such that target_formula(x) = target_value. Formula uses 'x' with basic math ops."""
    return _goal_seek(target_formula, target_value, initial_guess)


@mcp.tool()
def loan_amortization(
    principal: float,
    annual_rate: float,
    years: int,
    payments_per_year: int = 12,
) -> dict:
    """Generate a loan amortization schedule with payment breakdown."""
    return _loan_amortization(principal, annual_rate, years, payments_per_year)


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
