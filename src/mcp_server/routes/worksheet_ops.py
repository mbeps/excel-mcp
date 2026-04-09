from __future__ import annotations

from typing import Literal

from mcp.types import ToolAnnotations

import mcp_server.tools.worksheet_ops as _ws_ops

__all__ = [
    "worksheet_view",
    "worksheet_structure",
    "worksheet_print",
    "worksheet_transfer",
]


def worksheet_view(
    action: Literal["freeze", "auto_filter", "set_gridlines"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    cell_range: str | None = None,
    remove: bool = False,
    show: bool = True,
) -> str:
    """Toggle view-related settings such as freeze panes, auto-filter and gridlines.

    Args:
        action: "freeze", "auto_filter", or "set_gridlines".
        file_path: Workbook path.
        sheet_name: Worksheet name.
        cell_ref: For freeze, the pane freeze cell (omit to unfreeze).
        cell_range: For auto_filter set/remove.
        remove: For auto_filter, remove filter when True.
        show: For set_gridlines, show/hide gridlines.

    Returns:
        str: Result message.

    Notes:
        - Usually non-destructive aside from toggling UI settings stored in workbook.
    """
    if action == "freeze":
        return _ws_ops.freeze_panes(file_path, sheet_name, cell_ref)
    if action == "auto_filter":
        if not remove and cell_range is None:
            raise ValueError("cell_range is required for action='auto_filter' when remove=False.")
        return _ws_ops.set_auto_filter(file_path, sheet_name, cell_range, remove)
    if action == "set_gridlines":
        return _ws_ops.set_gridlines(file_path, sheet_name, show)
    raise ValueError(f"Unknown action: {action}")


def worksheet_structure(
    action: Literal[
        "insert_rows",
        "insert_cols",
        "delete_rows",
        "delete_cols",
        "set_row_height",
        "set_col_width",
        "group_rows",
        "group_cols",
        "ungroup_rows",
        "ungroup_cols",
    ],
    file_path: str,
    sheet_name: str,
    row: int | None = None,
    col: int | None = None,
    count: int | None = None,
    start_row: int | None = None,
    end_row: int | None = None,
    start_col: int | None = None,
    end_col: int | None = None,
    outline_level: int | None = None,
    hidden: bool | None = None,
    rows: list[int] | None = None,
    height: float | None = None,
    cols_list: list[str] | None = None,
    width: float | None = None,
) -> str:
    """Perform row/column insert/delete, grouping, and size adjustments.

    Args:
        action: One of insert/delete/group/ungroup/set_row_height/set_col_width.
        file_path: Workbook path.
        sheet_name: Worksheet name.
        row, col, count, start_row, end_row, start_col, end_col: Position arguments.
        outline_level, hidden, rows, cols_list, height, width: Operation-specific params.

    Returns:
        str: Result message.

    Notes:
        - Many actions are destructive (delete_rows/delete_cols) — document irreversible effects.
    """
    if action == "insert_rows":
        if row is None:
            raise ValueError("row is required for action='insert_rows'.")
        return _ws_ops.insert_rows(file_path, sheet_name, row, count or 1)
    if action == "delete_rows":
        if row is None:
            raise ValueError("row is required for action='delete_rows'.")
        return _ws_ops.delete_rows(file_path, sheet_name, row, count or 1)
    if action == "insert_cols":
        if col is None:
            raise ValueError("col is required for action='insert_cols'.")
        return _ws_ops.insert_cols(file_path, sheet_name, col, count or 1)
    if action == "delete_cols":
        if col is None:
            raise ValueError("col is required for action='delete_cols'.")
        return _ws_ops.delete_cols(file_path, sheet_name, col, count or 1)
    if action == "set_row_height":
        if rows is None:
            raise ValueError("rows is required for action='set_row_height'.")
        if height is None:
            raise ValueError("height is required for action='set_row_height'.")
        return _ws_ops.set_row_height(file_path, sheet_name, rows, height)
    if action == "set_col_width":
        if cols_list is None:
            raise ValueError("cols_list is required for action='set_col_width'.")
        if width is None:
            raise ValueError("width is required for action='set_col_width'.")
        return _ws_ops.set_col_width(file_path, sheet_name, cols_list, width)
    if action == "group_rows":
        if start_row is None:
            raise ValueError("start_row is required for action='group_rows'.")
        if end_row is None:
            raise ValueError("end_row is required for action='group_rows'.")
        return _ws_ops.group_rows(file_path, sheet_name, start_row, end_row, outline_level or 1, hidden or False)
    if action == "group_cols":
        if start_col is None:
            raise ValueError("start_col is required for action='group_cols'.")
        if end_col is None:
            raise ValueError("end_col is required for action='group_cols'.")
        return _ws_ops.group_cols(file_path, sheet_name, start_col, end_col, outline_level or 1, hidden or False)
    if action == "ungroup_rows":
        if start_row is None:
            raise ValueError("start_row is required for action='ungroup_rows'.")
        if end_row is None:
            raise ValueError("end_row is required for action='ungroup_rows'.")
        return _ws_ops.ungroup_rows(file_path, sheet_name, start_row, end_row)
    if action == "ungroup_cols":
        if start_col is None:
            raise ValueError("start_col is required for action='ungroup_cols'.")
        if end_col is None:
            raise ValueError("end_col is required for action='ungroup_cols'.")
        return _ws_ops.ungroup_cols(file_path, sheet_name, start_col, end_col)
    raise ValueError(f"Unknown action: {action}")


def worksheet_print(
    action: Literal["set_print_area", "set_page_setup", "set_print_titles", "add_page_break", "remove_page_break"],
    file_path: str,
    sheet_name: str,
    print_area: str | None = None,
    orientation: str | None = None,
    paper_size: int | None = None,
    fit_to_width: int | None = None,
    fit_to_height: int | None = None,
    title_rows: str | None = None,
    title_cols: str | None = None,
    row: int | None = None,
    col: int | None = None,
) -> str:
    """Configure print areas, page setup, print titles and manual page breaks.

    Args:
        action: "set_print_area", "set_page_setup", "set_print_titles", "add_page_break", "remove_page_break".
        file_path, sheet_name: Workbook and worksheet.
        print_area, orientation, paper_size, fit_to_width, fit_to_height, title_rows, title_cols, row, col: params.

    Returns:
        str: Result message.

    Notes:
        - Mostly metadata changes to the sheet's print settings.
    """
    if action == "set_print_area":
        if not print_area:
            raise ValueError("print_area is required for action='set_print_area'.")
        return _ws_ops.set_print_area(file_path, sheet_name, print_area)
    if action == "set_page_setup":
        return _ws_ops.set_page_setup(
            file_path, sheet_name, orientation or "portrait", paper_size or 1, fit_to_width, fit_to_height
        )
    if action == "set_print_titles":
        return _ws_ops.set_print_titles(file_path, sheet_name, title_rows, title_cols)
    if action == "add_page_break":
        return _ws_ops.add_page_break(file_path, sheet_name, row, col)
    if action == "remove_page_break":
        return _ws_ops.remove_page_break(file_path, sheet_name, row, col)
    raise ValueError(f"Unknown action: {action}")


def worksheet_transfer(
    action: Literal["copy_range_across", "copy_sheet_across", "merge_workbooks", "stack_sheets"],
    file_path: str | None = None,
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
    sheet_names: list[str] | None = None,
    dest_sheet: str | None = None,
    include_header: bool = True,
    output_path: str | None = None,
) -> str | dict:
    """Cross-sheet and cross-workbook copy/merge/stack operations.

    Args:
        action: "copy_range_across", "copy_sheet_across", "merge_workbooks", or "stack_sheets".
        file_path: Source workbook path (varies by action).
        source_sheet, source_range: Source identifiers for copy.
        target_sheet, target_start_cell: Destination identifiers.
        copy_values, copy_styles: Copy options.
        source_file, dest_file, dest_sheet_name: For copy_sheet_across.
        source_files, output_file, conflict_strategy: For merge_workbooks.
        sheet_names, dest_sheet, include_header, output_path: For stack_sheets.

    Returns:
        str or dict: Result metadata.

    Notes:
        - These operations can be expensive (many file opens) and destructive. Document conflict resolution strategies.
    """
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
    if action == "stack_sheets":
        if file_path is None:
            raise ValueError("file_path is required for action='stack_sheets'.")
        if sheet_names is None:
            raise ValueError("sheet_names is required for action='stack_sheets'.")
        if dest_sheet is None:
            raise ValueError("dest_sheet is required for action='stack_sheets'.")
        return _ws_ops.stack_sheets(file_path, sheet_names, dest_sheet, include_header, output_path)
    raise ValueError(f"Unknown action: {action}")


def register(mcp) -> None:
    """Register worksheet-operations tools on *mcp*."""
    mcp.tool()(worksheet_view)
    mcp.tool(annotations=ToolAnnotations(destructiveHint=True))(worksheet_structure)
    mcp.tool()(worksheet_print)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=False))(worksheet_transfer)
