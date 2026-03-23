"""Worksheet-level operations: freeze panes, filters, visibility, grouping, zoom, print setup."""

from __future__ import annotations

from logging import Logger

from openpyxl.worksheet.page import PageMargins

from mcp_server.utils.excel_helpers import (
    col_letter_to_index,
    get_sheet,
    index_to_col_letter,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def freeze_panes(file_path: str, sheet_name: str, cell_ref: str) -> str:
    """Freeze panes at the given cell reference (e.g. 'B2' freezes row 1 and column A)."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.freeze_panes = cell_ref
        save_workbook_safe(wb, file_path)
        logger.info("Froze panes at %s in '%s' of %s", cell_ref, sheet_name, file_path)
        return f"Panes frozen at '{cell_ref}' in sheet '{sheet_name}'."
    finally:
        wb.close()


def unfreeze_panes(file_path: str, sheet_name: str) -> str:
    """Remove freeze panes from a sheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.freeze_panes = None
        save_workbook_safe(wb, file_path)
        logger.info("Unfroze panes in '%s' of %s", sheet_name, file_path)
        return f"Panes unfrozen in sheet '{sheet_name}'."
    finally:
        wb.close()


def set_auto_filter(file_path: str, sheet_name: str, cell_range: str) -> str:
    """Enable auto-filter on a range (e.g. 'A1:E1' or 'A1:E100')."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.auto_filter.ref = cell_range
        save_workbook_safe(wb, file_path)
        logger.info("Set auto-filter %s in '%s' of %s", cell_range, sheet_name, file_path)
        return f"Auto-filter set on '{cell_range}' in sheet '{sheet_name}'."
    finally:
        wb.close()


def remove_auto_filter(file_path: str, sheet_name: str) -> str:
    """Remove auto-filter from a sheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.auto_filter.ref = None
        save_workbook_safe(wb, file_path)
        logger.info("Removed auto-filter in '%s' of %s", sheet_name, file_path)
        return f"Auto-filter removed from sheet '{sheet_name}'."
    finally:
        wb.close()


def hide_rows(file_path: str, sheet_name: str, start_row: int, end_row: int) -> str:
    """Hide rows from start_row to end_row (inclusive, 1-based)."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        for r in range(start_row, end_row + 1):
            ws.row_dimensions[r].hidden = True
        save_workbook_safe(wb, file_path)
        logger.info("Hid rows %d-%d in '%s' of %s", start_row, end_row, sheet_name, file_path)
        return f"Rows {start_row}-{end_row} hidden in sheet '{sheet_name}'."
    finally:
        wb.close()


def unhide_rows(file_path: str, sheet_name: str, start_row: int, end_row: int) -> str:
    """Unhide rows from start_row to end_row (inclusive, 1-based)."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        for r in range(start_row, end_row + 1):
            ws.row_dimensions[r].hidden = False
        save_workbook_safe(wb, file_path)
        logger.info("Unhid rows %d-%d in '%s' of %s", start_row, end_row, sheet_name, file_path)
        return f"Rows {start_row}-{end_row} unhidden in sheet '{sheet_name}'."
    finally:
        wb.close()


def hide_columns(file_path: str, sheet_name: str, start_col: str, end_col: str) -> str:
    """Hide columns from start_col to end_col (letters like 'A', 'D')."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        start_idx = col_letter_to_index(start_col)
        end_idx = col_letter_to_index(end_col)
        for i in range(start_idx, end_idx + 1):
            col = index_to_col_letter(i)
            ws.column_dimensions[col].hidden = True
        save_workbook_safe(wb, file_path)
        logger.info("Hid columns %s-%s in '%s' of %s", start_col, end_col, sheet_name, file_path)
        return f"Columns {start_col}-{end_col} hidden in sheet '{sheet_name}'."
    finally:
        wb.close()


def unhide_columns(file_path: str, sheet_name: str, start_col: str, end_col: str) -> str:
    """Unhide columns from start_col to end_col (letters like 'A', 'D')."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        start_idx = col_letter_to_index(start_col)
        end_idx = col_letter_to_index(end_col)
        for i in range(start_idx, end_idx + 1):
            col = index_to_col_letter(i)
            ws.column_dimensions[col].hidden = False
        save_workbook_safe(wb, file_path)
        logger.info("Unhid columns %s-%s in '%s' of %s", start_col, end_col, sheet_name, file_path)
        return f"Columns {start_col}-{end_col} unhidden in sheet '{sheet_name}'."
    finally:
        wb.close()


def group_rows(
    file_path: str,
    sheet_name: str,
    start_row: int,
    end_row: int,
    outline_level: int = 1,
    hidden: bool = False,
) -> str:
    """Group rows with an outline level."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        for r in range(start_row, end_row + 1):
            ws.row_dimensions[r].outlineLevel = outline_level
        if hidden:
            for r in range(start_row, end_row + 1):
                ws.row_dimensions[r].hidden = True
        save_workbook_safe(wb, file_path)
        logger.info(
            "Grouped rows %d-%d (level %d) in '%s' of %s",
            start_row,
            end_row,
            outline_level,
            sheet_name,
            file_path,
        )
        return f"Rows {start_row}-{end_row} grouped at outline level {outline_level} in sheet '{sheet_name}'."
    finally:
        wb.close()


def ungroup_rows(file_path: str, sheet_name: str, start_row: int, end_row: int) -> str:
    """Ungroup rows by resetting outline level to 0."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        for r in range(start_row, end_row + 1):
            ws.row_dimensions[r].outlineLevel = 0
        save_workbook_safe(wb, file_path)
        logger.info("Ungrouped rows %d-%d in '%s' of %s", start_row, end_row, sheet_name, file_path)
        return f"Rows {start_row}-{end_row} ungrouped in sheet '{sheet_name}'."
    finally:
        wb.close()


def group_columns(
    file_path: str,
    sheet_name: str,
    start_col: str,
    end_col: str,
    outline_level: int = 1,
    hidden: bool = False,
) -> str:
    """Group columns with an outline level."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        start_idx = col_letter_to_index(start_col)
        end_idx = col_letter_to_index(end_col)
        for i in range(start_idx, end_idx + 1):
            col = index_to_col_letter(i)
            ws.column_dimensions[col].outlineLevel = outline_level
        if hidden:
            for i in range(start_idx, end_idx + 1):
                col = index_to_col_letter(i)
                ws.column_dimensions[col].hidden = True
        save_workbook_safe(wb, file_path)
        logger.info(
            "Grouped columns %s-%s (level %d) in '%s' of %s",
            start_col,
            end_col,
            outline_level,
            sheet_name,
            file_path,
        )
        return f"Columns {start_col}-{end_col} grouped at outline level {outline_level} in sheet '{sheet_name}'."
    finally:
        wb.close()


def ungroup_columns(file_path: str, sheet_name: str, start_col: str, end_col: str) -> str:
    """Ungroup columns by resetting outline level to 0."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        start_idx = col_letter_to_index(start_col)
        end_idx = col_letter_to_index(end_col)
        for i in range(start_idx, end_idx + 1):
            col = index_to_col_letter(i)
            ws.column_dimensions[col].outlineLevel = 0
        save_workbook_safe(wb, file_path)
        logger.info("Ungrouped columns %s-%s in '%s' of %s", start_col, end_col, sheet_name, file_path)
        return f"Columns {start_col}-{end_col} ungrouped in sheet '{sheet_name}'."
    finally:
        wb.close()


def set_sheet_tab_color(file_path: str, sheet_name: str, color: str) -> str:
    """Set the sheet tab color (hex like 'FF0000')."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.sheet_properties.tabColor = color
        save_workbook_safe(wb, file_path)
        logger.info("Set tab color %s on '%s' in %s", color, sheet_name, file_path)
        return f"Tab color set to '{color}' for sheet '{sheet_name}'."
    finally:
        wb.close()


def hide_sheet(file_path: str, sheet_name: str, very_hidden: bool = False) -> str:
    """Hide a sheet. Use very_hidden=True for 'veryHidden' state."""
    wb = load_workbook_safe(file_path)
    try:
        visible_sheets = [s for s in wb.sheetnames if wb[s].sheet_state == "visible"]
        if sheet_name in visible_sheets and len(visible_sheets) <= 1:
            raise ValueError("Cannot hide the only visible sheet in the workbook.")
        ws = get_sheet(wb, sheet_name)
        ws.sheet_state = "veryHidden" if very_hidden else "hidden"
        save_workbook_safe(wb, file_path)
        state = "very hidden" if very_hidden else "hidden"
        logger.info("Set sheet '%s' to %s in %s", sheet_name, state, file_path)
        return f"Sheet '{sheet_name}' is now {state}."
    finally:
        wb.close()


def unhide_sheet(file_path: str, sheet_name: str) -> str:
    """Unhide a sheet by setting its state to visible."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.sheet_state = "visible"
        save_workbook_safe(wb, file_path)
        logger.info("Unhid sheet '%s' in %s", sheet_name, file_path)
        return f"Sheet '{sheet_name}' is now visible."
    finally:
        wb.close()


def move_sheet(file_path: str, sheet_name: str, offset: int) -> str:
    """Move a sheet by offset positions (positive=right, negative=left)."""
    wb = load_workbook_safe(file_path)
    try:
        get_sheet(wb, sheet_name)  # validate existence
        wb.move_sheet(sheet_name, offset)
        save_workbook_safe(wb, file_path)
        logger.info("Moved sheet '%s' by %d in %s", sheet_name, offset, file_path)
        return f"Sheet '{sheet_name}' moved by {offset} position(s)."
    finally:
        wb.close()


def set_zoom(file_path: str, sheet_name: str, zoom_scale: int) -> str:
    """Set sheet zoom level (10-400)."""
    if not 10 <= zoom_scale <= 400:
        raise ValueError(f"Zoom scale must be between 10 and 400, got {zoom_scale}.")
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.sheet_view.zoomScale = zoom_scale
        save_workbook_safe(wb, file_path)
        logger.info("Set zoom to %d%% on '%s' in %s", zoom_scale, sheet_name, file_path)
        return f"Zoom set to {zoom_scale}% for sheet '{sheet_name}'."
    finally:
        wb.close()


def show_gridlines(file_path: str, sheet_name: str, show: bool = True) -> str:
    """Show or hide gridlines on a sheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.sheet_view.showGridLines = show
        save_workbook_safe(wb, file_path)
        state = "shown" if show else "hidden"
        logger.info("Gridlines %s on '%s' in %s", state, sheet_name, file_path)
        return f"Gridlines {state} for sheet '{sheet_name}'."
    finally:
        wb.close()


def set_print_area(file_path: str, sheet_name: str, cell_range: str) -> str:
    """Set the print area for a sheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.print_area = cell_range
        save_workbook_safe(wb, file_path)
        logger.info("Set print area %s on '%s' in %s", cell_range, sheet_name, file_path)
        return f"Print area set to '{cell_range}' for sheet '{sheet_name}'."
    finally:
        wb.close()


def set_page_setup(
    file_path: str,
    sheet_name: str,
    orientation: str = "portrait",
    paper_size: int = 1,
    fit_to_width: int | None = None,
    fit_to_height: int | None = None,
) -> str:
    """Configure page setup: orientation, paper size, and fit-to-page scaling."""
    if orientation not in ("portrait", "landscape"):
        raise ValueError(f"Orientation must be 'portrait' or 'landscape', got '{orientation}'.")
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.page_setup.orientation = orientation
        ws.page_setup.paperSize = paper_size
        if fit_to_width is not None or fit_to_height is not None:
            ws.page_setup.fitToPage = True
            if fit_to_width is not None:
                ws.page_setup.fitToWidth = fit_to_width
            if fit_to_height is not None:
                ws.page_setup.fitToHeight = fit_to_height
        save_workbook_safe(wb, file_path)
        logger.info("Set page setup on '%s' in %s", sheet_name, file_path)
        return f"Page setup configured for sheet '{sheet_name}': orientation={orientation}, paper_size={paper_size}."
    finally:
        wb.close()


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
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.page_margins = PageMargins(
            top=top,
            bottom=bottom,
            left=left,
            right=right,
            header=header,
            footer=footer,
        )
        save_workbook_safe(wb, file_path)
        logger.info("Set page margins on '%s' in %s", sheet_name, file_path)
        return f"Page margins set for sheet '{sheet_name}'."
    finally:
        wb.close()


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
    """Set headers and footers for a sheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        if header_center is not None:
            ws.oddHeader.center.text = header_center
        if header_left is not None:
            ws.oddHeader.left.text = header_left
        if header_right is not None:
            ws.oddHeader.right.text = header_right
        if footer_center is not None:
            ws.oddFooter.center.text = footer_center
        if footer_left is not None:
            ws.oddFooter.left.text = footer_left
        if footer_right is not None:
            ws.oddFooter.right.text = footer_right
        save_workbook_safe(wb, file_path)
        logger.info("Set header/footer on '%s' in %s", sheet_name, file_path)
        return f"Header/footer configured for sheet '{sheet_name}'."
    finally:
        wb.close()
