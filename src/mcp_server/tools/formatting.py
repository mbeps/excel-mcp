"""Cell and sheet formatting operations."""

from __future__ import annotations

from logging import Logger

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from mcp_server.models.formatting import BorderStyle
from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


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
    preserve_existing: bool = False,
) -> str:
    """Apply formatting to a cell range (e.g. 'A1:C10' or 'A1')."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        font = Font(
            name=font_name,
            bold=bold,
            italic=italic,
            size=font_size,
            color=font_color,
            underline=underline,
            strike=strikethrough,
        )
        fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid") if bg_color else None
        alignment = Alignment(
            horizontal=horizontal_alignment,
            vertical=vertical_alignment,
            wrap_text=wrap_text,
            textRotation=text_rotation if text_rotation is not None else 0,
            indent=indent if indent is not None else 0,
            shrinkToFit=shrink_to_fit,
        )
        border = None
        has_per_side = any([top_border_style, bottom_border_style, left_border_style, right_border_style])
        if has_per_side:
            border = Border(
                left=(Side(style=left_border_style, color=border_color) if left_border_style else Side()),
                right=(Side(style=right_border_style, color=border_color) if right_border_style else Side()),
                top=(Side(style=top_border_style, color=border_color) if top_border_style else Side()),
                bottom=(Side(style=bottom_border_style, color=border_color) if bottom_border_style else Side()),
            )
        elif border_style:
            side = Side(style=border_style, color=border_color)
            border = Border(left=side, right=side, top=side, bottom=side)

        range_data = ws[cell_range]
        if not isinstance(range_data, tuple):
            range_data = ((range_data,),)
        for row in range_data:
            cells = row if isinstance(row, tuple) else (row,)
            for cell in cells:
                if preserve_existing:
                    ef = cell.font
                    cell.font = Font(
                        name=font_name if font_name is not None else ef.name,
                        bold=bold if bold else ef.bold,
                        italic=italic if italic else ef.italic,
                        size=font_size if font_size is not None else ef.size,
                        color=font_color if font_color is not None else ef.color,
                        underline=underline if underline is not None else ef.underline,
                        strike=strikethrough if strikethrough else ef.strike,
                    )
                    if bg_color:
                        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
                    ea = cell.alignment
                    cell.alignment = Alignment(
                        horizontal=horizontal_alignment if horizontal_alignment is not None else ea.horizontal,
                        vertical=vertical_alignment if vertical_alignment is not None else ea.vertical,
                        wrap_text=wrap_text if wrap_text else ea.wrap_text,
                        textRotation=text_rotation if text_rotation is not None else (ea.textRotation or 0),
                        indent=indent if indent is not None else (ea.indent or 0),
                        shrinkToFit=shrink_to_fit if shrink_to_fit else ea.shrinkToFit,
                    )
                    if border:
                        cell.border = border
                    if number_format:
                        cell.number_format = number_format
                else:
                    cell.font = font
                    if fill:
                        cell.fill = fill
                    cell.alignment = alignment
                    if border:
                        cell.border = border
                    if number_format:
                        cell.number_format = number_format

        save_workbook_safe(wb, file_path)
        logger.info("Formatted range %s in %s!%s", cell_range, file_path, sheet_name)
        return f"Formatting applied to {cell_range} in '{sheet_name}'."
    finally:
        wb.close()


def auto_fit_columns(file_path: str, sheet_name: str) -> str:
    """Auto-fit all column widths based on content length."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        for col_cells in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col_cells[0].column or 1)
            for cell in col_cells:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 2

        save_workbook_safe(wb, file_path)
        logger.info("Auto-fitted columns in %s!%s", file_path, sheet_name)
        return f"Auto-fitted all column widths in '{sheet_name}'."
    finally:
        wb.close()
