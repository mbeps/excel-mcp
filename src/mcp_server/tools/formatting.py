"""Cell and sheet formatting operations."""

from __future__ import annotations

from logging import Logger
from typing import Any, cast

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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
                left=Side(style=cast(Any, left_border_style), color=border_color) if left_border_style else Side(),
                right=Side(style=cast(Any, right_border_style), color=border_color) if right_border_style else Side(),
                top=Side(style=cast(Any, top_border_style), color=border_color) if top_border_style else Side(),
                bottom=(
                    Side(style=cast(Any, bottom_border_style), color=border_color) if bottom_border_style else Side()
                ),
            )
        elif border_style:
            side = Side(style=cast(Any, border_style), color=border_color)
            border = Border(left=side, right=side, top=side, bottom=side)

        for row in ws[cell_range]:
            cells = row if isinstance(row, tuple) else (row,)
            for cell in cells:
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


def set_column_width(file_path: str, sheet_name: str, column: str, width: float) -> str:
    """Set the width of a column (column letter like 'A')."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.column_dimensions[column.upper()].width = width
        save_workbook_safe(wb, file_path)
        logger.info("Set column %s width to %s in %s", column, width, sheet_name)
        return f"Column {column.upper()} width set to {width} in '{sheet_name}'."
    finally:
        wb.close()


def set_row_height(file_path: str, sheet_name: str, row: int, height: float) -> str:
    """Set the height of a row (1-based)."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.row_dimensions[row].height = height
        save_workbook_safe(wb, file_path)
        logger.info("Set row %d height to %s in %s", row, height, sheet_name)
        return f"Row {row} height set to {height} in '{sheet_name}'."
    finally:
        wb.close()


def merge_cells(file_path: str, sheet_name: str, cell_range: str) -> str:
    """Merge a range of cells (e.g. 'A1:C1')."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.merge_cells(cell_range)
        save_workbook_safe(wb, file_path)
        logger.info("Merged %s in %s", cell_range, sheet_name)
        return f"Merged cells {cell_range} in '{sheet_name}'."
    finally:
        wb.close()


def unmerge_cells(file_path: str, sheet_name: str, cell_range: str) -> str:
    """Unmerge a previously merged range."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.unmerge_cells(cell_range)
        save_workbook_safe(wb, file_path)
        logger.info("Unmerged %s in %s", cell_range, sheet_name)
        return f"Unmerged cells {cell_range} in '{sheet_name}'."
    finally:
        wb.close()


def get_cell_formatting(file_path: str, sheet_name: str, cell_ref: str) -> dict:
    """Read all formatting properties of a cell."""
    wb = load_workbook_safe(file_path, read_only=False)
    try:
        ws = get_sheet(wb, sheet_name)
        cell = ws[cell_ref]

        f = cell.font
        font_info = {
            "name": f.name,
            "size": f.size,
            "bold": f.bold,
            "italic": f.italic,
            "underline": f.underline,
            "strike": f.strike,
            "color": str(f.color.rgb) if f.color and f.color.rgb else None,
            "vertAlign": f.vertAlign,
        }

        fl = cell.fill
        fill_info = {
            "fill_type": fl.fill_type,
            "fgColor": str(fl.fgColor.rgb) if fl.fgColor and fl.fgColor.rgb and fl.fgColor.rgb != "00000000" else None,
            "bgColor": str(fl.bgColor.rgb) if fl.bgColor and fl.bgColor.rgb and fl.bgColor.rgb != "00000000" else None,
        }

        b = cell.border

        def _side_dict(s: Any) -> dict:
            return {
                "style": s.style,
                "color": str(s.color.rgb) if s.color and s.color.rgb else None,
            }

        border_info = {
            "left": _side_dict(b.left),
            "right": _side_dict(b.right),
            "top": _side_dict(b.top),
            "bottom": _side_dict(b.bottom),
        }

        a = cell.alignment
        alignment_info = {
            "horizontal": a.horizontal,
            "vertical": a.vertical,
            "wrap_text": a.wrap_text,
            "text_rotation": a.textRotation,
            "indent": a.indent,
            "shrink_to_fit": a.shrinkToFit,
        }

        p = cell.protection
        protection_info = {
            "locked": p.locked,
            "hidden": p.hidden,
        }

        return {
            "font": font_info,
            "fill": fill_info,
            "border": border_info,
            "alignment": alignment_info,
            "number_format": cell.number_format,
            "protection": protection_info,
        }
    finally:
        wb.close()


def list_merged_ranges(file_path: str, sheet_name: str) -> list[str]:
    """List all merged cell ranges in a sheet."""
    wb = load_workbook_safe(file_path, read_only=False)
    try:
        ws = get_sheet(wb, sheet_name)
        return [str(r) for r in ws.merged_cells.ranges]
    finally:
        wb.close()


def format_range_per_cell(
    file_path: str,
    sheet_name: str,
    start_cell: str,
    styles: list[list[dict | None]],
) -> str:
    """Apply per-cell formatting using a 2D array of style dicts matching range dimensions."""
    from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        col_letter, start_row = coordinate_from_string(start_cell)
        start_col = column_index_from_string(col_letter)

        cells_formatted = 0
        for row_offset, row_styles in enumerate(styles):
            for col_offset, style in enumerate(row_styles):
                if style is None:
                    continue

                cell = ws.cell(row=start_row + row_offset, column=start_col + col_offset)

                if any(k in style for k in ("font_name", "font_size", "bold", "italic", "font_color")):
                    cell.font = Font(
                        name=style.get("font_name"),
                        size=style.get("font_size"),
                        bold=style.get("bold", False),
                        italic=style.get("italic", False),
                        color=style.get("font_color"),
                    )

                if "fill_color" in style:
                    cell.fill = PatternFill(
                        start_color=style["fill_color"],
                        end_color=style["fill_color"],
                        fill_type="solid",
                    )

                if "number_format" in style:
                    cell.number_format = style["number_format"]

                if any(k in style for k in ("horizontal_alignment", "vertical_alignment")):
                    cell.alignment = Alignment(
                        horizontal=style.get("horizontal_alignment"),
                        vertical=style.get("vertical_alignment"),
                    )

                if "border_style" in style:
                    side = Side(style=cast(Any, style["border_style"]))
                    cell.border = Border(left=side, right=side, top=side, bottom=side)

                cells_formatted += 1

        save_workbook_safe(wb, file_path)
        logger.info("Formatted %d cells in %s!%s", cells_formatted, file_path, sheet_name)
        return f"Formatted {cells_formatted} cells starting at {start_cell} in '{sheet_name}'."
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
