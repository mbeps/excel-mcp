from __future__ import annotations

import mcp_server.tools.formatting as _formatting
from mcp_server.models.common import (
    BorderStyle,
    HorizontalAlignment,
    VerticalAlignment,
)

__all__ = [
    "format_cells",
    "auto_fit_columns",
    "copy_cell_format",
    "clear_cell_format",
    "apply_named_style",
]


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


def auto_fit_columns(file_path: str, sheet_name: str) -> str:
    """Auto-fit all column widths based on content length."""
    return _formatting.auto_fit_columns(file_path, sheet_name)


def copy_cell_format(
    file_path: str,
    sheet_name: str,
    source_cell: str,
    target_range: str,
) -> dict:
    """Copy all formatting from source_cell and apply it to every cell in target_range."""
    return _formatting.copy_cell_format(file_path, sheet_name, source_cell, target_range)


def clear_cell_format(
    file_path: str,
    sheet_name: str,
    range_str: str,
) -> dict:
    """Reset all formatting on cells in range_str without touching values."""
    return _formatting.clear_cell_format(file_path, sheet_name, range_str)


def apply_named_style(
    file_path: str,
    sheet_name: str,
    range_str: str,
    style_name: str,
) -> dict:
    """Apply a named built-in Excel style (e.g. 'Good', 'Bad', 'Neutral', 'Heading 1') to a cell range."""
    return _formatting.apply_named_style(file_path, sheet_name, range_str, style_name)


def register(mcp) -> None:
    """Register tools on *mcp*."""
    mcp.tool()(format_cells)
    mcp.tool()(auto_fit_columns)
    mcp.tool()(copy_cell_format)
    mcp.tool()(clear_cell_format)
    mcp.tool()(apply_named_style)
