from __future__ import annotations

from datetime import date, datetime
from typing import Literal

from mcp.types import ToolAnnotations

import mcp_server.tools.cell_ops as _cell_ops
from mcp_server.models.pivot_etl import ChunkReadResult

__all__ = [
    "read_cells",
    "write_cells",
    "clear_range",
    "copy_range",
    "find_replace",
    "transpose_range",
]


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


def clear_range(file_path: str, sheet_name: str, start_cell: str, end_cell: str) -> str:
    """Clear all values in a rectangular cell range."""
    return _cell_ops.clear_range(file_path, sheet_name, start_cell, end_cell)


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


def find_replace(
    file_path: str,
    sheet_name: str,
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    match_entire_cell: bool = False,
    search_formulas: bool = False,
) -> dict:
    """Find and replace text across a worksheet. Set match_case for case-sensitive search."""
    return _cell_ops.find_replace(
        file_path,
        sheet_name,
        find_text,
        replace_text,
        match_case,
        match_entire_cell,
        search_formulas,
    )


def transpose_range(
    file_path: str,
    sheet_name: str,
    source_range: str,
    target_cell: str,
    source_sheet: str | None = None,
    paste_values_only: bool = False,
) -> dict:
    """Transpose (swap rows and columns) of source_range and write starting at target_cell."""
    return _cell_ops.transpose_range(
        file_path,
        sheet_name,
        source_range,
        target_cell,
        source_sheet,
        paste_values_only,
    )


def register(mcp) -> None:
    """Register tools on *mcp*."""
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(read_cells)
    mcp.tool()(write_cells)
    mcp.tool(annotations=ToolAnnotations(destructiveHint=True))(clear_range)
    mcp.tool()(copy_range)
    mcp.tool()(find_replace)
    mcp.tool()(transpose_range)
