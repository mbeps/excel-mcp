from __future__ import annotations

from datetime import date, datetime
from typing import Any, Literal

from mcp.types import ToolAnnotations

import mcp_server.tools.cell_ops as _cell_ops
from mcp_server.models.cell_ops import ChunkReadResult

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

    Args:
        mode: One of "single", "range", or "chunked". Determines the behaviour and required parameters.
            - "single": returns a single cell value. Requires `cell_ref`.
            - "range": returns a rectangular range. Requires `start_cell` and `end_cell`.
            - "chunked": returns a chunked reader result for large sheets. Optional: `start_row`, `chunk_size`.
        file_path: Path to the workbook (validated by utils).
        sheet_name: Worksheet name.
        cell_ref: Cell reference for single-cell reads (e.g. "A1").
        start_cell: Top-left cell for range reads.
        end_cell: Bottom-right cell for range reads.
        include_formula: If True, include formula text when available.
        include_metadata: If True, include additional metadata (styles, comment presence, etc.).
        show_formula: For range reads, include formulas instead of values when True.
        show_style: For range reads, include style information when True.
        output_format: Format for range output; typically "json".
        max_cells: Maximum cells to return for range reads (to avoid huge payloads).
        start_row: For chunked reads, starting row index (0-based).
        chunk_size: Number of rows per chunk for chunked reads.

    Returns:
        dict or ChunkReadResult: Single value, range payload, or chunked reader object depending on `mode`.

    Raises:
        ValueError: If required parameters for the chosen `mode` are missing or `mode` is unknown.

    Notes:
        - Read-only: this function only reads workbook data and should not mutate files.
        - Dispatch mapping: "single"→`tools.cell_ops.read_cell`, "range"→`tools.cell_ops.read_range`, "chunked"→`tools.cell_ops.read_file_chunked`.
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
    """Write data or operations to worksheet cells.

    Args:
        mode: One of "single", "range", "series", "merge", or "unmerge".
            - "single": write a single cell. Requires `cell_ref` and `value`.
            - "range": write a 2D array starting at `start_cell`. Requires `start_cell` and `data`.
            - "series": fill a series from `start_cell`. Requires `start_cell` and `count`.
            - "merge": merge a cell range. Requires `range_string`.
            - "unmerge": unmerge a cell range. Requires `range_string`.
        file_path: Path to workbook to mutate.
        sheet_name: Target worksheet name.
        cell_ref: Single-cell reference for "single" and some formula operations.
        value: Value to write for single-cell writes.
        start_cell: Top-left cell for range/series writes.
        data: 2D array of values for "range" mode.
        series_type: Series type for "series" (e.g. "number").
        count: Number of series entries to write (required for "series").
        step: Increment for numeric series or string step expression.
        direction: "down" or "right" for series growth.
        start_value: Optional starting value for series.
        range_string: Range string for merge/unmerge (e.g. "A1:D1").

    Returns:
        str or dict: Success message or structured result depending on operation.

    Raises:
        ValueError: If required parameters for the chosen `mode` are missing or `mode` is unknown.

    Notes:
        - This route mutates workbook files (destructive operations).
        - Dispatch mapping: "single"→`tools.cell_ops.write_cell`, "range"→`tools.cell_ops.write_range`,
          "series"→`tools.cell_ops.fill_series`, "merge"→`tools.cell_ops.merge_cells`, "unmerge"→`tools.cell_ops.unmerge_cells`.
        - Recommend adding short examples in the underlying tools for common series patterns.
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
    """Clear values from a rectangular range of cells.

    Args:
        file_path: Path to workbook.
        sheet_name: Worksheet name.
        start_cell: Top-left cell of range to clear.
        end_cell: Bottom-right cell of range to clear.

    Returns:
        str: Success message.

    Raises:
        (Propagates exceptions from the underlying tools on I/O or invalid ranges.)

    Notes:
        - Destructive: clears cell values (but not necessarily styles) — mention in API docs.
    """
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
    """Copy a range of cells to a destination range (same workbook or across sheets).

    Args:
        file_path: Source workbook path (when copying within same workbook).
        source_sheet: Source worksheet name.
        source_range: Range to copy (e.g. "A1:B10").
        dest_sheet: Destination worksheet name.
        dest_range: Destination top-left target (e.g. "C1").
        copy_values: If True, copy values.
        copy_styles: If True, copy styles.
        paste_values_only: If True, read resolved values and write values (not formulas).

    Returns:
        str: Success message.

    Notes:
        - May open the workbook twice if `paste_values_only` is True (read-only pass then write pass).
    """
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
    regex: bool = False,
) -> dict:
    """Find and replace text across a worksheet.

    Args:
        file_path: Path to workbook to modify.
        sheet_name: Worksheet name.
        find_text: Substring or pattern to find.
        replace_text: Replacement text.
        match_case: Case-sensitive search when True.
        match_entire_cell: Match entire cell contents exactly when True.
        search_formulas: Also search within formulas when True.
        regex: Treat find_text as a regular expression when True.

    Returns:
        dict: {"count": int, "cells": ["A1", ...]} detailing replacements.

    Notes:
        - Destructive: modifies cells in-place.
        - Consider returning per-cell before/after pairs for audit logging in high-risk contexts.
    """
    return _cell_ops.find_replace(
        file_path,
        sheet_name,
        find_text,
        replace_text,
        match_case,
        match_entire_cell,
        search_formulas,
        regex,
    )


def transpose_range(
    file_path: str,
    sheet_name: str,
    source_range: str,
    target_cell: str,
    source_sheet: str | None = None,
    paste_values_only: bool = False,
) -> dict:
    """Transpose a source range (rows↔columns) and write starting at `target_cell`.

    Args:
        file_path: Path to workbook.
        sheet_name: Target worksheet name (destination for transposed data).
        source_range: Range to transpose (on `sheet_name` by default).
        target_cell: Top-left cell for the transposed output.
        source_sheet: Optional source sheet name if different from `sheet_name`.
        paste_values_only: If True, paste only resolved values (no formulas).

    Returns:
        dict: Summary including target range written and number of cells.

    Notes:
        - Destructive: overwrites destination cells.
    """
    return _cell_ops.transpose_range(
        file_path,
        sheet_name,
        source_range,
        target_cell,
        source_sheet,
        paste_values_only,
    )


def register(mcp: Any) -> None:
    """Register tools on *mcp*."""
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(read_cells)
    mcp.tool()(write_cells)
    mcp.tool(annotations=ToolAnnotations(destructiveHint=True))(clear_range)
    mcp.tool()(copy_range)
    mcp.tool()(find_replace)
    mcp.tool()(transpose_range)
