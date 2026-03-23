"""Cell-level read/write operations and chunked reading."""

from __future__ import annotations

import copy
import math
from logging import Logger

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
    validate_file_path,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def read_cell(file_path: str, sheet_name: str, cell_ref: str) -> dict:
    """Read a single cell and return its value and data type."""
    wb = load_workbook_safe(file_path, read_only=True)
    try:
        ws = get_sheet(wb, sheet_name)
        cell = ws[cell_ref]
        return {
            "cell_ref": cell_ref,
            "value": cell.value,
            "data_type": cell.data_type,
        }
    finally:
        wb.close()


def write_cell(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
    value: str | int | float | bool | None,
) -> str:
    """Write a value to a single cell."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws[cell_ref] = value
        save_workbook_safe(wb, file_path)
        logger.info("Wrote to %s!%s", sheet_name, cell_ref)
        return f"Cell {cell_ref} in '{sheet_name}' set to {value!r}."
    finally:
        wb.close()


def read_range(
    file_path: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    show_formula: bool = False,
    show_style: bool = False,
    output_format: str = "json",
    max_cells: int | None = None,
) -> dict:
    """Read a rectangular range and return rows as a list of lists.

    Args:
        show_formula: When True, return stored formulas instead of values.
        show_style: When True, include style metadata per cell.
        output_format: 'json' (default) or 'html' for HTML table output.
        max_cells: If set, cap total cells returned; excess rows are truncated.
    """
    from openpyxl.utils import coordinate_to_tuple, get_column_letter

    wb = load_workbook_safe(file_path, read_only=not show_style)
    try:
        ws = get_sheet(wb, sheet_name)

        start_r, start_c = coordinate_to_tuple(start_cell)
        end_r, end_c = coordinate_to_tuple(end_cell)
        total_rows_available = end_r - start_r + 1
        col_count_available = end_c - start_c + 1
        total_cells_available = total_rows_available * col_count_available

        read_end_cell = end_cell
        truncated = False
        if max_cells is not None and total_cells_available > max_cells:
            max_rows = max(1, max_cells // col_count_available)
            new_end_r = start_r + max_rows - 1
            read_end_cell = f"{get_column_letter(end_c)}{new_end_r}"
            truncated = True

        rows: list[list] = []
        styles: list[list[dict]] = []
        for row in ws[f"{start_cell}:{read_end_cell}"]:
            row_values = []
            row_styles = []
            for cell in row:
                if show_formula and isinstance(cell.value, str) and cell.value.startswith("="):
                    row_values.append(cell.value)
                else:
                    row_values.append(cell.value)
                if show_style:
                    font = cell.font
                    fill = cell.fill
                    alignment = cell.alignment
                    fg_color = None
                    if fill and fill.fgColor and fill.fgColor.rgb and fill.fgColor.rgb != "00000000":
                        fg_color = str(fill.fgColor.rgb)
                    font_color = None
                    if font and font.color and font.color.rgb:
                        font_color = str(font.color.rgb)
                    row_styles.append(
                        {
                            "font_name": font.name if font else None,
                            "font_size": font.size if font else None,
                            "bold": font.bold if font else False,
                            "italic": font.italic if font else False,
                            "font_color": font_color,
                            "fill_color": fg_color,
                            "number_format": cell.number_format,
                            "alignment": {
                                "horizontal": alignment.horizontal if alignment else None,
                                "vertical": alignment.vertical if alignment else None,
                                "wrap_text": alignment.wrap_text if alignment else None,
                            },
                        }
                    )
            rows.append(row_values)
            if show_style:
                styles.append(row_styles)

        if output_format == "html":
            html_parts = ["<table>"]
            for row_data in rows:
                html_parts.append("<tr>")
                for val in row_data:
                    escaped = str(val) if val is not None else ""
                    escaped = escaped.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                    html_parts.append(f"<td>{escaped}</td>")
                html_parts.append("</tr>")
            html_parts.append("</table>")
            result: dict = {
                "html": "\n".join(html_parts),
                "row_count": len(rows),
                "col_count": len(rows[0]) if rows else 0,
            }
        else:
            result = {
                "rows": rows,
                "row_count": len(rows),
                "col_count": len(rows[0]) if rows else 0,
            }

        if show_style:
            result["styles"] = styles
        result["truncated"] = truncated
        if truncated:
            result["total_cells_available"] = total_cells_available
        return result
    finally:
        wb.close()


def write_range(file_path: str, sheet_name: str, start_cell: str, data: list[list]) -> str:
    """Write a 2D array starting from start_cell."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        start = ws[start_cell]
        start_row, start_col = start.row, start.column

        for r_offset, row_data in enumerate(data):
            for c_offset, value in enumerate(row_data):
                ws.cell(row=start_row + r_offset, column=start_col + c_offset, value=value)

        save_workbook_safe(wb, file_path)
        rows_written = len(data)
        cols_written = len(data[0]) if data else 0
        logger.info("Wrote %dx%d range at %s!%s", rows_written, cols_written, sheet_name, start_cell)
        return f"Wrote {rows_written} rows x {cols_written} cols starting at {start_cell} in '{sheet_name}'."
    finally:
        wb.close()


def clear_range(file_path: str, sheet_name: str, start_cell: str, end_cell: str) -> str:
    """Clear all values in a rectangular range."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        for row in ws[f"{start_cell}:{end_cell}"]:
            for cell in row:
                cell.value = None

        save_workbook_safe(wb, file_path)
        logger.info("Cleared range %s:%s in %s", start_cell, end_cell, sheet_name)
        return f"Cleared range {start_cell}:{end_cell} in '{sheet_name}'."
    finally:
        wb.close()


def read_cell_detailed(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
    data_only: bool = False,
) -> dict:
    """Read a cell with full detail including type, formula, comments, and merge status."""
    wb = load_workbook_safe(file_path, data_only=data_only)
    try:
        ws = get_sheet(wb, sheet_name)
        cell = ws[cell_ref]
        value = cell.value
        raw_type = cell.data_type

        type_map = {
            "s": "string",
            "n": "number",
            "b": "boolean",
            "d": "datetime",
            "f": "formula",
            "e": "error",
        }
        data_type = type_map.get(raw_type, "empty") if value is not None else "empty"

        formula = None
        if data_only:
            wb_formula = load_workbook_safe(file_path, data_only=False)
            try:
                ws_f = get_sheet(wb_formula, sheet_name)
                f_cell = ws_f[cell_ref]
                if f_cell.data_type == "f":
                    formula = str(f_cell.value)
            finally:
                wb_formula.close()
        elif raw_type == "f":
            formula = str(value)

        is_merged = False
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                is_merged = True
                break

        return {
            "cell_ref": cell_ref,
            "value": value,
            "data_type": data_type,
            "formula": formula,
            "has_comment": cell.comment is not None,
            "has_hyperlink": cell.hyperlink is not None,
            "number_format": cell.number_format,
            "is_merged": is_merged,
        }
    finally:
        wb.close()


def read_file_chunked(
    file_path: str,
    sheet_name: str,
    start_row: int = 0,
    chunk_size: int = 1000,
) -> dict:
    """Read a sheet in chunks using pandas for performance."""
    import pandas as pd

    path = validate_file_path(file_path)

    try:
        df = pd.read_excel(path, sheet_name=sheet_name, engine="calamine")
    except ImportError:
        logger.warning("calamine engine not available, falling back to openpyxl")
        df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")

    total_rows = len(df)
    end_row = start_row + chunk_size
    chunk = df.iloc[start_row:end_row]
    has_more = end_row < total_rows
    total_pages = math.ceil(total_rows / chunk_size) if chunk_size > 0 else 1
    current_page = (start_row // chunk_size) + 1 if chunk_size > 0 else 1

    actual_chunk_size = len(chunk)
    return {
        "rows": chunk.to_dict(orient="records"),
        "chunk_start": start_row,
        "chunk_size": actual_chunk_size,
        "has_more": has_more,
        "next_start_row": start_row + actual_chunk_size if has_more else None,
        "next_offset": end_row if has_more else None,
        "total_rows": total_rows,
        "current_page": current_page,
        "total_pages": total_pages,
    }


def copy_range(
    file_path: str,
    source_sheet: str,
    source_range: str,
    dest_sheet: str,
    dest_range: str,
    copy_values: bool = True,
    copy_styles: bool = True,
) -> str:
    """Copy cells from source range to destination within same or across sheets."""
    wb = load_workbook_safe(file_path)
    try:
        ws_src = get_sheet(wb, source_sheet)
        ws_dst = get_sheet(wb, dest_sheet)

        # Parse source range
        if ":" in source_range:
            src_start, src_end = source_range.split(":")
        else:
            src_start = src_end = source_range

        src_cells = list(ws_src[f"{src_start}:{src_end}"])

        # Parse destination top-left cell
        dest_cell = ws_dst[dest_range.split(":")[0]]
        dest_start_row = dest_cell.row
        dest_start_col = dest_cell.column

        cells_copied = 0
        for r_offset, row in enumerate(src_cells):
            for c_offset, cell in enumerate(row):
                dst = ws_dst.cell(
                    row=dest_start_row + r_offset,
                    column=dest_start_col + c_offset,
                )
                if copy_values:
                    dst.value = cell.value
                if copy_styles:
                    dst.font = copy.copy(cell.font)
                    dst.fill = copy.copy(cell.fill)
                    dst.border = copy.copy(cell.border)
                    dst.number_format = cell.number_format
                    dst.alignment = copy.copy(cell.alignment)
                cells_copied += 1

        save_workbook_safe(wb, file_path)
        logger.info(
            "Copied %d cells from %s!%s to %s!%s",
            cells_copied,
            source_sheet,
            source_range,
            dest_sheet,
            dest_range,
        )
        return f"Copied {cells_copied} cells from '{source_sheet}'!{source_range} to '{dest_sheet}'!{dest_range}."
    finally:
        wb.close()


def delete_range(
    file_path: str,
    sheet_name: str,
    range_str: str,
    shift_direction: str = "up",
) -> str:
    """Delete range contents and shift remaining cells up or left."""
    if shift_direction not in ("up", "left"):
        return f"Invalid shift_direction '{shift_direction}'. Use 'up' or 'left'."

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        if ":" in range_str:
            start_ref, end_ref = range_str.split(":")
        else:
            start_ref = end_ref = range_str

        start = ws[start_ref]
        end = ws[end_ref]
        min_row, max_row = start.row, end.row
        min_col, max_col = start.column, end.column
        num_rows = max_row - min_row + 1
        num_cols = max_col - min_col + 1

        if shift_direction == "up":
            ws.delete_rows(min_row, num_rows)
        else:
            ws.delete_cols(min_col, num_cols)

        save_workbook_safe(wb, file_path)
        logger.info(
            "Deleted range %s in %s, shifted %s",
            range_str,
            sheet_name,
            shift_direction,
        )
        return f"Deleted range {range_str} in '{sheet_name}' and shifted cells {shift_direction}."
    finally:
        wb.close()


def get_file_info(file_path: str) -> dict:
    """Analyze a file and return metadata to help decide how to read it."""
    import os

    path = validate_file_path(file_path)
    size_bytes = os.path.getsize(path)

    if size_bytes < 1024:
        size_display = f"{size_bytes} bytes"
    elif size_bytes < 1024 * 1024:
        size_display = f"{size_bytes / 1024:.1f} KB"
    else:
        size_display = f"{size_bytes / (1024 * 1024):.1f} MB"

    wb = load_workbook_safe(file_path, read_only=True)
    try:
        sheets_info = []
        total_cells = 0
        for name in wb.sheetnames:
            ws = wb[name]
            max_row = ws.max_row or 0
            max_col = ws.max_column or 0
            cell_count = max_row * max_col
            total_cells += cell_count
            sheets_info.append(
                {
                    "name": name,
                    "max_row": max_row,
                    "max_column": max_col,
                    "cell_count": cell_count,
                }
            )

        estimated_memory_mb = round(total_cells * 50 / (1024 * 1024), 2)
        if total_cells <= 10000:
            recommended_chunk = 1000
        elif total_cells <= 100000:
            recommended_chunk = 500
        else:
            recommended_chunk = 200

        return {
            "file_path": str(path),
            "size_bytes": size_bytes,
            "size_display": size_display,
            "sheet_count": len(wb.sheetnames),
            "sheets": sheets_info,
            "total_cells": total_cells,
            "estimated_memory_mb": estimated_memory_mb,
            "recommended_chunk_size": recommended_chunk,
        }
    finally:
        wb.close()


def read_ranges_batch(
    file_path: str,
    sheet_name: str,
    ranges: list[str],
    include_empty: bool = True,
) -> dict:
    """Read multiple non-contiguous ranges in a single workbook load."""
    wb = load_workbook_safe(file_path, read_only=True, data_only=True)
    try:
        ws = get_sheet(wb, sheet_name)
        results: dict[str, list[list]] = {}
        for range_str in ranges:
            raw = ws[range_str]
            # Single cell returns a Cell object; wrap it
            if not isinstance(raw, tuple):
                value = raw.value if raw.value is not None else ""
                results[range_str] = [[value]]
            else:
                rows: list[list] = []
                for row in raw:
                    if isinstance(row, tuple):
                        rows.append([c.value if c.value is not None else "" for c in row])
                    else:
                        rows.append([row.value if row.value is not None else ""])
                results[range_str] = rows
        logger.info("Read %d ranges from %s!%s", len(ranges), sheet_name, file_path)
        return {"results": results}
    finally:
        wb.close()
