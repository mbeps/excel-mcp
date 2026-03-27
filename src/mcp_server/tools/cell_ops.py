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


def read_cell(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
    include_formula: bool = False,
    include_metadata: bool = False,
) -> dict:
    """Read a single cell. Set include_formula=True to get the stored formula string.
    Set include_metadata=True to also return is_merged, has_comment, has_hyperlink, number_format."""
    wb = load_workbook_safe(file_path, read_only=not include_metadata)
    try:
        ws = get_sheet(wb, sheet_name)
        cell = ws[cell_ref]
        result = {
            "cell_ref": cell_ref,
            "value": cell.value,
            "data_type": cell.data_type,
        }
        if include_formula and isinstance(cell.value, str) and cell.value.startswith("="):
            result["formula"] = cell.value
        elif include_formula:
            result["formula"] = None
        if include_metadata:
            is_merged = any(cell.coordinate in mr for mr in ws.merged_cells.ranges)
            result["is_merged"] = is_merged
            result["has_comment"] = cell.comment is not None
            result["has_hyperlink"] = cell.hyperlink is not None
            result["number_format"] = cell.number_format
        return result
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


def fill_series(
    file_path: str,
    sheet_name: str,
    start_cell: str,
    series_type: str,
    count: int,
    step: float | str = 1,
    direction: str = "down",
    start_value: float | int | None = None,
) -> dict:
    """Fill a series of values starting from start_cell.

    series_type: 'number' (1,2,3...), 'date' (requires step as offset like '1D','1M','1Y'),
                 'text_increment' (A1, A2, A3...), 'custom' (fill count cells with the step value)
    direction: 'down' or 'right'
    """
    import re as _re

    from openpyxl.utils import column_index_from_string, get_column_letter
    from openpyxl.utils.cell import coordinate_from_string

    if direction not in ("down", "right"):
        raise ValueError("direction must be 'down' or 'right'")

    col_letter, row_num = coordinate_from_string(start_cell.upper())
    col_idx = column_index_from_string(col_letter)

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        start_val = ws.cell(row=row_num, column=col_idx).value
        if start_value is not None:
            start_val = start_value

        values_written: list = []

        if series_type == "number":
            step_val = float(step)
            if start_val is None:
                start_val = 0
            current = float(start_val)
            for i in range(count):
                raw = current + step_val * i
                write_val: float | int = int(raw) if raw == int(raw) else raw
                if direction == "down":
                    ws.cell(row=row_num + i, column=col_idx, value=write_val)
                else:
                    ws.cell(row=row_num, column=col_idx + i, value=write_val)
                values_written.append(write_val)

        elif series_type == "date":
            import pandas as pd

            if start_val is None:
                raise ValueError("start_cell must contain a date value for series_type='date'")
            dates = pd.date_range(start=start_val, periods=count, freq=str(step))
            for i, dt in enumerate(dates):
                val = dt.to_pydatetime()
                if direction == "down":
                    ws.cell(row=row_num + i, column=col_idx, value=val)
                else:
                    ws.cell(row=row_num, column=col_idx + i, value=val)
                values_written.append(str(dt.date()))

        elif series_type == "text_increment":
            if start_val is None:
                raise ValueError("start_cell must contain a text value for series_type='text_increment'")
            text_val = str(start_val)
            m = _re.match(r"^(.*?)(\d+)$", text_val)
            if not m:
                raise ValueError(f"Cannot extract trailing number from '{text_val}' for text_increment series")
            prefix = m.group(1)
            num = int(m.group(2))
            width = len(m.group(2))
            for i in range(count):
                val_str = f"{prefix}{str(num + i).zfill(width)}"
                if direction == "down":
                    ws.cell(row=row_num + i, column=col_idx, value=val_str)
                else:
                    ws.cell(row=row_num, column=col_idx + i, value=val_str)
                values_written.append(val_str)

        elif series_type == "custom":
            for i in range(count):
                if direction == "down":
                    ws.cell(row=row_num + i, column=col_idx, value=step)
                else:
                    ws.cell(row=row_num, column=col_idx + i, value=step)
                values_written.append(step)

        else:
            raise ValueError(
                f"Invalid series_type '{series_type}'. Must be 'number', 'date', 'text_increment', or 'custom'"
            )

        if direction == "down":
            end_cell = f"{col_letter}{row_num + count - 1}"
        else:
            end_cell = f"{get_column_letter(col_idx + count - 1)}{row_num}"

        save_workbook_safe(wb, file_path)
        logger.info("Filled %s series from %s to %s in %s!%s", series_type, start_cell, end_cell, sheet_name, file_path)
        return {"start_cell": start_cell, "end_cell": end_cell, "count": count, "values_written": values_written}
    finally:
        wb.close()
