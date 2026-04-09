"""Low-level cell and range read/write helpers, including chunked reads and copy/paste semantics (supports paste_values_only).

Functions mutate workbooks when writing and use ``load_workbook_safe()``/``save_workbook_safe()`` for safety.
"""

from __future__ import annotations

import copy
import math
from logging import Logger

from mcp_server.models.cell_ops import CellStyleInfo
from mcp_server.models.common import CellScalar
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
) -> dict[str, CellScalar | str | bool | None]:
    """Read a single cell value with optional metadata and formula exposure.

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Worksheet name.
        cell_ref (str): A1-style cell reference (e.g. 'B2').
        include_formula (bool): If True return the stored formula string when present.
        include_metadata (bool): If True return merged/comment/hyperlink/number_format metadata.

    Returns:
        dict: {cell_ref, value, data_type, ...}.

    Raises:
        FileNotFoundError/PermissionError: on invalid file path or I/O.

    Remarks:
        - Uses ``load_workbook_safe()`` (may open read_only when include_metadata is False). Always closes workbook.
    """
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
    """Write a Python scalar value to a single cell and save the workbook.

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Worksheet to write.
        cell_ref (str): A1 cell reference.
        value (scalar): Value to write (int/float/str/bool/None).

    Returns:
        str: Success message.

    Raises:
        FileNotFoundError, PermissionError: on save.

    Remarks:
        - Mutates workbook; uses ``load_workbook_safe()`` and ``save_workbook_safe()``.
    """
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
) -> dict[str, list[list[CellScalar]] | list[list[CellStyleInfo]] | str | int | bool]:
    """Read a rectangular A1 range from a sheet returning rows as lists. Supports returning style metadata and HTML output.

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Worksheet to read.
        start_cell (str): Top-left cell.
        end_cell (str): Bottom-right cell.
        show_formula (bool): Return formula strings when present instead of evaluated values.
        show_style (bool): Include per-cell style metadata (requires non-read-only openpyxl workbook).
        output_format (str): 'json' (rows returned) or 'html' (HTML table string).
        max_cells (int | None): Maximum total cells to return (truncates rows if exceeded).

    Returns:
        dict: rows/styles/html, row_count, col_count, truncated flag.

    Raises:
        ValueError: on invalid cell coordinates.

    Remarks:
        - Uses ``load_workbook_safe()`` with read_only=False when styles requested.
        - Be cautious with very large ranges (use ``max_cells`` to limit memory usage).
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

        rows: list[list[CellScalar]] = []
        styles: list[list[CellStyleInfo]] = []
        for row in ws[f"{start_cell}:{read_end_cell}"]:
            row_values: list[CellScalar] = []
            row_styles: list[CellStyleInfo] = []
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


def write_range(file_path: str, sheet_name: str, start_cell: str, data: list[list[CellScalar]]) -> str:
    """Write a 2D list of values to a sheet starting at ``start_cell`` and save.

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Worksheet to write.
        start_cell (str): Top-left anchor for the provided 2D data.
        data (list[list[scalar]]): Rows to write.

    Returns:
        str: Summary of rows x cols written.

    Raises:
        FileNotFoundError/PermissionError: on write.

    Remarks:
        - Mutates workbook via openpyxl; calls ``save_workbook_safe()``.
    """
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
    """Clear all values in an A1 rectangular range.

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Worksheet name.
        start_cell (str): Top-left.
        end_cell (str): Bottom-right.

    Returns:
        str: Status message.

    Raises:
        FileNotFoundError/PermissionError: on save.

    Remarks:
        - Mutates workbook and saves changes.
    """
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
) -> dict[str, list[dict[str, CellScalar]] | int | bool | None]:
    """Read a worksheet in chunks using pandas to prevent huge memory use.

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Worksheet to read.
        start_row (int): 0-based start index.
        chunk_size (int): Number of rows to read.

    Returns:
        dict: rows as dictionaries, chunk metadata (has_more, next_start_row, total_rows, paging).

    Raises:
        ValueError: if chunk_size invalid.

    Remarks:
        - Uses pandas.read_excel with ``calamine`` or ``openpyxl`` engine; ``validate_file_path()`` is invoked.
    """
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
    paste_values_only: bool = False,
) -> str:
    """Copy cells from a source range to a destination range; can copy formulas, styles or paste values only.

    Args:
        file_path (str): Path to workbook.
        source_sheet (str): Source worksheet name.
        source_range (str): A1 source range (e.g. 'A1:C10').
        dest_sheet (str): Destination worksheet name.
        dest_range (str): Top-left target cell or range where content will be written.
        copy_values (bool): Whether to copy values (default True).
        copy_styles (bool): Whether to copy styles (default True).
        paste_values_only (bool): If True, use ``data_only=True`` pass to read computed values and paste those.

    Returns:
        str: Human readable summary including number of cells copied.

    Raises:
        ValueError: on invalid ranges or sheet names.
        FileNotFoundError/PermissionError: on write.

    Remarks:
        - When ``paste_values_only`` is True the function reads the source sheet with ``data_only=True`` to obtain computed values.
        - Uses ``load_workbook_safe()``/``save_workbook_safe()`` and always closes workbooks.
    """
    if ":" in source_range:
        src_start, src_end = source_range.split(":")
    else:
        src_start = src_end = source_range
    range_key = f"{src_start}:{src_end}"

    # Pre-read computed values when paste_values_only is requested
    computed: list[list] | None = None
    if paste_values_only:
        wb_data = load_workbook_safe(file_path, data_only=True)
        try:
            ws_data = get_sheet(wb_data, source_sheet)
            computed = [[c.value for c in row] for row in ws_data[range_key]]
        finally:
            wb_data.close()

    wb = load_workbook_safe(file_path)
    try:
        ws_src = get_sheet(wb, source_sheet)
        ws_dst = get_sheet(wb, dest_sheet)

        src_cells = list(ws_src[range_key])

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
                    dst.value = computed[r_offset][c_offset] if paste_values_only else cell.value  # type: ignore[index]
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
) -> dict[str, str | int | list[CellScalar]]:
    """Auto-fill a sequence of cells with numeric, date, text-increment or custom series types.

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Worksheet to modify.
        start_cell (str): Anchor cell to begin the series.
        series_type (str): 'number', 'date', 'text_increment', or 'custom'.
        count (int): Number of values to generate.
        step (float | str): Step increment (numeric) or offset spec for dates (e.g. '1D').
        direction (str): 'down' or 'right'.
        start_value (optional): Override starting value.

    Returns:
        dict: {start_cell, end_cell, count, values_written}.

    Raises:
        ValueError: for invalid parameters or when start value types cannot be parsed.

    Remarks:
        - Mutates workbook and saves changes.
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

        values_written: list[CellScalar] = []

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


def merge_cells(file_path: str, sheet_name: str, range_string: str) -> dict[str, str]:
    """Merge or unmerge a given range of cells.

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Worksheet to modify.
        range_string (str): Range to merge/unmerge (e.g. 'A1:D1').

    Returns:
        dict: status/message.

    Raises:
        ValueError: if range invalid.

    Remarks:
        - Mutates workbook; uses ``validate_file_path()`` and ``load_workbook_safe()``.
    """
    validate_file_path(file_path)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.merge_cells(range_string)
        save_workbook_safe(wb, file_path)
        logger.info("Merged cells %s in %s!%s", range_string, sheet_name, file_path)
        return {"status": "success", "message": f"Merged cells {range_string}"}
    finally:
        wb.close()


def unmerge_cells(file_path: str, sheet_name: str, range_string: str) -> dict[str, str]:
    """Merge or unmerge a given range of cells.

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Worksheet to modify.
        range_string (str): Range to merge/unmerge (e.g. 'A1:D1').

    Returns:
        dict: status/message.

    Raises:
        ValueError: if range invalid.

    Remarks:
        - Mutates workbook; uses ``validate_file_path()`` and ``load_workbook_safe()``.
    """
    validate_file_path(file_path)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.unmerge_cells(range_string)
        save_workbook_safe(wb, file_path)
        logger.info("Unmerged cells %s in %s!%s", range_string, sheet_name, file_path)
        return {"status": "success", "message": f"Unmerged cells {range_string}"}
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# New cell-operation features
# ---------------------------------------------------------------------------


def fill_formula(
    file_path: str,
    sheet_name: str,
    source_cell: str,
    target_range: str,
) -> dict[str, int | str]:
    """Translate a source formula to every cell in a target range using openpyxl's Translator (excel-like drag-fill).

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Worksheet containing the formula to copy.
        source_cell (str): Cell with the source formula (must start with '=').
        target_range (str): A1 range to fill.

    Returns:
        dict: {'cells_filled': int, 'source_formula': str}.

    Raises:
        ValueError: if source_cell does not contain a formula.

    Remarks:
        - Uses openpyxl.formula.translate.Translator and mutates workbook.
    """
    from openpyxl.formula.translate import Translator

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        source_formula = ws[source_cell.upper()].value
        if not isinstance(source_formula, str) or not source_formula.startswith("="):
            raise ValueError(f"Source cell {source_cell} does not contain a formula.")

        range_key = target_range if ":" in target_range else f"{target_range}:{target_range}"
        translator = Translator(source_formula, source_cell.upper())
        filled = 0
        for row in ws[range_key]:
            for cell in row:
                cell.value = translator.translate_formula(cell.coordinate)
                filled += 1

        save_workbook_safe(wb, file_path)
        logger.info("fill_formula: filled %d cells in %s", filled, sheet_name)
        return {"cells_filled": filled, "source_formula": source_formula}
    finally:
        wb.close()


def find_replace(
    file_path: str,
    sheet_name: str,
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    match_entire_cell: bool = False,
    search_formulas: bool = False,
) -> dict[str, int | list[str]]:
    """Find and replace text in cell values (and optionally formulas) across a sheet.

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Worksheet to scan.
        find_text (str): Text to search for.
        replace_text (str): Replacement text.
        match_case (bool): Case-sensitive matching.
        match_entire_cell (bool): Replace only when full cell equals needle.
        search_formulas (bool): If True search & replace inside formula strings.

    Returns:
        dict: {'replacements_made': int, 'cells_modified': list[str]}.

    Remarks:
        - Mutates workbook and saves; uses ``load_workbook_safe()`` and ``save_workbook_safe()`` only when replacements occurred.
    """
    import re as _re

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        modified: list[str] = []
        needle = find_text if match_case else find_text.lower()
        flags = 0 if match_case else _re.IGNORECASE

        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if val is None:
                    continue
                is_formula = isinstance(val, str) and val.startswith("=")
                if is_formula and not search_formulas:
                    continue
                content = str(val)
                compare = content if match_case else content.lower()

                if match_entire_cell:
                    if compare == needle:
                        cell.value = replace_text
                        modified.append(cell.coordinate)
                else:
                    if needle in compare:
                        cell.value = _re.sub(_re.escape(find_text), replace_text, content, flags=flags)
                        modified.append(cell.coordinate)

        if modified:
            save_workbook_safe(wb, file_path)
        logger.info("find_replace: %d replacements in %s", len(modified), sheet_name)
        return {"replacements_made": len(modified), "cells_modified": modified}
    finally:
        wb.close()


def transpose_range(
    file_path: str,
    sheet_name: str,
    source_range: str,
    target_cell: str,
    source_sheet: str | None = None,
    paste_values_only: bool = False,
) -> dict[str, list[int]]:
    """Transpose a rectangular source_range and write it starting at target_cell. Optionally uses computed values when paste_values_only=True.

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Destination worksheet name.
        source_range (str): Source A1 range.
        target_cell (str): Top-left destination cell.
        source_sheet (str | None): Source sheet name (defaults to sheet_name).
        paste_values_only (bool): If True use computed values (data_only read) instead of formulas.

    Returns:
        dict: {'source_shape': [rows, cols], 'target_shape': [cols, rows]}.

    Remarks:
        - Uses ``load_workbook_safe()`` possibly twice (for data_only read) and mutates workbook to write.
    """
    src_sheet = source_sheet if source_sheet is not None else sheet_name
    range_key = source_range if ":" in source_range else f"{source_range}:{source_range}"

    data: list[list]
    if paste_values_only:
        wb_data = load_workbook_safe(file_path, data_only=True)
        try:
            ws_data = get_sheet(wb_data, src_sheet)
            data = [[cell.value for cell in row] for row in ws_data[range_key]]
        finally:
            wb_data.close()

    wb = load_workbook_safe(file_path)
    try:
        ws_src = get_sheet(wb, src_sheet)
        ws_dst = get_sheet(wb, sheet_name)

        if not paste_values_only:
            data = [[cell.value for cell in row] for row in ws_src[range_key]]

        n_rows = len(data)
        n_cols = len(data[0]) if data else 0

        dst = ws_dst[target_cell.upper()]
        dst_row, dst_col = dst.row, dst.column

        for r in range(n_rows):
            for c in range(n_cols):
                ws_dst.cell(row=dst_row + c, column=dst_col + r, value=data[r][c])

        save_workbook_safe(wb, file_path)
        logger.info("transpose_range: %dx%d → %dx%d in %s", n_rows, n_cols, n_cols, n_rows, sheet_name)
        return {"source_shape": [n_rows, n_cols], "target_shape": [n_cols, n_rows]}
    finally:
        wb.close()


def auto_sum(
    file_path: str,
    sheet_name: str,
    cell: str,
    source_range: str | None = None,
) -> dict[str, str]:
    """Write a `=SUM(...)` formula into a target cell. If ``source_range`` is omitted the function heuristically finds adjacent numeric cells to sum.

    Args:
        file_path (str): Path to workbook.
        sheet_name (str): Worksheet name.
        cell (str): Target cell to write formula.
        source_range (str | None): Optional explicit range to sum.

    Returns:
        dict: {'formula': str, 'cell': str}.

    Raises:
        ValueError: if no numeric neighbours are found and source_range was not provided.

    Remarks:
        - Mutates workbook and saves change.
    """
    from openpyxl.utils import column_index_from_string, get_column_letter
    from openpyxl.utils.cell import coordinate_from_string

    cell_upper = cell.upper()

    if source_range is not None:
        formula = f"=SUM({source_range})"
        wb = load_workbook_safe(file_path)
        try:
            ws = get_sheet(wb, sheet_name)
            ws[cell_upper] = formula
            save_workbook_safe(wb, file_path)
        finally:
            wb.close()
        return {"formula": formula, "cell": cell_upper}

    col_letter, row_num = coordinate_from_string(cell_upper)
    col_idx = column_index_from_string(col_letter)

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        # Scan upward for consecutive numeric cells
        top = row_num - 1
        while top >= 1 and isinstance(ws.cell(row=top, column=col_idx).value, (int, float)):
            top -= 1
        top += 1  # first numeric row, or row_num if none found

        if top < row_num:
            range_str = f"{col_letter}{top}:{col_letter}{row_num - 1}"
        else:
            # Scan left for consecutive numeric cells in the same row
            left = col_idx - 1
            while left >= 1 and isinstance(ws.cell(row=row_num, column=left).value, (int, float)):
                left -= 1
            left += 1

            if left < col_idx:
                range_str = f"{get_column_letter(left)}{row_num}:{get_column_letter(col_idx - 1)}{row_num}"
            else:
                raise ValueError(f"No numeric cells found adjacent to '{cell}'. Provide source_range explicitly.")

        formula = f"=SUM({range_str})"
        ws[cell_upper] = formula
        save_workbook_safe(wb, file_path)
        logger.info("auto_sum: wrote %s into %s!%s", formula, sheet_name, cell_upper)
    finally:
        wb.close()

    return {"formula": formula, "cell": cell_upper}
