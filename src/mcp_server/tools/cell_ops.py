"""Cell-level read/write operations and chunked reading."""

from __future__ import annotations

from logging import Logger

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
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


def read_range(file_path: str, sheet_name: str, start_cell: str, end_cell: str) -> dict:
    """Read a rectangular range and return rows as a list of lists."""
    wb = load_workbook_safe(file_path, read_only=True)
    try:
        ws = get_sheet(wb, sheet_name)
        rows = []
        for row in ws[f"{start_cell}:{end_cell}"]:
            rows.append([cell.value for cell in row])
        return {
            "rows": rows,
            "row_count": len(rows),
            "col_count": len(rows[0]) if rows else 0,
        }
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

    from mcp_server.utils.excel_helpers import validate_file_path

    path = validate_file_path(file_path)

    try:
        df = pd.read_excel(path, sheet_name=sheet_name, engine="calamine")
    except ImportError:
        logger.warning("calamine engine not available, falling back to openpyxl")
        df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")

    end_row = start_row + chunk_size
    chunk = df.iloc[start_row:end_row]

    return {
        "rows": chunk.to_dict(orient="records"),
        "chunk_start": start_row,
        "chunk_size": len(chunk),
        "has_more": end_row < len(df),
    }
