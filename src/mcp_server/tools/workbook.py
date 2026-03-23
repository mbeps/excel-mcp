"""Workbook-level operations: metadata, sheets, creation."""

from __future__ import annotations

from logging import Logger

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
    validate_file_path,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def get_workbook_metadata(file_path: str) -> dict:
    """Return workbook metadata: sheets, active sheet, and named ranges."""
    wb = load_workbook_safe(file_path, read_only=True)
    try:
        sheets = []
        for ws in wb.worksheets:
            sheets.append(
                {
                    "name": ws.title,
                    "min_row": ws.min_row,
                    "max_row": ws.max_row,
                    "min_col": ws.min_column,
                    "max_col": ws.max_column,
                }
            )

        named_ranges = [{"name": nr.name, "destination": str(nr.attr_text)} for nr in wb.defined_names.values()]

        return {
            "sheets": sheets,
            "active_sheet": wb.active.title if wb.active else None,
            "named_ranges": named_ranges,
        }
    finally:
        wb.close()


def list_sheets(file_path: str) -> list[dict]:
    """Return a list of sheet info dicts with name and dimensions."""
    wb = load_workbook_safe(file_path, read_only=True)
    try:
        return [
            {
                "name": ws.title,
                "min_row": ws.min_row,
                "max_row": ws.max_row,
                "min_col": ws.min_column,
                "max_col": ws.max_column,
            }
            for ws in wb.worksheets
        ]
    finally:
        wb.close()


def create_workbook(file_path: str, sheet_names: list[str] | None = None) -> dict:
    """Create a new .xlsx workbook with optional sheet names."""
    validate_file_path(file_path, must_exist=False)
    wb = Workbook()

    if sheet_names:
        for name in sheet_names:
            wb.create_sheet(title=name)
        if "Sheet" not in sheet_names and "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    save_workbook_safe(wb, file_path)
    logger.info("Created workbook: %s", file_path)
    return {"file_path": file_path, "sheets": wb.sheetnames}


def get_sheet_summary(file_path: str, sheet_name: str) -> dict:
    """Return summary of a sheet: name, dimensions, headers, used range."""
    wb = load_workbook_safe(file_path, read_only=True)
    try:
        ws = get_sheet(wb, sheet_name)
        min_row = ws.min_row or 1
        max_row = ws.max_row or 1
        min_col = ws.min_column or 1
        max_col = ws.max_column or 1

        headers = []
        for row in ws.iter_rows(min_row=min_row, max_row=min_row, min_col=min_col, max_col=max_col):
            headers = [cell.value for cell in row]

        used_range = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"

        return {
            "name": ws.title,
            "row_count": max_row - min_row + 1,
            "col_count": max_col - min_col + 1,
            "headers": [str(h) if h is not None else "" for h in headers],
            "used_range": used_range,
        }
    finally:
        wb.close()


def rename_sheet(file_path: str, old_name: str, new_name: str) -> str:
    """Rename a worksheet."""
    wb = load_workbook_safe(file_path)
    ws = get_sheet(wb, old_name)
    ws.title = new_name
    save_workbook_safe(wb, file_path)
    logger.info("Renamed sheet '%s' to '%s' in %s", old_name, new_name, file_path)
    return f"Sheet '{old_name}' renamed to '{new_name}'."


def delete_sheet(file_path: str, sheet_name: str) -> str:
    """Delete a worksheet. Raises ValueError if it's the only sheet."""
    wb = load_workbook_safe(file_path)
    if len(wb.sheetnames) == 1:
        raise ValueError("Cannot delete the only sheet in the workbook.")
    ws = get_sheet(wb, sheet_name)
    del wb[ws.title]
    save_workbook_safe(wb, file_path)
    logger.info("Deleted sheet '%s' from %s", sheet_name, file_path)
    return f"Sheet '{sheet_name}' deleted."


def copy_sheet(file_path: str, source_sheet: str, new_name: str) -> str:
    """Copy a sheet within the same workbook."""
    wb = load_workbook_safe(file_path)
    ws = get_sheet(wb, source_sheet)
    target = wb.copy_worksheet(ws)
    target.title = new_name
    save_workbook_safe(wb, file_path)
    logger.info("Copied sheet '%s' as '%s' in %s", source_sheet, new_name, file_path)
    return f"Sheet '{source_sheet}' copied as '{new_name}'."
