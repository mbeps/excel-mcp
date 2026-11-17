"""Sheet and cell protection operations."""

from __future__ import annotations

from logging import Logger

from openpyxl.styles import Protection

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def protect_sheet(
    file_path: str,
    sheet_name: str,
    password: str | None = None,
    allow_formatting_cells: bool = False,
    allow_formatting_columns: bool = False,
    allow_formatting_rows: bool = False,
    allow_insert_columns: bool = False,
    allow_insert_rows: bool = False,
    allow_delete_columns: bool = False,
    allow_delete_rows: bool = False,
    allow_sort: bool = False,
    allow_filter: bool = False,
) -> str:
    """Enable sheet protection with configurable permissions."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        ws.protection.sheet = True
        if password:
            ws.protection.set_password(password)
        ws.protection.formatCells = allow_formatting_cells
        ws.protection.formatColumns = allow_formatting_columns
        ws.protection.formatRows = allow_formatting_rows
        ws.protection.insertColumns = allow_insert_columns
        ws.protection.insertRows = allow_insert_rows
        ws.protection.deleteColumns = allow_delete_columns
        ws.protection.deleteRows = allow_delete_rows
        ws.protection.sort = allow_sort
        ws.protection.autoFilter = allow_filter

        save_workbook_safe(wb, file_path)
        logger.info("Protected sheet '%s' in %s", sheet_name, file_path)
        return f"Sheet '{sheet_name}' is now protected."
    finally:
        wb.close()


def unprotect_sheet(file_path: str, sheet_name: str, password: str | None = None) -> str:
    """Remove sheet protection."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        ws.protection.sheet = False
        ws.protection._password = None  # noqa: SLF001  # clear without triggering hash
        ws.protection.algorithmName = None
        ws.protection.hashValue = None
        ws.protection.saltValue = None
        ws.protection.spinCount = None

        save_workbook_safe(wb, file_path)
        logger.info("Unprotected sheet '%s' in %s", sheet_name, file_path)
        return f"Sheet '{sheet_name}' protection removed."
    finally:
        wb.close()


def protect_cells(
    file_path: str,
    sheet_name: str,
    locked_range: str,
    unlocked_ranges: list[str] | None = None,
) -> str:
    """Lock specific cells and optionally unlock others.

    Note: cell locking only takes effect when sheet protection is enabled.
    """
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        locked_target = f"{locked_range}:{locked_range}" if ":" not in locked_range else locked_range
        for row in ws[locked_target]:
            cells = row if isinstance(row, tuple) else (row,)
            for cell in cells:
                existing_hidden = bool(cell.protection.hidden) if cell.protection else False
                cell.protection = Protection(locked=True, hidden=existing_hidden)

        if unlocked_ranges:
            for ur in unlocked_ranges:
                # Normalise single-cell refs to range format so ws[] returns a tuple
                target = f"{ur}:{ur}" if ":" not in ur else ur
                for row in ws[target]:
                    cells = row if isinstance(row, tuple) else (row,)
                    for cell in cells:
                        existing_hidden = bool(cell.protection.hidden) if cell.protection else False
                        cell.protection = Protection(locked=False, hidden=existing_hidden)

        save_workbook_safe(wb, file_path)
        locked_msg = f"Locked '{locked_range}'"
        unlocked_msg = f", unlocked {unlocked_ranges}" if unlocked_ranges else ""
        logger.info("Updated cell protection on sheet '%s' in %s", sheet_name, file_path)
        return f"{locked_msg}{unlocked_msg} on sheet '{sheet_name}'. Enable sheet protection for this to take effect."
    finally:
        wb.close()
