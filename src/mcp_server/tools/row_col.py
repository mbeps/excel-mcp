"""Row and column insertion/deletion operations."""

from __future__ import annotations

from logging import Logger

from mcp_server.utils.excel_helpers import (
    col_letter_to_index,
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def insert_rows(file_path: str, sheet_name: str, row_index: int, count: int = 1) -> str:
    """Insert empty rows at the given 1-based position."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.insert_rows(row_index, amount=count)
        save_workbook_safe(wb, file_path)
        logger.info("Inserted %d row(s) at row %d in '%s'", count, row_index, sheet_name)
        return f"Inserted {count} row(s) at row {row_index} in '{sheet_name}'."
    finally:
        wb.close()


def delete_rows(file_path: str, sheet_name: str, row_index: int, count: int = 1) -> str:
    """Delete rows starting at the given 1-based position."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.delete_rows(row_index, amount=count)
        save_workbook_safe(wb, file_path)
        logger.info("Deleted %d row(s) at row %d in '%s'", count, row_index, sheet_name)
        return f"Deleted {count} row(s) at row {row_index} in '{sheet_name}'."
    finally:
        wb.close()


def insert_cols(file_path: str, sheet_name: str, col_index: int, count: int = 1) -> str:
    """Insert empty columns at the given 1-based position."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.insert_cols(col_index, amount=count)
        save_workbook_safe(wb, file_path)
        logger.info("Inserted %d column(s) at col %d in '%s'", count, col_index, sheet_name)
        return f"Inserted {count} column(s) at column {col_index} in '{sheet_name}'."
    finally:
        wb.close()


def delete_cols(file_path: str, sheet_name: str, col_index: int, count: int = 1) -> str:
    """Delete columns starting at the given 1-based position."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.delete_cols(col_index, amount=count)
        save_workbook_safe(wb, file_path)
        logger.info("Deleted %d column(s) at col %d in '%s'", count, col_index, sheet_name)
        return f"Deleted {count} column(s) at column {col_index} in '{sheet_name}'."
    finally:
        wb.close()


def insert_columns_by_letter(file_path: str, sheet_name: str, column: str, count: int = 1) -> str:
    """Insert columns before the specified column letter (e.g. 'C')."""
    col_index = col_letter_to_index(column.upper())
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.insert_cols(col_index, amount=count)
        save_workbook_safe(wb, file_path)
        logger.info("Inserted %d column(s) at column %s in '%s'", count, column.upper(), sheet_name)
        return f"Inserted {count} column(s) at column {column.upper()} in '{sheet_name}'."
    finally:
        wb.close()


def delete_columns_by_letter(file_path: str, sheet_name: str, column: str, count: int = 1) -> str:
    """Delete columns starting at the specified column letter (e.g. 'C')."""
    col_index = col_letter_to_index(column.upper())
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.delete_cols(col_index, amount=count)
        save_workbook_safe(wb, file_path)
        logger.info("Deleted %d column(s) at column %s in '%s'", count, column.upper(), sheet_name)
        return f"Deleted {count} column(s) at column {column.upper()} in '{sheet_name}'."
    finally:
        wb.close()
