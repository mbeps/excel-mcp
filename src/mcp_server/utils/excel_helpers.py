"""Shared helpers for Excel file operations."""

from __future__ import annotations

from logging import Logger
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv", ".xlsm"}
MAX_ROWS_DEFAULT = 10000


def validate_file_path(file_path: str, must_exist: bool = True) -> Path:
    """Validate and resolve a file path. Raises ValueError for invalid paths."""
    path = Path(file_path).resolve()
    if must_exist and not path.exists():
        raise ValueError(f"File not found: {file_path}")
    if path.suffix.lower() not in ALLOWED_EXTENSIONS:
        raise ValueError(f"Unsupported file type: {path.suffix}. Allowed: {ALLOWED_EXTENSIONS}")
    if must_exist and not path.is_file():
        raise ValueError(f"Path is not a file: {file_path}")
    if not must_exist:
        path.parent.mkdir(parents=True, exist_ok=True)
    return path


def load_workbook_safe(file_path: str, read_only: bool = False, data_only: bool = False) -> Workbook:
    """Load an Excel workbook with validation and error handling."""
    path = validate_file_path(file_path)
    try:
        return openpyxl.load_workbook(str(path), read_only=read_only, data_only=data_only)
    except Exception as e:
        raise ValueError(f"Failed to open workbook '{file_path}': {e}") from e


def save_workbook_safe(wb: Workbook, file_path: str) -> None:
    """Save a workbook with error handling."""
    path = validate_file_path(file_path, must_exist=False)
    try:
        wb.save(str(path))
    except Exception as e:
        raise ValueError(f"Failed to save workbook '{file_path}': {e}") from e


def get_sheet(wb: Workbook, sheet_name: str) -> Worksheet:
    """Get a worksheet by name, raising ValueError if not found."""
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
    return wb[sheet_name]


def col_letter_to_index(letter: str) -> int:
    """Convert column letter(s) to 1-based index. A=1, B=2, ..., Z=26, AA=27."""
    return column_index_from_string(letter)


def index_to_col_letter(index: int) -> str:
    """Convert 1-based column index to letter(s). 1=A, 2=B, ..., 26=Z, 27=AA."""
    return get_column_letter(index)
