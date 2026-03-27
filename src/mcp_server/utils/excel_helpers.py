"""Shared helpers for Excel file operations."""

from __future__ import annotations

import re
from logging import Logger
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv", ".xlsm"}
MAX_ROWS_DEFAULT = 10000
MAX_ROWS_WRITE = 50000


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


def load_workbook_safe(
    file_path: str,
    read_only: bool = False,
    data_only: bool = False,
    keep_vba: bool | None = None,
) -> Workbook:
    """Load an Excel workbook with validation and error handling."""
    path = validate_file_path(file_path)
    effective_keep_vba = keep_vba if keep_vba is not None else (path.suffix.lower() == ".xlsm")
    try:
        return openpyxl.load_workbook(str(path), read_only=read_only, data_only=data_only, keep_vba=effective_keep_vba)
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


def read_sheet_df(file_path: str, sheet_name: str, header_row: int = 1) -> pd.DataFrame:
    """Read a sheet into a pandas DataFrame.

    Args:
        header_row: 1-based row number of the header. 0 means no header.
    """
    path = validate_file_path(file_path)
    header = header_row - 1 if header_row >= 1 else None
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        try:
            return pd.read_excel(path, sheet_name=sheet_name, header=header, engine="calamine")
        except Exception:
            pass
    return pd.read_excel(path, sheet_name=sheet_name, header=header, engine="openpyxl")


_CELL_REF_RE = re.compile(r"^([A-Z]{1,3})(\d+)$", re.IGNORECASE)
MAX_COL_INDEX = 16384  # XFD
MAX_ROW = 1048576


def _validate_single_cell(ref: str) -> dict:
    """Validate a single cell reference like 'A1' or 'XFD1048576'."""
    m = _CELL_REF_RE.match(ref)
    if not m:
        return {"valid": False, "message": f"Invalid cell reference format: '{ref}'"}
    col_letters = m.group(1).upper()
    row_num = int(m.group(2))
    try:
        col_idx = column_index_from_string(col_letters)
    except ValueError:
        return {"valid": False, "message": f"Invalid column letters: '{col_letters}'"}
    if col_idx > MAX_COL_INDEX:
        return {"valid": False, "message": f"Column '{col_letters}' exceeds max (XFD)"}
    if row_num < 1 or row_num > MAX_ROW:
        return {"valid": False, "message": f"Row {row_num} out of range (1-{MAX_ROW})"}
    return {"valid": True, "column": col_letters, "row": row_num, "col_index": col_idx}


def validate_excel_range(range_str: str) -> dict:
    """Validate A1-style range notation.

    Accepts single cell refs ('A1') and ranges ('A1:C10').
    Returns dict with 'valid', 'message', and parsed info.
    """
    range_str = range_str.strip()
    if not range_str:
        return {"valid": False, "message": "Range string is empty"}

    parts = range_str.split(":")
    if len(parts) > 2:
        return {"valid": False, "message": f"Invalid range format: '{range_str}'"}

    start = _validate_single_cell(parts[0])
    if not start["valid"]:
        return start

    if len(parts) == 1:
        return {
            "valid": True,
            "message": "Valid single cell reference",
            "start_cell": parts[0].upper(),
            "end_cell": None,
            "is_range": False,
        }

    end = _validate_single_cell(parts[1])
    if not end["valid"]:
        return end

    return {
        "valid": True,
        "message": "Valid range",
        "start_cell": parts[0].upper(),
        "end_cell": parts[1].upper(),
        "is_range": True,
    }
