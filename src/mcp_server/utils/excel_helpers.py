"""Utilities for safe and convenient Excel/CSV file operations.

This module provides small, focused helpers used across the MCP server to
validate file paths, safely load and save workbooks, convert Excel column
notations, and read worksheets into pandas DataFrames. Functions here enforce
whitelisted file extensions and conservative defaults to avoid excessive memory
use in typical server environments.
"""

from __future__ import annotations

import copy
import json
import os
import re
import tempfile
from logging import Logger
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
from openpyxl.utils import column_index_from_string
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from mcp_server.models.workbook import ValidationRangeResult
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

"""
ALLOWED_EXTENSIONS (set[str]): Whitelisted file extensions accepted by helpers in this module.

Purpose:
    Restrict accepted workbook and CSV file formats so callers and downstream libraries
    (openpyxl / pandas / calamine) can rely on predictable parsing behaviour.

Typical values:
    {'.xlsx', '.xls', '.csv', '.xlsm'}
"""
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv", ".xlsm"}

_SERVER_TEMP_PREFIX = os.path.join(tempfile.gettempdir(), "excel_mcp_")


def validate_file_path(file_path: str, must_exist: bool = True) -> Path:
    """
    Validate and resolve a filesystem path to an allowed Excel/CSV file.

    Args:
        file_path (str): The path (absolute or relative) to validate.
        must_exist (bool): If True (default), the function requires the file to exist.
            If False, the parent directory will be created if necessary and no exist check
            is performed.

    Returns:
        pathlib.Path: Resolved Path object pointing to the validated file.

    Raises:
        ValueError: If the file does not exist (when must_exist=True), the path is not a
            regular file, the file extension is not in `ALLOWED_EXTENSIONS`, or the path is
            outside directories permitted by the `EXCEL_MCP_ALLOWED_DIRS` environment variable.
    """
    path = Path(file_path).resolve()
    _is_server_temp = str(path).startswith(_SERVER_TEMP_PREFIX)

    if path.suffix.lower() not in ALLOWED_EXTENSIONS:
        raise ValueError(f"Unsupported file type: {path.suffix}. Allowed: {ALLOWED_EXTENSIONS}")
    if must_exist and not path.exists():
        raise ValueError(f"File not found: {file_path}")
    if must_exist and not path.is_file():
        raise ValueError(f"Path is not a file: {file_path}")

    if _is_server_temp:
        return path

    allowed_dirs_env = os.environ.get("EXCEL_MCP_ALLOWED_DIRS", "").strip()
    if allowed_dirs_env:
        allowed_dirs = [Path(d.strip()).resolve() for d in allowed_dirs_env.split(",") if d.strip()]
        if not any(path == ad or ad in path.parents for ad in allowed_dirs):
            raise ValueError("File path is outside allowed directories")

    if not must_exist:
        path.parent.mkdir(parents=True, exist_ok=True)
    return path


def load_workbook_safe(
    file_path: str,
    read_only: bool = False,
    data_only: bool = False,
    keep_vba: bool | None = None,
) -> Workbook:
    """
    Load an Excel workbook with validation and robust error handling.

    Args:
        file_path (str): Path to the workbook to open.
        read_only (bool): If True, openpyxl will open the workbook in read-only mode.
        data_only (bool): If True, formulas will not be evaluated; cell values are the cached
            results saved in the file.
        keep_vba (bool | None): If True, preserve VBA content; if None the function will
            preserve VBA for `.xlsm` files and not for others.

    Returns:
        openpyxl.workbook.Workbook: Loaded workbook instance.

    Raises:
        ValueError: If `validate_file_path` fails or openpyxl fails to open the workbook.
    """
    path = validate_file_path(file_path)
    effective_keep_vba = keep_vba if keep_vba is not None else (path.suffix.lower() == ".xlsm")
    try:
        return openpyxl.load_workbook(str(path), read_only=read_only, data_only=data_only, keep_vba=effective_keep_vba)
    except Exception as e:
        raise ValueError(f"Failed to open workbook '{file_path}': {e}") from e


def save_workbook_safe(wb: Workbook, file_path: str) -> None:
    """
    Save an openpyxl `Workbook` to disk with validation and error wrapping.

    Args:
        wb (openpyxl.workbook.Workbook): Workbook instance to save.
        file_path (str): Target path. Parent directories will be created if necessary.

    Returns:
        None

    Raises:
        ValueError: If the path validation fails or saving raises an exception.
    """
    path = validate_file_path(file_path, must_exist=False)
    try:
        wb.save(str(path))
    except Exception as e:
        raise ValueError(f"Failed to save workbook '{file_path}': {e}") from e


def get_sheet(wb: Workbook, sheet_name: str) -> Worksheet:
    """
    Return a worksheet by name from a loaded `Workbook`.

    Args:
        wb (openpyxl.workbook.Workbook): Open workbook.
        sheet_name (str): Name of the worksheet to retrieve.

    Returns:
        openpyxl.worksheet.worksheet.Worksheet: The requested worksheet object.

    Raises:
        ValueError: If the sheet name is not present; the error message should include
            the available sheet names.
    """
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
    return wb[sheet_name]


def col_letter_to_index(letter: str) -> int:
    """
    Convert Excel column letter(s) to a 1-based column index.

    Examples:
        'A' -> 1, 'Z' -> 26, 'AA' -> 27

    Args:
        letter (str): Column letters in A1 notation.

    Returns:
        int: 1-based column index.
    """
    return int(column_index_from_string(letter))


def read_sheet_df(file_path: str, sheet_name: str, header_row: int | bool = 1) -> pd.DataFrame:
    """
    Read a worksheet into a pandas DataFrame, preferring the 'calamine' engine for speed
    and falling back to 'openpyxl' when needed.

    Args:
        file_path (str): Path to the workbook or CSV file.
        sheet_name (str): Name of the worksheet to read.
        header_row (int): 1-based row number containing headers. Use 0 to indicate there is
            no header row.

    Returns:
        pandas.DataFrame: DataFrame containing the sheet's data.

    Raises:
        ValueError: If `validate_file_path` rejects the path.

    Notes:
        The function attempts to use the 'calamine' engine for Excel files, with an
        OpenPyXL fallback for cases the first attempt fails.
    """
    path = validate_file_path(file_path)
    header = header_row - 1 if header_row >= 1 else None
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        try:
            return pd.read_excel(path, sheet_name=sheet_name, header=header, engine="calamine")
        except Exception:
            pass
    return pd.read_excel(path, sheet_name=sheet_name, header=header, engine="openpyxl")


_CELL_REF_RE: re.Pattern[str] = re.compile(r"^([A-Z]{1,3})(\d+)$", re.IGNORECASE)

"""
MAX_COL_INDEX (int): Maximum supported Excel column index (1-based). Corresponds to
column 'XFD'.

Value:
    16384
"""
MAX_COL_INDEX: int = 16384  # XFD

"""
MAX_ROW (int): Maximum supported Excel row index.

Value:
    1048576
"""
MAX_ROW: int = 1048576


def _validate_single_cell(ref: str) -> ValidationRangeResult:
    """Validate a single cell reference like 'A1' or 'XFD1048576'."""
    m = _CELL_REF_RE.match(ref)
    if not m:
        return {"valid": False, "message": f"Invalid cell reference format: '{ref}'"}
    col_letters: str = m.group(1).upper()
    row_num: int = int(m.group(2))
    try:
        col_idx: int = column_index_from_string(col_letters)
    except ValueError:
        return {"valid": False, "message": f"Invalid column letters: '{col_letters}'"}
    if col_idx > MAX_COL_INDEX:
        return {"valid": False, "message": f"Column '{col_letters}' exceeds max (XFD)"}
    if row_num < 1 or row_num > MAX_ROW:
        return {"valid": False, "message": f"Row {row_num} out of range (1-{MAX_ROW})"}
    return {"valid": True, "column": col_letters, "row": row_num, "col_index": col_idx}


def validate_excel_range(range_str: str) -> ValidationRangeResult:
    """Validate A1-style range notation.

    Accepts single cell refs ('A1') and ranges ('A1:C10').
    Returns dict with 'valid', 'message', and parsed info.
    """
    range_str = range_str.strip()
    if not range_str:
        return {"valid": False, "message": "Range string is empty"}

    parts: list[str] = range_str.split(":")
    if len(parts) > 2:
        return {"valid": False, "message": f"Invalid range format: '{range_str}'"}

    start = _validate_single_cell(parts[0])
    if not start["valid"]:
        return {
            "valid": False,
            "message": str(start.get("message", "Unknown error")),
        }

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
        return {
            "valid": False,
            "message": str(end.get("message", "Unknown error")),
        }

    return {
        "valid": True,
        "message": "Valid range",
        "start_cell": parts[0].upper(),
        "end_cell": parts[1].upper(),
        "is_range": True,
    }


def copy_cell_style(src_cell: Any, dest_cell: Any) -> None:
    """Copy all style attributes from src_cell to dest_cell."""
    dest_cell.font = copy.copy(src_cell.font)
    dest_cell.fill = copy.copy(src_cell.fill)
    dest_cell.border = copy.copy(src_cell.border)
    dest_cell.alignment = copy.copy(src_cell.alignment)
    dest_cell.number_format = src_cell.number_format


def load_hidden_json(wb: Workbook, sheet_name: str) -> dict:
    """Read and deserialise JSON stored in cell A1 of a hidden sheet.

    Returns empty dict if the sheet doesn't exist or the cell is empty.
    """
    if sheet_name not in wb.sheetnames:
        return {}
    ws = wb[sheet_name]
    raw = ws["A1"].value
    if not raw:
        return {}
    try:
        return json.loads(str(raw))  # type: ignore[no-any-return]
    except (json.JSONDecodeError, TypeError):
        return {}


def save_hidden_json(wb: Workbook, sheet_name: str, data: dict) -> None:
    """Serialise data to JSON and write to cell A1 of a hidden sheet (creating if needed)."""
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)
        ws.sheet_state = "hidden"
    ws["A1"] = json.dumps(data)
