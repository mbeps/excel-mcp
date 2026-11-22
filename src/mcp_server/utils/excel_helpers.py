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
import math
import os
import re
import tempfile
import threading
import weakref
from logging import Logger
from pathlib import Path
from typing import Any

import numpy as np
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

_SERVER_TEMP_DIR: Path = Path(tempfile.gettempdir()).resolve()


def validate_file_path(
    file_path: str,
    must_exist: bool = True,
    allowed_extensions: set[str] | None = None,
) -> Path:
    """
    Validate and resolve a filesystem path to an allowed Excel/CSV file.

    Args:
        file_path (str): The path (absolute or relative) to validate.
        must_exist (bool): If True (default), the function requires the file to exist.
            If False, the parent directory will be created if necessary and no exist check
            is performed.
        allowed_extensions (set[str] | None): Optional extension whitelist overriding
            ``ALLOWED_EXTENSIONS``.

    Returns:
        pathlib.Path: Resolved Path object pointing to the validated file.

    Raises:
        ValueError: If the path is a URL, is outside ``EXCEL_MCP_ALLOWED_DIRS``, the file
            extension is not allowed, or the file does not exist / is not a regular file.
    """
    if file_path.strip().lower().startswith(("http://", "https://")):
        raise ValueError(
            "Expected a local workspace file path, not a URL. Call upload_file first to stage remote files."
        )

    path = Path(file_path).resolve()
    _is_server_temp = path.parent == _SERVER_TEMP_DIR and path.name.startswith("excel_mcp_")

    effective_extensions = allowed_extensions if allowed_extensions is not None else ALLOWED_EXTENSIONS
    if path.suffix.lower() not in effective_extensions:
        raise ValueError(f"Unsupported file type: {path.suffix}. Allowed: {effective_extensions}")

    if not _is_server_temp:
        allowed_dirs_env = os.environ.get("EXCEL_MCP_ALLOWED_DIRS", "").strip()
        if allowed_dirs_env:
            allowed_dirs = [Path(d.strip()).resolve() for d in allowed_dirs_env.split(",") if d.strip()]
            if not any(path == ad or ad in path.parents for ad in allowed_dirs):
                raise ValueError("File path is outside allowed directories")

    if must_exist and not path.exists():
        raise ValueError(f"File not found: {file_path}")
    if must_exist and not path.is_file():
        raise ValueError(f"Path is not a file: {file_path}")

    if not must_exist and not _is_server_temp:
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


_FILE_LOCKS_GUARD: threading.Lock = threading.Lock()
# ponytail: in-process per-file mutex. Ceiling: single process only. Upgrade path: use filelock if multi-process MCP deployment needed.
_FILE_LOCKS: weakref.WeakValueDictionary[Path, threading.Lock] = weakref.WeakValueDictionary()


def _get_file_lock(path: Path) -> threading.Lock:
    """Return a per-file threading Lock to serialize concurrent writes to the same workbook."""
    with _FILE_LOCKS_GUARD:
        lock = _FILE_LOCKS.get(path)
        if lock is None:
            lock = threading.Lock()
            _FILE_LOCKS[path] = lock
        return lock


_EXCEL_SAVE_EXTENSIONS: frozenset[str] = frozenset({".xlsx", ".xlsm"})


def save_workbook_safe(wb: Workbook, file_path: str) -> None:
    """
    Save an openpyxl `Workbook` to disk with validation and robust atomic error handling.

    Args:
        wb (openpyxl.workbook.Workbook): Workbook instance to save.
        file_path (str): Target path. Parent directories will be created if necessary.

    Returns:
        None

    Raises:
        ValueError: If the workbook is read-only, the path validation fails, the target
            extension is not an Excel format, or saving raises an exception.
    """
    if getattr(wb, "read_only", False):
        raise ValueError("Cannot save a read-only workbook")
    path = validate_file_path(file_path, must_exist=False)
    if path.suffix.lower() not in _EXCEL_SAVE_EXTENSIONS:
        raise ValueError(f"Cannot save workbook to non-Excel extension: {path.suffix}")
    lock = _get_file_lock(path)
    with lock:
        tmp_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(dir=path.parent, delete=False, suffix=".tmp") as tmp:
                tmp_path = Path(tmp.name)
            wb.save(str(tmp_path))
            os.replace(tmp_path, path)
        except Exception as e:
            raise ValueError(f"Failed to save workbook '{file_path}': {e}") from e
        finally:
            if tmp_path is not None and tmp_path.exists():
                try:
                    tmp_path.unlink()
                except OSError:
                    pass


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


def read_sheet_df(file_path: str, sheet_name: str | int, header_row: int | bool = 1) -> pd.DataFrame:
    """
    Read a worksheet or CSV into a pandas DataFrame, preferring 'calamine' for Excel
    and falling back to 'openpyxl' for .xlsx/.xlsm when needed.

    Args:
        file_path (str): Path to the workbook or CSV file.
        sheet_name (str | int): Name or 0-based index of the worksheet to read (ignored for CSV).
        header_row (int): 1-based row number containing headers. Use 0 to indicate there is
            no header row.

    Returns:
        pandas.DataFrame: DataFrame containing the sheet's data.

    Raises:
        ValueError: If `validate_file_path` rejects the path, `header_row` is invalid,
            or reading fails.
    """
    if isinstance(header_row, bool) or header_row < 0:
        raise ValueError(f"header_row must be a non-negative integer, got {header_row!r}")
    path = validate_file_path(file_path)
    header = header_row - 1 if header_row >= 1 else None
    suffix = path.suffix.lower()

    if suffix == ".csv":
        for encoding in ("utf-8", "latin-1"):
            try:
                return pd.read_csv(path, header=header, encoding=encoding)
            except UnicodeDecodeError:
                continue
            except Exception as e:
                raise ValueError(f"Failed to read CSV '{file_path}': {e}") from e
        raise ValueError(f"Failed to decode CSV '{file_path}' as UTF-8 or latin-1")

    try:
        return pd.read_excel(path, sheet_name=sheet_name, header=header, engine="calamine")
    except Exception as calamine_err:
        if suffix not in (".xlsx", ".xlsm"):
            raise ValueError(
                f"Failed to read sheet '{sheet_name}' from '{file_path}': {calamine_err}"
            ) from calamine_err
    try:
        return pd.read_excel(path, sheet_name=sheet_name, header=header, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Failed to read sheet '{sheet_name}' from '{file_path}': {e}") from e


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
    col_idx: int = column_index_from_string(col_letters)
    if col_idx > MAX_COL_INDEX:
        return {"valid": False, "message": f"Column '{col_letters}' exceeds max (XFD)"}
    if row_num < 1 or row_num > MAX_ROW:
        return {"valid": False, "message": f"Row {row_num} out of range (1-{MAX_ROW})"}
    return {"valid": True, "column": col_letters, "row": row_num, "col_index": col_idx}


def validate_single_cell(cell_ref: str) -> str:
    """Validate that *cell_ref* is a single A1 cell reference (not a multi-cell range)."""
    cleaned = cell_ref.strip()
    if ":" in cleaned:
        raise ValueError(f"Expected a single cell reference, got range '{cell_ref}'.")
    res = validate_excel_range(cleaned)
    if not res["valid"]:
        raise ValueError(f"Invalid cell reference '{cell_ref}': {res['message']}")
    return cleaned.upper()


def validate_excel_range(range_str: str) -> ValidationRangeResult:
    """Validate A1-style range notation.

    Accepts single cell refs ('A1') and ranges ('A1:C10').
    Returns dict with 'valid', 'message', and parsed info.
    """
    range_str = range_str.strip()
    if not range_str:
        return {"valid": False, "message": "Range string is empty"}

    parts: list[str] = [p.strip() for p in range_str.split(":")]
    if len(parts) > 2:
        return {"valid": False, "message": f"Invalid range format: '{range_str}'"}

    start = _validate_single_cell(parts[0])
    if not start["valid"]:
        return {
            "valid": False,
            "message": start.get("message", "Unknown error"),
        }

    start_cell = f"{start['column']}{start['row']}"
    if len(parts) == 1:
        return {
            "valid": True,
            "message": "Valid single cell reference",
            "start_cell": start_cell,
            "end_cell": None,
            "is_range": False,
        }

    end = _validate_single_cell(parts[1])
    if not end["valid"]:
        return {
            "valid": False,
            "message": end.get("message", "Unknown error"),
        }

    if end["col_index"] < start["col_index"] or end["row"] < start["row"]:
        return {
            "valid": False,
            "message": f"Invalid range '{range_str}': end cell precedes start cell",
        }

    end_cell = f"{end['column']}{end['row']}"
    return {
        "valid": True,
        "message": "Valid range",
        "start_cell": start_cell,
        "end_cell": end_cell,
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
        return json.loads(str(raw))
    except (json.JSONDecodeError, TypeError):
        return {}


_EXCEL_CELL_CHAR_LIMIT = 32767


def save_hidden_json(wb: Workbook, sheet_name: str, data: dict) -> None:
    """Serialise data to JSON and write to cell A1 of a hidden sheet (creating if needed)."""
    payload = json.dumps(data)
    if len(payload) > _EXCEL_CELL_CHAR_LIMIT:
        raise ValueError(
            f"Metadata for '{sheet_name}' exceeds Excel's {_EXCEL_CELL_CHAR_LIMIT}-character cell limit "
            f"({len(payload)} chars)"
        )
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)
        ws.sheet_state = "hidden"
    ws["A1"] = payload


def sanitize_excel_value(val: Any) -> Any:
    """Convert NaN/pd.NA/pd.NaT to None and NumPy scalars to Python primitives."""
    if val is None or val is pd.NA or val is pd.NaT:
        return None
    if isinstance(val, np.bool_):
        return bool(val)
    if isinstance(val, bool):
        return val
    if isinstance(val, np.integer):
        return int(val)
    if isinstance(val, np.floating):
        f = float(val)
        return None if math.isnan(f) else f
    if isinstance(val, float):
        return None if math.isnan(val) else val
    if isinstance(val, pd.Timestamp):
        return None if pd.isna(val) else val.to_pydatetime()
    return val


_EXCEL_COL_RE: re.Pattern[str] = re.compile(r"^[A-Za-z]{1,3}$")


def resolve_dataframe_column(df: pd.DataFrame, col_ref: str) -> Any:
    """Resolve a column reference by exact header, case-insensitive header, or Excel letter."""
    if col_ref in df.columns:
        return col_ref
    for col in df.columns:
        if str(col) == col_ref:
            return col
    lower_ref = col_ref.lower()
    for col in df.columns:
        if str(col).lower() == lower_ref:
            return col
    stripped = col_ref.strip()
    if _EXCEL_COL_RE.match(stripped):
        idx = int(column_index_from_string(stripped.upper())) - 1
        if 0 <= idx < len(df.columns):
            return df.columns[idx]
    raise ValueError(f"Column '{col_ref}' not found. Available: {list(df.columns)}")


def write_dataframe_to_workbook(
    df: pd.DataFrame,
    file_path: str,
    sheet_name: str | None = None,
    output_file: str | None = None,
    start_row: int = 1,
    include_header: bool = True,
) -> Path:
    """Write a DataFrame to an Excel workbook, preserving other sheets on existing files."""
    src_path = validate_file_path(file_path)
    dest_str = output_file if output_file is not None else str(src_path)
    dest_path = validate_file_path(dest_str, must_exist=False)
    same_file = os.path.realpath(src_path) == os.path.realpath(dest_path)

    if same_file or dest_path.exists():
        wb = load_workbook_safe(str(dest_path))
        if sheet_name is not None and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        elif sheet_name is not None:
            ws = wb.create_sheet(sheet_name)
        else:
            ws = wb.active
            if ws is None:  # pragma: no cover
                ws = wb.create_sheet("Sheet1")
    else:
        wb = Workbook()
        ws = wb.active
        if ws is None:  # pragma: no cover
            ws = wb.create_sheet("Sheet1")
        if sheet_name is not None:
            ws.title = sheet_name

    try:
        if ws.max_row >= start_row:
            ws.delete_rows(start_row, ws.max_row - start_row + 1)

        if include_header:
            for c_idx, col_name in enumerate(df.columns, start=1):
                ws.cell(row=start_row, column=c_idx, value=sanitize_excel_value(col_name))

        data_start_row = start_row + 1 if include_header else start_row
        for r_offset, row_vals in enumerate(df.itertuples(index=False, name=None)):
            for c_idx, val in enumerate(row_vals, start=1):
                ws.cell(row=data_start_row + r_offset, column=c_idx, value=sanitize_excel_value(val))

        save_workbook_safe(wb, str(dest_path))
    finally:
        wb.close()

    return dest_path
