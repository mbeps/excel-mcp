"""Data cleaning pipeline operations."""

from __future__ import annotations

import re
from collections.abc import Callable
from logging import Logger

import pandas as pd

from mcp_server.utils.excel_helpers import (
    load_workbook_safe,
    read_sheet_df,
    save_workbook_safe,
    validate_file_path,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

ALL_OPERATIONS = [
    "trim_whitespace",
    "remove_empty_rows",
    "remove_empty_columns",
    "normalize_text",
    "fix_numbers",
    "remove_duplicates",
    "fill_missing",
]


def _col_letters_to_names(df: pd.DataFrame, columns: list[str] | None) -> list[str] | None:
    """Convert column letters like ['A', 'B'] to actual DataFrame column names."""
    if columns is None:
        return None
    result = []
    all_cols = list(df.columns)
    for letter in columns:
        idx = 0
        value = 0
        for ch in letter.upper():
            value = value * 26 + (ord(ch) - ord("A") + 1)
        idx = value - 1
        if idx < 0 or idx >= len(all_cols):
            raise ValueError(f"Column letter '{letter}' is out of range. Sheet has {len(all_cols)} columns.")
        result.append(all_cols[idx])
    return result


def _get_target_cols(df: pd.DataFrame, col_names: list[str] | None) -> list[str]:
    """Return target column names, or all columns if none specified."""
    if col_names is None:
        return list(df.columns)
    return col_names


def _trim_whitespace(df: pd.DataFrame, target_cols: list[str]) -> tuple[pd.DataFrame, int]:
    count = 0
    for col in target_cols:
        col_dtype = df[col].dtype
        if col_dtype is object:
            # Object dtype may contain mixed types — only process actual strings
            str_mask = df[col].apply(lambda x: isinstance(x, str))
            if not str_mask.any():
                continue
            original = df[col][str_mask]
            stripped = original.str.strip()
            count += int((original != stripped).sum())
            df.loc[str_mask, col] = stripped
        elif isinstance(col_dtype, pd.StringDtype):
            # Pandas 2.2+ nullable StringDtype — fully vectorized
            stripped = df[col].str.strip()
            mask = df[col].notna() & (df[col] != stripped)
            count += int(mask.sum())
            df[col] = stripped
    return df, count


def _remove_empty_rows(df: pd.DataFrame, target_cols: list[str]) -> tuple[pd.DataFrame, int]:
    before = len(df)
    mask = df[target_cols].isna().all(axis=1) | df[target_cols].apply(
        lambda row: all((pd.isna(v) or (isinstance(v, str) and v.strip() == "")) for v in row),
        axis=1,
    )
    df = df[~mask].reset_index(drop=True)
    return df, before - len(df)


def _remove_empty_columns(df: pd.DataFrame, target_cols: list[str]) -> tuple[pd.DataFrame, int]:
    count = 0
    cols_to_drop = []
    for col in target_cols:
        if df[col].isna().all() or df[col].apply(lambda x: isinstance(x, str) and x.strip() == "").all():
            cols_to_drop.append(col)
            count += 1
    df = df.drop(columns=cols_to_drop)
    return df, count


def _normalize_text(df: pd.DataFrame, target_cols: list[str]) -> tuple[pd.DataFrame, int]:
    count = 0
    for col in target_cols:
        if df[col].dtype is not object and not isinstance(df[col].dtype, pd.StringDtype):
            continue
        original = df[col].copy()
        # Vectorised: strip whitespace, lowercase, collapse internal whitespace
        cleaned = df[col].astype(str).where(df[col].notna(), other=pd.NA)
        cleaned = cleaned.str.strip().str.lower().str.replace(r"\s+", " ", regex=True)
        # Restore non-string / NA values
        cleaned = cleaned.where(original.notna(), other=original)
        changed = (original != cleaned) & original.notna()
        count += int(changed.sum())
        df[col] = cleaned
    return df, count


def _fix_numbers(df: pd.DataFrame, target_cols: list[str]) -> tuple[pd.DataFrame, int]:
    count = 0
    for col in target_cols:
        conversions: list[tuple[int, int | float]] = []
        for idx in df.index:
            val = df.at[idx, col]
            if not isinstance(val, str):
                continue
            cleaned = val.strip().replace(",", "").replace("$", "").replace("€", "").replace("£", "")
            if cleaned == "":
                continue
            try:
                num = int(cleaned) if "." not in cleaned else float(cleaned)
                conversions.append((idx, num))
            except ValueError:
                continue
        if conversions:
            df[col] = df[col].astype(object)
            for idx, num in conversions:
                df.at[idx, col] = num
            count += len(conversions)
    return df, count


def _remove_duplicates(df: pd.DataFrame, target_cols: list[str]) -> tuple[pd.DataFrame, int]:
    before = len(df)
    df = df.drop_duplicates(subset=target_cols, keep="first").reset_index(drop=True)
    return df, before - len(df)


def _fill_missing(df: pd.DataFrame, target_cols: list[str], fill_value: str = "") -> tuple[pd.DataFrame, int]:
    count = 0
    for col in target_cols:
        # Convert column to object dtype so we can safely assign fill_value
        if df[col].dtype != object:
            df[col] = df[col].astype(object)
        for idx in df.index:
            val = df.at[idx, col]
            if pd.isna(val) or (isinstance(val, str) and val.strip() == ""):
                df.at[idx, col] = fill_value
                count += 1
    return df, count


def _fill_missing_with_strategy(
    df: pd.DataFrame, target_cols: list[str], strategy: str, fill_value: str = ""
) -> tuple[pd.DataFrame, int]:
    """Fill missing values using a named strategy."""
    if strategy == "value":
        return _fill_missing(df, target_cols, fill_value=fill_value)

    count = 0
    if strategy == "ffill":
        before = int(df[target_cols].isna().sum().sum())
        for col in target_cols:
            df[col] = df[col].ffill()
        after = int(df[target_cols].isna().sum().sum())
        count = before - after
    elif strategy == "bfill":
        before = int(df[target_cols].isna().sum().sum())
        for col in target_cols:
            df[col] = df[col].bfill()
        after = int(df[target_cols].isna().sum().sum())
        count = before - after
    elif strategy == "mean":
        for col in target_cols:
            if pd.api.types.is_numeric_dtype(df[col]):
                na_count = int(df[col].isna().sum())
                if na_count > 0:
                    df[col] = df[col].fillna(df[col].mean())
                    count += na_count
    elif strategy == "median":
        for col in target_cols:
            if pd.api.types.is_numeric_dtype(df[col]):
                na_count = int(df[col].isna().sum())
                if na_count > 0:
                    df[col] = df[col].fillna(df[col].median())
                    count += na_count
    elif strategy == "mode":
        for col in target_cols:
            mode_series = df[col].mode()
            if not mode_series.empty:
                na_count = int(df[col].isna().sum())
                if na_count > 0:
                    df[col] = df[col].fillna(mode_series.iloc[0])
                    count += na_count
    else:
        raise ValueError(
            f"Unknown fill_missing_strategy '{strategy}'. Allowed: ffill, bfill, mean, median, mode, value"
        )
    return df, count


_OPERATION_MAP: dict[str, Callable[[pd.DataFrame, list[str]], tuple[pd.DataFrame, int]]] = {
    "trim_whitespace": _trim_whitespace,
    "remove_empty_rows": _remove_empty_rows,
    "remove_empty_columns": _remove_empty_columns,
    "normalize_text": _normalize_text,
    "fix_numbers": _fix_numbers,
    "remove_duplicates": _remove_duplicates,
}


def _collect_sample_changes(
    original: pd.DataFrame, cleaned: pd.DataFrame, max_samples: int = 10
) -> list[dict[str, str | int]]:
    """Collect sample before/after changes for preview mode."""
    samples: list[dict] = []
    shared_cols = [c for c in original.columns if c in cleaned.columns]
    compare_len = min(len(original), len(cleaned))

    for col in shared_cols:
        if len(samples) >= max_samples:
            break
        for i in range(compare_len):
            if len(samples) >= max_samples:
                break
            orig_val = original.iloc[i][col]
            clean_val = cleaned.iloc[i][col]
            orig_str = str(orig_val) if not pd.isna(orig_val) else "<empty>"
            clean_str = str(clean_val) if not pd.isna(clean_val) else "<empty>"
            if orig_str != clean_str:
                samples.append(
                    {
                        "row": i + 1,
                        "column": col,
                        "before": orig_str,
                        "after": clean_str,
                    }
                )
    return samples


def data_cleaner(
    file_path: str,
    sheet_name: str = "Sheet1",
    operations: list[str] | None = None,
    columns: list[str] | None = None,
    preview: bool = False,
    output_file: str | None = None,
    header_row: int = 1,
    fill_missing_strategy: str = "value",
    fill_value: str | None = None,
) -> dict[str, object]:
    """Run a configurable data cleaning pipeline on a sheet."""
    validate_file_path(file_path)

    ops = operations if operations is not None else list(ALL_OPERATIONS)
    valid_ops = set(ALL_OPERATIONS)
    for op in ops:
        if op not in valid_ops:
            raise ValueError(f"Unknown operation '{op}'. Allowed: {ALL_OPERATIONS}")

    df = read_sheet_df(file_path, sheet_name, header_row)
    original_df = df.copy()
    rows_before = len(df)

    col_names = _col_letters_to_names(df, columns)
    changes: dict[str, int] = {}

    for op in ops:
        target = _get_target_cols(df, col_names)
        # Filter target cols to those still present in df
        target = [c for c in target if c in df.columns]
        if not target:
            changes[op] = 0
            continue
        if op == "fill_missing":
            resolved_fill = fill_value if fill_value is not None else ""
            df, count = _fill_missing_with_strategy(df, target, fill_missing_strategy, fill_value=resolved_fill)
        else:
            df, count = _OPERATION_MAP[op](df, target)
        changes[op] = count

    rows_after = len(df)

    result: dict[str, object] = {
        "operations_applied": ops,
        "changes": changes,
        "rows_before": rows_before,
        "rows_after": rows_after,
        "preview_only": preview,
    }

    if preview:
        result["sample_changes"] = _collect_sample_changes(original_df, df)
        return result

    # Write output
    save_path = output_file if output_file else file_path
    if output_file:
        validate_file_path(output_file, must_exist=False)

    if save_path == file_path:
        wb = load_workbook_safe(file_path)
        # Clear and rewrite the target sheet
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
                for cell in row:
                    cell.value = None
        else:
            ws = wb.create_sheet(title=sheet_name)
    else:
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name

    try:
        # Write headers
        for c_idx, col_name in enumerate(df.columns, start=1):
            ws.cell(row=1, column=c_idx, value=col_name)

        # Write data
        for r_idx, row_data in enumerate(df.values.tolist(), start=2):
            for c_idx, val in enumerate(row_data, start=1):
                # Convert numpy types to native Python for openpyxl
                if hasattr(val, "item"):
                    val = val.item()
                ws.cell(row=r_idx, column=c_idx, value=val)

        save_workbook_safe(wb, save_path)
        logger.info("Cleaned data written to %s", save_path)
    finally:
        wb.close()

    return result


# ---------------------------------------------------------------------------
# Private helpers for new standalone functions
# ---------------------------------------------------------------------------


def _resolve_col(df: pd.DataFrame, col: str) -> str:
    """Resolve a column reference (letter like 'A' or direct column name) to a df column name."""
    if col in df.columns:
        return col
    if re.match(r"^[A-Za-z]+$", col):
        resolved = _col_letters_to_names(df, [col])
        if resolved:
            return resolved[0]
    raise ValueError(f"Column '{col}' not found. Available columns: {list(df.columns)}")


def _write_df_to_workbook(
    df: pd.DataFrame,
    file_path: str,
    sheet_name: str,
    output_file: str | None,
) -> str:
    """Write DataFrame back to a workbook sheet; returns the save path."""
    save_path = output_file if output_file else file_path
    if output_file:
        validate_file_path(output_file, must_exist=False)

    if save_path == file_path:
        wb = load_workbook_safe(file_path)
    else:
        import openpyxl

        wb = openpyxl.Workbook()

    try:
        if save_path == file_path:
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
                    for cell in row:
                        cell.value = None
            else:
                ws = wb.create_sheet(title=sheet_name)
        else:
            ws = wb.active
            ws.title = sheet_name

        for c_idx, col_name in enumerate(df.columns, start=1):
            ws.cell(row=1, column=c_idx, value=col_name)

        for r_idx, row_data in enumerate(df.values.tolist(), start=2):
            for c_idx, val in enumerate(row_data, start=1):
                if hasattr(val, "item"):
                    val = val.item()
                ws.cell(row=r_idx, column=c_idx, value=val)

        save_workbook_safe(wb, save_path)
    finally:
        wb.close()

    return save_path


# ---------------------------------------------------------------------------
# Public standalone functions
# ---------------------------------------------------------------------------


def split_column(
    file_path: str,
    sheet_name: str = "Sheet1",
    column: str = "A",
    delimiter: str = ",",
    new_columns: list[str] | None = None,
    drop_original: bool = True,
    output_file: str | None = None,
    header_row: int = 1,
) -> dict[str, str | int | list[str]]:
    """Split a text column into multiple columns by delimiter.

    Returns: {new_columns: list[str], rows_affected: int, output_file: str}
    """
    validate_file_path(file_path)
    df = read_sheet_df(file_path, sheet_name, header_row)
    col_name = _resolve_col(df, column)

    split_df = df[col_name].astype(str).str.split(delimiter, expand=True)
    num_parts = split_df.shape[1]

    if new_columns:
        names = list(new_columns[:num_parts])
        # Pad with auto-generated names if fewer than num_parts were supplied
        for i in range(len(names), num_parts):
            names.append(f"{col_name}_{i + 1}")
    else:
        names = [f"{col_name}_{i + 1}" for i in range(num_parts)]

    split_df.columns = pd.Index(names)

    col_idx = df.columns.get_loc(col_name)
    for i, new_col in enumerate(names):
        df.insert(col_idx + 1 + i, new_col, split_df[new_col])

    if drop_original:
        df = df.drop(columns=[col_name])

    save_path = _write_df_to_workbook(df, file_path, sheet_name, output_file)
    logger.info("split_column: split '%s' into %d columns in %s", col_name, num_parts, save_path)

    return {
        "new_columns": names,
        "rows_affected": len(df),
        "output_file": save_path,
    }


def parse_date_column(
    file_path: str,
    sheet_name: str,
    column: str,
    output_column: str | None = None,
    output_format: str = "%Y-%m-%d",
    dayfirst: bool = False,
    header_row: int = 1,
) -> dict[str, str | int]:
    """Parse and normalize mixed date formats in a column to a standard format."""
    from datetime import datetime

    from dateutil import parser as dateutil_parser

    df = read_sheet_df(file_path, sheet_name, header_row)
    column = _resolve_col(df, column)
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found. Available: {list(df.columns)}")

    parsed_values: list[datetime | None] = []
    for val in df[column]:
        if pd.isna(val) or (isinstance(val, str) and val.strip() == ""):
            parsed_values.append(None)
        elif isinstance(val, (datetime, pd.Timestamp)):
            parsed_values.append(val if isinstance(val, datetime) else val.to_pydatetime())
        else:
            try:
                parsed_values.append(dateutil_parser.parse(str(val), dayfirst=dayfirst))
            except (ValueError, OverflowError):
                parsed_values.append(None)

    parsed = pd.Series(parsed_values, index=df.index)
    formatted = parsed.apply(lambda v: v.strftime(output_format) if v is not None and not pd.isna(v) else None)

    write_col = output_column if output_column else column

    wb = load_workbook_safe(file_path)
    try:
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else None
        if ws is None:
            raise ValueError(f"Sheet '{sheet_name}' not found.")
        header_cells = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]
        if write_col in header_cells:
            col_idx = header_cells.index(write_col) + 1
        else:
            col_idx = ws.max_column + 1
            ws.cell(row=header_row, column=col_idx, value=write_col)
        for r_idx, v in enumerate(formatted.tolist(), start=header_row + 1):
            ws.cell(row=r_idx, column=col_idx, value=v)
        save_workbook_safe(wb, file_path)
    finally:
        wb.close()

    parsed_count = int(parsed.notna().sum())
    failed_count = int(parsed.isna().sum())  # None values count as NaN in pd.Series
    logger.info("parse_date_column: parsed %d/%d dates in column '%s'", parsed_count, len(df), column)
    return {"column": column, "parsed_count": parsed_count, "failed_count": failed_count, "output_column": write_col}
