"""Data cleaning pipeline operations."""

from __future__ import annotations

import re
from collections.abc import Callable
from logging import Logger

import pandas as pd

from mcp_server.utils.excel_helpers import (
    load_workbook_safe,
    read_sheet_df,
    resolve_dataframe_column,
    sanitize_excel_value,
    save_workbook_safe,
    validate_file_path,
    write_dataframe_to_workbook,
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


def _get_target_cols(df: pd.DataFrame, col_names: list[str] | None) -> list[str]:
    """Return target column names, or all columns if none specified."""
    if col_names is None:
        return list(df.columns)
    return col_names


def _trim_whitespace(df: pd.DataFrame, target_cols: list[str]) -> tuple[pd.DataFrame, int]:
    count = 0
    for col in target_cols:
        col_dtype = df[col].dtype
        if pd.api.types.is_object_dtype(col_dtype):
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
        if df[col].apply(lambda x: pd.isna(x) or (isinstance(x, str) and x.strip() == "")).all():
            cols_to_drop.append(col)
            count += 1
    df = df.drop(columns=cols_to_drop)
    return df, count


def _normalize_text(df: pd.DataFrame, target_cols: list[str]) -> tuple[pd.DataFrame, int]:
    count = 0
    for col in target_cols:
        col_dtype = df[col].dtype
        if pd.api.types.is_object_dtype(col_dtype):
            str_mask = df[col].apply(lambda x: isinstance(x, str))
            if not str_mask.any():
                continue
            original = df[col][str_mask]
            cleaned = original.str.strip().str.lower().str.replace(r"\s+", " ", regex=True)
            count += int((original != cleaned).sum())
            df.loc[str_mask, col] = cleaned
        elif isinstance(col_dtype, pd.StringDtype):
            original = df[col]
            cleaned = original.str.strip().str.lower().str.replace(r"\s+", " ", regex=True)
            changed = original.notna() & (original != cleaned)
            count += int(changed.sum())
            df[col] = cleaned
    return df, count


def _parse_numeric_string(raw: str) -> int | float | None:
    """Parse a currency/formatted/scientific numeric string, or return None."""
    cleaned = raw.strip()
    for sym in ("$", "€", "£", "¥"):
        cleaned = cleaned.replace(sym, "")
    cleaned = cleaned.strip()
    if cleaned == "":
        return None
    if "." in cleaned and "," in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        if cleaned.count(",") == 1 and len(cleaned.split(",")[1]) != 3:
            cleaned = cleaned.replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    try:
        if "." in cleaned or "e" in cleaned.lower():
            return float(cleaned)
        return int(cleaned)
    except ValueError:
        return None


def _fix_numbers(df: pd.DataFrame, target_cols: list[str]) -> tuple[pd.DataFrame, int]:
    count = 0
    for col in target_cols:
        conversions: list[tuple[int, int | float]] = []
        for idx in df.index:
            val = df.at[idx, col]
            if not isinstance(val, str):
                continue
            num = _parse_numeric_string(val)
            if num is not None:
                conversions.append((idx, num))
        if conversions:
            df[col] = df[col].astype(object)
            for idx, num in conversions:
                df.at[idx, col] = num
            count += len(conversions)
            if all(pd.isna(v) or (isinstance(v, (int, float)) and not isinstance(v, bool)) for v in df[col]):
                df[col] = pd.to_numeric(df[col])
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
            if pd.api.types.is_numeric_dtype(df[col]) and not pd.api.types.is_bool_dtype(df[col]):
                na_count = int(df[col].isna().sum())
                if na_count > 0:
                    fill_val = df[col].mean()
                    if pd.notna(fill_val):
                        df[col] = df[col].fillna(fill_val)
                        count += na_count
    elif strategy == "median":
        for col in target_cols:
            if pd.api.types.is_numeric_dtype(df[col]) and not pd.api.types.is_bool_dtype(df[col]):
                na_count = int(df[col].isna().sum())
                if na_count > 0:
                    fill_val = df[col].median()
                    if pd.notna(fill_val):
                        df[col] = df[col].fillna(fill_val)
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
    # Normalize aliases
    ops = [("trim_whitespace" if o == "trim" else o) for o in ops]
    valid_ops = set(ALL_OPERATIONS)
    for op in ops:
        if op not in valid_ops:
            raise ValueError(f"Unknown operation '{op}'. Allowed: {ALL_OPERATIONS}")

    df = read_sheet_df(file_path, sheet_name, header_row)
    original_df = df.copy()
    rows_before = len(df)

    col_names: list[str] | None
    if columns is None:
        col_names = None
    else:
        col_names = [str(resolve_dataframe_column(df, col_ref)) for col_ref in columns]
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
    save_path = write_dataframe_to_workbook(
        df,
        file_path,
        sheet_name=sheet_name,
        output_file=output_file,
        start_row=header_row,
        include_header=True,
    )
    logger.info("Cleaned data written to %s", save_path)

    return result


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
    col_name = str(resolve_dataframe_column(df, column))

    str_series = df[col_name].where(df[col_name].notna(), other=pd.NA).astype("string")
    split_df = str_series.str.split(delimiter, expand=True)
    num_parts = split_df.shape[1]

    if new_columns:
        names = list(new_columns[:num_parts])
        # Pad with auto-generated names if fewer than num_parts were supplied
        for i in range(len(names), num_parts):
            names.append(f"{col_name}_{i + 1}")
    else:
        names = [f"{col_name}_{i + 1}" for i in range(num_parts)]

    split_df.columns = pd.Index(names)

    col_idx = int(df.columns.get_loc(col_name))
    for i, new_col in enumerate(names):
        df.insert(col_idx + 1 + i, new_col, split_df[new_col])

    if drop_original:
        df = df.drop(columns=[col_name])

    save_path = str(
        write_dataframe_to_workbook(
            df,
            file_path,
            sheet_name=sheet_name,
            output_file=output_file,
            start_row=header_row,
            include_header=True,
        )
    )
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

    validate_file_path(file_path)
    df = read_sheet_df(file_path, sheet_name, header_row)
    col_key = str(resolve_dataframe_column(df, column))

    parsed_values: list[datetime | None] = []
    raw_values = df[col_key].tolist()
    for val in raw_values:
        if pd.isna(val) or (isinstance(val, str) and val.strip() == ""):
            parsed_values.append(None)
        elif isinstance(val, (datetime, pd.Timestamp)):
            parsed_values.append(val if isinstance(val, datetime) else val.to_pydatetime())
        elif isinstance(val, (int, float, bool)):
            parsed_values.append(None)
        else:
            s = str(val).strip()
            if re.fullmatch(r"[+-]?\d+(?:\.\d+)?", s) and not (len(s) == 8 and s.isdigit()):
                parsed_values.append(None)
                continue
            try:
                # ISO format (YYYY-MM-DD...) is unambiguous — parse without dayfirst
                if re.match(r"^\d{4}-\d{2}-\d{2}", s):
                    parsed_values.append(dateutil_parser.parse(s))
                else:
                    parsed_values.append(dateutil_parser.parse(s, dayfirst=dayfirst))
            except (ValueError, OverflowError):
                parsed_values.append(None)

    parsed = pd.Series(parsed_values, index=df.index)
    write_col = output_column if output_column else col_key

    wb = load_workbook_safe(file_path)
    try:
        ws = wb[sheet_name]
        header_cells = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]
        if write_col in header_cells:
            col_idx = header_cells.index(write_col) + 1
        else:
            col_idx = ws.max_column + 1
            ws.cell(row=header_row, column=col_idx, value=write_col)
        for i, (parsed_dt, raw_val) in enumerate(zip(parsed_values, raw_values)):
            r_idx = header_row + 1 + i
            if parsed_dt is not None:
                out_val: object = parsed_dt.strftime(output_format)
            elif output_column is None and not (
                pd.isna(raw_val) or (isinstance(raw_val, str) and raw_val.strip() == "")
            ):
                out_val = sanitize_excel_value(raw_val)
            else:
                out_val = None
            ws.cell(row=r_idx, column=col_idx, value=out_val)
        save_workbook_safe(wb, file_path)
    finally:
        wb.close()

    parsed_count = int(parsed.notna().sum())
    failed_count = int(parsed.isna().sum())  # None values count as NaN in pd.Series
    logger.info("parse_date_column: parsed %d/%d dates in column '%s'", parsed_count, len(df), col_key)
    return {"column": col_key, "parsed_count": parsed_count, "failed_count": failed_count, "output_column": write_col}
