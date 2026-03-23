"""Data cleaning pipeline operations."""

from __future__ import annotations

import re
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
        for idx in df.index:
            val = df.at[idx, col]
            if isinstance(val, str) and val != val.strip():
                df.at[idx, col] = val.strip()
                count += 1
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
        for idx in df.index:
            val = df.at[idx, col]
            if isinstance(val, str):
                normalized = re.sub(r"\s+", " ", val.lower().strip())
                if val != normalized:
                    df.at[idx, col] = normalized
                    count += 1
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


def _fill_missing(df: pd.DataFrame, target_cols: list[str]) -> tuple[pd.DataFrame, int]:
    count = 0
    for col in target_cols:
        # Convert column to object dtype so we can safely assign "N/A"
        if df[col].dtype != object:
            df[col] = df[col].astype(object)
        for idx in df.index:
            val = df.at[idx, col]
            if pd.isna(val) or (isinstance(val, str) and val.strip() == ""):
                df.at[idx, col] = "N/A"
                count += 1
    return df, count


_OPERATION_MAP = {
    "trim_whitespace": _trim_whitespace,
    "remove_empty_rows": _remove_empty_rows,
    "remove_empty_columns": _remove_empty_columns,
    "normalize_text": _normalize_text,
    "fix_numbers": _fix_numbers,
    "remove_duplicates": _remove_duplicates,
    "fill_missing": _fill_missing,
}


def _collect_sample_changes(original: pd.DataFrame, cleaned: pd.DataFrame, max_samples: int = 10) -> list[dict]:
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
) -> dict:
    """Run a configurable data cleaning pipeline on a sheet."""
    validate_file_path(file_path)

    ops = operations if operations is not None else list(ALL_OPERATIONS)
    for op in ops:
        if op not in _OPERATION_MAP:
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
        df, count = _OPERATION_MAP[op](df, target)
        changes[op] = count

    rows_after = len(df)

    result: dict = {
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

    return result
