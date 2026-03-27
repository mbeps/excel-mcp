"""Cross-file operations: bulk aggregate, bulk filter, and data consistency validation."""

from __future__ import annotations

from logging import Logger

import pandas as pd

from mcp_server.utils.excel_helpers import (
    load_workbook_safe,
    read_sheet_df,
    validate_file_path,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

AGGREGATE_OPS = {"sum", "mean", "min", "max", "count"}
FILTER_OPERATORS = {"equals", "not_equals", "greater_than", "less_than", "contains"}


def _write_df_to_new_file(df: pd.DataFrame, output_file: str, sheet_name: str = "Sheet1") -> None:
    validate_file_path(output_file, must_exist=False)
    df.to_excel(output_file, index=False, sheet_name=sheet_name, engine="openpyxl")


def bulk_aggregate_multi_files(
    file_paths: list[str],
    column: str,
    operation: str = "sum",
    sheet_name: str = "Sheet1",
    header_row: int = 1,
    output_file: str | None = None,
) -> dict:
    """Aggregate the same column across multiple files."""
    if operation not in AGGREGATE_OPS:
        raise ValueError(f"Unsupported operation '{operation}'. Allowed: {AGGREGATE_OPS}")
    if not file_paths:
        raise ValueError("file_paths must not be empty.")

    per_file: list[dict] = []
    all_values: list[float] = []

    for fp in file_paths:
        df = read_sheet_df(fp, sheet_name, header_row)
        if column not in df.columns:
            raise ValueError(f"Column '{column}' not found in '{fp}'. Available: {list(df.columns)}")

        col_data = pd.to_numeric(df[column], errors="coerce").dropna()
        values = col_data.tolist()
        all_values.extend(values)

        file_result: dict = {"file": fp, "row_count": len(df)}
        if operation == "sum":
            file_result["value"] = float(col_data.sum())
        elif operation == "mean":
            file_result["value"] = float(col_data.mean()) if len(col_data) > 0 else 0.0
        elif operation == "min":
            file_result["value"] = float(col_data.min()) if len(col_data) > 0 else 0.0
        elif operation == "max":
            file_result["value"] = float(col_data.max()) if len(col_data) > 0 else 0.0
        elif operation == "count":
            file_result["value"] = len(col_data)

        per_file.append(file_result)

    series = pd.Series(all_values)
    if operation == "sum":
        aggregate = float(series.sum())
    elif operation == "mean":
        aggregate = float(series.mean()) if len(series) > 0 else 0.0
    elif operation == "min":
        aggregate = float(series.min()) if len(series) > 0 else 0.0
    elif operation == "max":
        aggregate = float(series.max()) if len(series) > 0 else 0.0
    else:  # count
        aggregate = len(series)

    result = {
        "column": column,
        "operation": operation,
        "per_file": per_file,
        "aggregate": aggregate,
        "total_files": len(file_paths),
    }

    if output_file:
        summary_df = pd.DataFrame(per_file)
        summary_df.loc[len(summary_df)] = {
            "file": "TOTAL",
            "row_count": sum(r["row_count"] for r in per_file),
            "value": aggregate,
        }
        _write_df_to_new_file(summary_df, output_file, sheet_name="Summary")
        logger.info("Aggregate summary written to %s", output_file)

    return result


def bulk_filter_multi_files(
    file_paths: list[str],
    column: str,
    operator: str,
    value: str | float,
    sheet_name: str = "Sheet1",
    header_row: int = 1,
    output_file: str | None = None,
) -> dict:
    """Filter rows across multiple files based on a condition."""
    if operator not in FILTER_OPERATORS:
        raise ValueError(f"Unsupported operator '{operator}'. Allowed: {FILTER_OPERATORS}")
    if not file_paths:
        raise ValueError("file_paths must not be empty.")

    per_file: list[dict] = []
    all_matched = pd.DataFrame()

    for fp in file_paths:
        df = read_sheet_df(fp, sheet_name, header_row)
        if column not in df.columns:
            raise ValueError(f"Column '{column}' not found in '{fp}'. Available: {list(df.columns)}")

        col = df[column]
        if operator == "equals":
            mask = col.astype(str) == str(value)
        elif operator == "not_equals":
            mask = col.astype(str) != str(value)
        elif operator == "greater_than":
            mask = pd.to_numeric(col, errors="coerce") > float(value)
        elif operator == "less_than":
            mask = pd.to_numeric(col, errors="coerce") < float(value)
        else:  # contains
            mask = col.astype(str).str.contains(str(value), case=False, na=False)

        matched = df[mask]
        per_file.append(
            {
                "file": fp,
                "total_rows": len(df),
                "matched_rows": len(matched),
                "data": matched.values.tolist(),
            }
        )
        if not matched.empty:
            tagged = matched.copy()
            tagged.insert(0, "_source_file", fp)
            all_matched = pd.concat([all_matched, tagged], ignore_index=True)

    total_matched = sum(f["matched_rows"] for f in per_file)

    result = {
        "column": column,
        "operator": operator,
        "value": value,
        "per_file": per_file,
        "total_matched": total_matched,
        "total_files": len(file_paths),
    }

    if output_file and not all_matched.empty:
        _write_df_to_new_file(all_matched, output_file, sheet_name="FilteredResults")
        logger.info("Filtered results written to %s", output_file)

    return result


def validate_data_consistency(
    file_paths: list[str],
    key_column: str,
    check_columns: list[str] | None = None,
    sheet_name: str = "Sheet1",
    header_row: int = 1,
) -> dict:
    """Cross-file referential integrity check on a key column."""
    if not file_paths:
        raise ValueError("file_paths must not be empty.")

    # Load all DataFrames and collect keys
    file_data: list[tuple[str, pd.DataFrame]] = []
    all_keys: set = set()

    for fp in file_paths:
        df = read_sheet_df(fp, sheet_name, header_row)
        if key_column not in df.columns:
            raise ValueError(f"Key column '{key_column}' not found in '{fp}'. Available: {list(df.columns)}")
        if check_columns:
            for cc in check_columns:
                if cc not in df.columns:
                    raise ValueError(f"Check column '{cc}' not found in '{fp}'. Available: {list(df.columns)}")
        file_data.append((fp, df))
        all_keys.update(df[key_column].dropna().unique().tolist())

    # Find missing keys per file
    missing_keys: dict[str, list] = {}
    for fp, df in file_data:
        file_keys = set(df[key_column].dropna().unique().tolist())
        missing = sorted(str(k) for k in all_keys - file_keys)
        if missing:
            missing_keys[fp] = missing

    # Check value consistency across files for matching keys
    mismatched_values: list[dict] = []
    if check_columns:
        # Build a lookup: key -> {file -> {col: value}}
        key_values: dict[str, dict[str, dict[str, object]]] = {}
        for fp, df in file_data:
            for _, row in df.iterrows():
                key = row[key_column]
                if pd.isna(key):
                    continue
                key_str = str(key)
                if key_str not in key_values:
                    key_values[key_str] = {}
                vals = {}
                for cc in check_columns:
                    v = row[cc]
                    vals[cc] = v if not pd.isna(v) else None
                key_values[key_str][fp] = vals

        for key_str, file_vals in key_values.items():
            if len(file_vals) < 2:
                continue
            files_list = list(file_vals.keys())
            for cc in check_columns:
                seen_values = {}
                for fp in files_list:
                    v = file_vals[fp].get(cc)
                    v_str = str(v) if v is not None else "<empty>"
                    seen_values[fp] = v_str
                unique_vals = set(seen_values.values())
                if len(unique_vals) > 1:
                    mismatched_values.append(
                        {
                            "key": key_str,
                            "column": cc,
                            "values_by_file": seen_values,
                        }
                    )

    consistent = len(missing_keys) == 0 and len(mismatched_values) == 0

    return {
        "consistent": consistent,
        "total_keys": len(all_keys),
        "total_files": len(file_paths),
        "missing_keys": missing_keys,
        "mismatched_values": mismatched_values,
    }


def compare_workbooks(
    file_path_a: str,
    file_path_b: str,
    sheet_name: str | None = None,
    output_file: str | None = None,
) -> dict:
    """Compare two workbooks cell-by-cell. If sheet_name given, compare only that sheet."""
    wb_a = load_workbook_safe(file_path_a, data_only=True)
    wb_b = load_workbook_safe(file_path_b, data_only=True)
    try:
        if sheet_name:
            if sheet_name not in wb_a.sheetnames or sheet_name not in wb_b.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' not found in both workbooks.")
            sheets_to_compare = [sheet_name]
        else:
            sheets_to_compare = [s for s in wb_a.sheetnames if s in wb_b.sheetnames]

        differences: list[dict] = []
        for sn in sheets_to_compare:
            ws_a = wb_a[sn]
            ws_b = wb_b[sn]
            max_row = max(ws_a.max_row or 0, ws_b.max_row or 0)
            max_col = max(ws_a.max_column or 0, ws_b.max_column or 0)
            for row in range(1, max_row + 1):
                for col in range(1, max_col + 1):
                    val_a = ws_a.cell(row=row, column=col).value
                    val_b = ws_b.cell(row=row, column=col).value
                    if val_a != val_b:
                        from openpyxl.utils import get_column_letter

                        cell_ref = f"{get_column_letter(col)}{row}"
                        differences.append(
                            {
                                "sheet": sn,
                                "cell": cell_ref,
                                "value_a": val_a,
                                "value_b": val_b,
                            }
                        )

        result: dict = {
            "differences": differences,
            "total_differences": len(differences),
            "sheets_compared": sheets_to_compare,
            "identical": len(differences) == 0,
        }

        if output_file and differences:
            diff_df = pd.DataFrame(differences)
            _write_df_to_new_file(diff_df, output_file, sheet_name="Differences")
            logger.info("Comparison differences written to %s", output_file)

        logger.info("Compared %d sheets; %d differences found", len(sheets_to_compare), len(differences))
        return result
    finally:
        wb_a.close()
        wb_b.close()
