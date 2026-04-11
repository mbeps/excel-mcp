"""Cross-file operations: bulk aggregate, bulk filter, and data consistency validation."""

from __future__ import annotations

from logging import Logger

import pandas as pd

from mcp_server.models.multi_file import (
    MultiFileFilterPerFileResult,
    MultiFilePerFileResult,
    WorkbookDiff,
)
from mcp_server.utils.excel_helpers import (
    load_workbook_safe,
    read_sheet_df,
    validate_file_path,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

AGGREGATE_OPS = {"sum", "mean", "min", "max", "count"}
FILTER_OPERATORS = {
    "equals",
    "not_equals",
    "greater_than",
    "less_than",
    "greater_than_or_equal",
    "less_than_or_equal",
    "contains",
}
OPERATOR_ALIASES: dict[str, str] = {
    "==": "equals",
    "!=": "not_equals",
    ">": "greater_than",
    "<": "less_than",
    ">=": "greater_than_or_equal",
    "<=": "less_than_or_equal",
}


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
) -> dict[str, str | float | int | list[MultiFilePerFileResult]]:
    """Aggregate the same column across multiple files."""
    if operation not in AGGREGATE_OPS:
        raise ValueError(f"Unsupported operation '{operation}'. Allowed: {AGGREGATE_OPS}")
    if not file_paths:
        raise ValueError("file_paths must not be empty.")

    per_file: list[MultiFilePerFileResult] = []
    all_values: list[float] = []

    for fp in file_paths:
        df = read_sheet_df(fp, sheet_name, header_row)
        if column not in df.columns:
            raise ValueError(f"Column '{column}' not found in '{fp}'. Available: {list(df.columns)}")

        col_data = pd.to_numeric(df[column], errors="coerce").dropna()
        values = col_data.tolist()
        all_values.extend(values)

        file_value: float
        if operation == "sum":
            file_value = float(col_data.sum())
        elif operation == "mean":
            file_value = float(col_data.mean()) if len(col_data) > 0 else 0.0
        elif operation == "min":
            file_value = float(col_data.min()) if len(col_data) > 0 else 0.0
        elif operation == "max":
            file_value = float(col_data.max()) if len(col_data) > 0 else 0.0
        else:  # count
            file_value = float(len(col_data))

        per_file.append(MultiFilePerFileResult(file=fp, row_count=len(df), value=file_value))

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

    result: dict[str, str | float | int | list[MultiFilePerFileResult]] = {
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
) -> dict[str, str | float | int | list[MultiFileFilterPerFileResult]]:
    """Filter rows across multiple files based on a condition."""
    operator = OPERATOR_ALIASES.get(operator, operator)
    if operator not in FILTER_OPERATORS:
        raise ValueError(f"Unsupported operator '{operator}'. Allowed: {FILTER_OPERATORS}")
    if not file_paths:
        raise ValueError("file_paths must not be empty.")

    per_file: list[dict] = []
    matched_frames: list[pd.DataFrame] = []

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
        elif operator == "greater_than_or_equal":
            mask = pd.to_numeric(col, errors="coerce") >= float(value)
        elif operator == "less_than_or_equal":
            mask = pd.to_numeric(col, errors="coerce") <= float(value)
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
            matched_frames.append(tagged)

    total_matched = sum(f["matched_rows"] for f in per_file)

    result = {
        "column": column,
        "operator": operator,
        "value": value,
        "per_file": per_file,
        "total_matched": total_matched,
        "total_files": len(file_paths),
    }

    if output_file and matched_frames:
        all_matched = pd.concat(matched_frames, ignore_index=True)
        _write_df_to_new_file(all_matched, output_file, sheet_name="FilteredResults")
        logger.info("Filtered results written to %s", output_file)

    return result


def validate_data_consistency(
    file_paths: list[str],
    key_column: str,
    check_columns: list[str] | None = None,
    sheet_name: str = "Sheet1",
    header_row: int = 1,
) -> dict[str, bool | int | dict[str, list[str]] | list[dict[str, str | dict[str, str]]]]:
    """Cross-file schema and referential integrity check.

    Performs two layers of validation:
    1. **Schema validation** — compares column names (and inferred dtypes) across files.
       When ``check_columns`` is provided, verifies those columns exist in every file.
       When omitted, verifies all files share the same column set.
    2. **Referential integrity** — checks that every unique key value in ``key_column``
       appears in all files, and that ``check_columns`` values agree across files for
       matching keys.
    """
    if not file_paths:
        raise ValueError("file_paths must not be empty.")

    # --- Phase 1: Schema validation ---
    file_data: list[tuple[str, pd.DataFrame]] = []
    file_schemas: dict[str, list[str]] = {}
    schema_mismatches: list[dict] = []

    for fp in file_paths:
        df = read_sheet_df(fp, sheet_name, header_row)
        file_data.append((fp, df))
        file_schemas[fp] = list(df.columns)

    # Check that key_column exists in every file
    for fp, df in file_data:
        if key_column not in df.columns:
            schema_mismatches.append(
                {
                    "file": fp,
                    "issue": "missing_key_column",
                    "columns": [key_column],
                    "available": list(df.columns),
                }
            )

    # Check that required columns exist in every file
    if check_columns:
        for fp, df in file_data:
            missing_cols = [c for c in check_columns if c not in df.columns]
            if missing_cols:
                schema_mismatches.append(
                    {
                        "file": fp,
                        "issue": "missing_columns",
                        "columns": missing_cols,
                        "available": list(df.columns),
                    }
                )
    else:
        # Compare all files against the union of columns
        all_columns: set[str] = set()
        for cols in file_schemas.values():
            all_columns.update(cols)
        for fp, cols in file_schemas.items():
            missing_cols = sorted(all_columns - set(cols))
            if missing_cols:
                schema_mismatches.append(
                    {
                        "file": fp,
                        "issue": "missing_columns",
                        "columns": missing_cols,
                        "available": cols,
                    }
                )

    # --- Phase 2: Referential integrity (only if key_column exists in all files) ---
    all_keys: set = set()
    key_valid_files = [(fp, df) for fp, df in file_data if key_column in df.columns]

    if len(key_valid_files) < len(file_data):
        # key_column missing from some files — already captured in schema_mismatches
        missing_keys: dict[str, list] = {}
    else:
        for fp, df in key_valid_files:
            all_keys.update(df[key_column].dropna().unique().tolist())

        missing_keys = {}
        for fp, df in key_valid_files:
            file_keys = set(df[key_column].dropna().unique().tolist())
            missing = sorted(str(k) for k in all_keys - file_keys)
            if missing:
                missing_keys[fp] = missing

    # Check value consistency across files for matching keys
    mismatched_values: list[dict] = []
    effective_check_cols = check_columns or []
    if effective_check_cols and key_valid_files:
        key_values: dict[str, dict[str, dict[str, object]]] = {}
        for fp, df in key_valid_files:
            for _, row in df.iterrows():
                key = row[key_column]
                if pd.isna(key):
                    continue
                key_str = str(key)
                if key_str not in key_values:
                    key_values[key_str] = {}
                vals = {}
                for cc in effective_check_cols:
                    if cc in df.columns:
                        v = row[cc]
                        vals[cc] = v if not pd.isna(v) else None
                key_values[key_str][fp] = vals

        for key_str, file_vals in key_values.items():
            if len(file_vals) < 2:
                continue
            files_list = list(file_vals.keys())
            for cc in effective_check_cols:
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

    consistent = len(missing_keys) == 0 and len(mismatched_values) == 0 and len(schema_mismatches) == 0

    return {
        "consistent": consistent,
        "total_keys": len(all_keys),
        "total_files": len(file_paths),
        "schema_mismatches": schema_mismatches,
        "missing_keys": missing_keys,
        "mismatched_values": mismatched_values,
    }


def compare_workbooks(
    file_path_a: str,
    file_path_b: str,
    sheet_name: str | None = None,
    output_file: str | None = None,
    sheet_name_b: str | None = None,
) -> dict[str, list[WorkbookDiff] | int | list[str] | bool]:
    """Compare two workbooks cell-by-cell.

    If both sheet_name and sheet_name_b are given, compare that single pair
    (sheet_name from file A vs sheet_name_b from file B).
    If only sheet_name is given, compare that sheet in both workbooks (must exist in both).
    If neither is given, auto-detect common sheet names.
    """
    wb_a = load_workbook_safe(file_path_a, data_only=True)
    wb_b = load_workbook_safe(file_path_b, data_only=True)
    try:
        if sheet_name and sheet_name_b:
            if sheet_name not in wb_a.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' not found in workbook A.")
            if sheet_name_b not in wb_b.sheetnames:
                raise ValueError(f"Sheet '{sheet_name_b}' not found in workbook B.")
            sheet_pairs = [(sheet_name, sheet_name_b)]
        elif sheet_name:
            if sheet_name not in wb_a.sheetnames or sheet_name not in wb_b.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' not found in both workbooks.")
            sheet_pairs = [(sheet_name, sheet_name)]
        else:
            sheet_pairs = [(s, s) for s in wb_a.sheetnames if s in wb_b.sheetnames]

        differences: list[dict] = []
        sheets_compared: list[str] = []
        for sn_a, sn_b in sheet_pairs:
            ws_a = wb_a[sn_a]
            ws_b = wb_b[sn_b]
            label = sn_a if sn_a == sn_b else f"{sn_a} vs {sn_b}"
            sheets_compared.append(label)
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
                                "sheet": label,
                                "cell": cell_ref,
                                "value_a": val_a,
                                "value_b": val_b,
                            }
                        )

        result: dict = {
            "differences": differences,
            "total_differences": len(differences),
            "sheets_compared": sheets_compared,
            "identical": len(differences) == 0,
        }

        if output_file and differences:
            diff_df = pd.DataFrame(differences)
            _write_df_to_new_file(diff_df, output_file, sheet_name="Differences")
            logger.info("Comparison differences written to %s", output_file)

        logger.info("Compared %d sheets; %d differences found", len(sheets_compared), len(differences))
        return result
    finally:
        wb_a.close()
        wb_b.close()
