"""Pivot table, ETL, and data transformation operations."""

from __future__ import annotations

import ast
from logging import Logger

import pandas as pd

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
    validate_file_path,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

VALID_KEEP = {"first", "last", False}

_FORBIDDEN_NAMES: frozenset[str] = frozenset(
    {
        "__builtins__",
        "__import__",
        "exec",
        "eval",
        "open",
        "system",
        "getattr",
        "setattr",
        "delattr",
        "globals",
        "locals",
        "compile",
        "breakpoint",
        "input",
        "print",
        "exit",
        "quit",
    }
)

_ALLOWED_AST_NODES = (
    ast.Expression,
    ast.BinOp,
    ast.UnaryOp,
    ast.Constant,
    ast.Name,
    ast.Load,
    ast.Add,
    ast.Sub,
    ast.Mult,
    ast.Div,
    ast.Pow,
    ast.Mod,
    ast.FloorDiv,
    ast.USub,
    ast.UAdd,
)


def _validate_eval_expression(expression: str) -> None:
    """Validate expression using AST whitelist — only arithmetic and column names allowed."""
    try:
        tree = ast.parse(expression, mode="eval")
    except SyntaxError as e:
        raise ValueError(f"Invalid expression syntax: {e}") from e

    for node in ast.walk(tree):
        if not isinstance(node, _ALLOWED_AST_NODES):
            raise ValueError(
                f"Expression contains disallowed operation '{type(node).__name__}'. "
                "Only column references and basic arithmetic are allowed."
            )
        if isinstance(node, ast.Name) and node.id in _FORBIDDEN_NAMES:
            raise ValueError(
                f"Expression references forbidden name '{node.id}'. "
                "Only column references and basic arithmetic are allowed."
            )


def _read_sheet_df(file_path: str, sheet_name: str, has_header: bool = True) -> pd.DataFrame:
    from mcp_server.utils.excel_helpers import read_sheet_df

    return read_sheet_df(file_path, sheet_name, header_row=1 if has_header else 0)


def _write_df_to_sheet(wb, sheet_name: str, df: pd.DataFrame) -> None:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)

    # Write headers
    for c_idx, col_name in enumerate(df.columns, start=1):
        ws.cell(row=1, column=c_idx, value=col_name)

    # Write data
    for r_idx, row_data in enumerate(df.values.tolist(), start=2):
        for c_idx, val in enumerate(row_data, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)


def create_pivot_table(
    file_path: str,
    sheet_name: str,
    index_cols: list[str],
    value_cols: list[str],
    aggfunc: str | dict = "sum",
    output_sheet: str | None = None,
    output_file: str | None = None,
    include_margins: bool = False,
    column_field: str | None = None,
) -> dict:
    """Create a static pivot table using pandas and write to a sheet or file."""
    df = _read_sheet_df(file_path, sheet_name)

    all_required_cols = index_cols + value_cols + ([column_field] if column_field else [])
    for col in all_required_cols:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found. Available: {list(df.columns)}")

    pivot_kwargs: dict = {"margins": True, "margins_name": "Total"} if include_margins else {}
    if column_field:
        pivot_kwargs["columns"] = column_field
    pivot = pd.pivot_table(df, index=index_cols, values=value_cols, aggfunc=aggfunc, **pivot_kwargs)
    pivot = pivot.reset_index()

    # Flatten multi-level column headers produced by column_field
    if column_field:
        pivot.columns = [
            "_".join(str(c) for c in col).strip("_") if isinstance(col, tuple) else str(col) for col in pivot.columns
        ]

    if output_file:
        validate_file_path(output_file, must_exist=False)
        pivot.to_excel(output_file, index=False, engine="openpyxl")
        logger.info("Pivot table written to %s", output_file)
    elif output_sheet:
        wb = load_workbook_safe(file_path)
        _write_df_to_sheet(wb, output_sheet, pivot)
        save_workbook_safe(wb, file_path)
        logger.info("Pivot table written to sheet '%s' in %s", output_sheet, file_path)

    return {
        "data": pivot.to_dict(orient="records"),
        "index_columns": index_cols,
        "value_columns": value_cols,
        "column_field": column_field,
        "operation": aggfunc,
    }


def unpivot_data(
    file_path: str,
    sheet_name: str,
    id_vars: list[str],
    value_vars: list[str],
    var_name: str = "Variable",
    value_name: str = "Value",
) -> dict:
    """Unpivot (melt) data from wide to long format."""
    df = _read_sheet_df(file_path, sheet_name)

    for col in id_vars + value_vars:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found. Available: {list(df.columns)}")

    melted = pd.melt(df, id_vars=id_vars, value_vars=value_vars, var_name=var_name, value_name=value_name)
    return {
        "data": melted.to_dict(orient="records"),
        "row_count": len(melted),
    }


def merge_datasets(
    file_path: str,
    sheet1: str,
    sheet2: str,
    join_key: str | list[str],
    how: str = "left",
    output_sheet: str | None = None,
    suffixes: tuple[str, str] = ("_x", "_y"),
) -> dict:
    """Merge two sheets like a SQL join."""
    valid_how = {"left", "right", "inner", "outer"}
    if how not in valid_how:
        raise ValueError(f"Unsupported join type '{how}'. Allowed: {valid_how}")

    df1 = _read_sheet_df(file_path, sheet1)
    df2 = _read_sheet_df(file_path, sheet2)

    keys = [join_key] if isinstance(join_key, str) else join_key
    for k in keys:
        if k not in df1.columns:
            raise ValueError(f"Key '{k}' not found in sheet '{sheet1}'. Available: {list(df1.columns)}")
        if k not in df2.columns:
            raise ValueError(f"Key '{k}' not found in sheet '{sheet2}'. Available: {list(df2.columns)}")

    if df1[keys].duplicated().any() or df2[keys].duplicated().any():
        logger.warning("Merge keys are not unique in one or both datasets. Result may contain unexpected rows.")

    merged = pd.merge(df1, df2, on=keys, how=how, suffixes=suffixes)

    if output_sheet:
        wb = load_workbook_safe(file_path)
        _write_df_to_sheet(wb, output_sheet, merged)
        save_workbook_safe(wb, file_path)
        logger.info("Merged result written to sheet '%s'", output_sheet)

    return {
        "data": merged.to_dict(orient="records"),
        "row_count": len(merged),
    }


def add_computed_column(
    file_path: str,
    sheet_name: str,
    new_column_name: str,
    expression: str,
    has_header: bool = True,
) -> str:
    """Add a computed column using a pandas-eval expression like 'Revenue - Cost'."""
    df = _read_sheet_df(file_path, sheet_name, has_header)
    _validate_eval_expression(expression)

    try:
        df[new_column_name] = df.eval(expression, engine="python")
    except Exception as e:
        raise ValueError(f"Failed to evaluate expression '{expression}': {e}") from e

    wb = load_workbook_safe(file_path)
    ws = get_sheet(wb, sheet_name)

    # Write header for new column
    new_col_idx = len(df.columns)
    if has_header:
        ws.cell(row=1, column=new_col_idx, value=new_column_name)

    start_row = 2 if has_header else 1
    for r_idx, val in enumerate(df[new_column_name].tolist(), start=start_row):
        ws.cell(row=r_idx, column=new_col_idx, value=val)

    save_workbook_safe(wb, file_path)
    logger.info("Added column '%s' to %s!%s", new_column_name, sheet_name, file_path)
    return f"Added column '{new_column_name}' ({expression}) to '{sheet_name}' ({len(df)} rows)."


def deduplicate_data(
    file_path: str,
    sheet_name: str,
    columns: list[str] | None = None,
    keep: str = "first",
) -> str:
    """Remove duplicate rows from a sheet."""
    if keep not in ("first", "last", False):
        raise ValueError(f"Invalid keep value '{keep}'. Allowed: 'first', 'last', or False.")

    df = _read_sheet_df(file_path, sheet_name)

    if columns:
        for c in columns:
            if c not in df.columns:
                raise ValueError(f"Column '{c}' not found. Available: {list(df.columns)}")

    original_count = len(df)
    df = df.drop_duplicates(subset=columns, keep=keep if keep else False)
    removed = original_count - len(df)

    wb = load_workbook_safe(file_path)
    ws = get_sheet(wb, sheet_name)

    # Clear existing data below header
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.value = None

    for r_idx, row_data in enumerate(df.values.tolist(), start=2):
        for c_idx, val in enumerate(row_data, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)

    save_workbook_safe(wb, file_path)
    logger.info("Removed %d duplicates from %s!%s", removed, sheet_name, file_path)
    return f"Removed {removed} duplicate row(s) from '{sheet_name}'. {len(df)} rows remain."
