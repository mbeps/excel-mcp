"""Sandboxed custom Python/pandas code execution against Excel files."""

from __future__ import annotations

import ast
from logging import Logger
from typing import Any

import numpy as np
import pandas as pd

from mcp_server.utils.excel_helpers import validate_file_path
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

_BLOCKED_FUNC_NAMES: frozenset[str] = frozenset(
    {
        "exec",
        "eval",
        "compile",
        "open",
        "__import__",
        "globals",
        "locals",
        "getattr",
        "setattr",
        "delattr",
        "vars",
        "dir",
        "breakpoint",
    }
)

_BLOCKED_MODULE_ATTRS: frozenset[str] = frozenset(
    {
        "os",
        "sys",
        "subprocess",
        "shutil",
        "pathlib",
    }
)

_BLOCKED_IO_ATTRS: frozenset[str] = frozenset(
    {
        # pandas I/O
        "read_csv",
        "to_csv",
        "read_excel",
        "to_excel",
        "read_parquet",
        "to_parquet",
        "read_sql",
        "to_sql",
        "read_json",
        "to_json",
        "read_html",
        "to_html",
        "read_pickle",
        "to_pickle",
        "read_clipboard",
        "to_clipboard",
        "read_feather",
        "to_feather",
        "read_hdf",
        "to_hdf",
        "read_orc",
        "to_orc",
        "read_sas",
        "read_spss",
        "read_stata",
        "to_stata",
        "read_gbq",
        "to_gbq",
        "read_fwf",
        "read_table",
        "ExcelFile",
        "ExcelWriter",
        "HDFStore",
        # numpy I/O
        "load",
        "save",
        "savez",
        "savetxt",
        "loadtxt",
        "genfromtxt",
        "fromfile",
        "tofile",
    }
)

_SAFE_BUILTINS: dict[str, Any] = {
    "len": len,
    "range": range,
    "enumerate": enumerate,
    "zip": zip,
    "map": map,
    "filter": filter,
    "sorted": sorted,
    "reversed": reversed,
    "min": min,
    "max": max,
    "sum": sum,
    "abs": abs,
    "round": round,
    "int": int,
    "float": float,
    "str": str,
    "bool": bool,
    "list": list,
    "dict": dict,
    "tuple": tuple,
    "set": set,
    "isinstance": isinstance,
    "type": type,
    "print": print,
    "any": any,
    "all": all,
    "ValueError": ValueError,
    "TypeError": TypeError,
    "KeyError": KeyError,
    "IndexError": IndexError,
}


def _validate_code(code: str) -> None:
    """Parse and validate user code via AST analysis. Raises ValueError on unsafe patterns."""
    try:
        tree = ast.parse(code, mode="exec")
    except SyntaxError as e:
        raise ValueError(f"Syntax error in code: {e}") from e

    for node in ast.walk(tree):
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            raise ValueError("Blocked: import statements are not allowed")

        if isinstance(node, ast.Attribute):
            if isinstance(node.attr, str) and node.attr.startswith("__") and node.attr.endswith("__"):
                raise ValueError(f"Blocked: dunder attribute access '{node.attr}' is not allowed")
            if isinstance(node.value, ast.Name) and node.value.id in _BLOCKED_MODULE_ATTRS:
                raise ValueError(f"Blocked: attribute access on '{node.value.id}' is not allowed")
            if node.attr in _BLOCKED_IO_ATTRS:
                raise ValueError(f"Blocked: I/O method '{node.attr}' is not allowed in sandbox")

        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            if node.func.id in _BLOCKED_FUNC_NAMES:
                raise ValueError(f"Blocked: calling '{node.func.id}()' is not allowed")


def _serialize_result(value: Any) -> Any:
    """Convert a result value to a JSON-serializable form."""
    if isinstance(value, pd.DataFrame):
        if len(value) > 50:
            return {
                "type": "DataFrame",
                "shape": list(value.shape),
                "columns": list(value.columns),
                "preview": value.head(10).to_dict(orient="records"),
            }
        return {
            "type": "DataFrame",
            "shape": list(value.shape),
            "columns": list(value.columns),
            "data": value.to_dict(orient="records"),
        }
    if isinstance(value, pd.Series):
        return {"type": "Series", "name": value.name, "data": value.tolist()}
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating,)):
        return float(value)
    if isinstance(value, np.ndarray):
        return value.tolist()
    if isinstance(value, (dict, list, str, int, float, bool)) or value is None:
        return value
    return str(value)


def execute_custom_code(
    file_path: str,
    code: str,
    sheet: str | None = None,
    output_file: str | None = None,
) -> dict[str, Any]:
    """Execute custom Python/pandas code against an Excel file in a sandboxed environment.

    The code runs with access to ``df`` (the sheet data), ``pd`` (pandas),
    ``np`` (numpy), and a ``result`` variable.  Set ``result`` in your code to
    control the return value.  If ``result`` is a DataFrame it will be written
    back; if it is another serialisable type it is returned directly.

    Args:
        file_path: Path to the Excel/CSV file.
        code: Python code to execute.
        sheet: Optional sheet name (defaults to first sheet).
        output_file: Optional path for writing results. Defaults to overwriting
            the input file when the result is a DataFrame.

    Returns:
        A dict with ``status``, ``message``, and optionally ``result``.
    """
    # 1. Validate file path
    try:
        validated_path = validate_file_path(file_path)
    except ValueError as e:
        return {"status": "error", "message": str(e)}

    # 2. Validate code safety
    try:
        _validate_code(code)
    except ValueError as e:
        return {"status": "error", "message": str(e)}

    # 3. Read data
    try:
        sheet_name: str | int = sheet if sheet is not None else 0
        df = pd.read_excel(validated_path, sheet_name=sheet_name, engine="calamine")
    except Exception as e:
        return {"status": "error", "message": f"Failed to read file: {e}"}

    # 4. Execute in restricted namespace
    namespace: dict[str, Any] = {
        "df": df,
        "pd": pd,
        "np": np,
        "result": None,
    }

    original_id = id(namespace["df"])
    try:
        exec(code, {"__builtins__": _SAFE_BUILTINS}, namespace)  # noqa: S102
    except Exception as e:
        logger.warning("Custom code execution failed: %s", e)
        return {"status": "error", "message": f"Execution error: {e}"}

    # 5. Process result
    user_result = namespace["result"]
    modified_df = namespace["df"]

    write_path = output_file or str(validated_path)

    try:
        if user_result is not None:
            if isinstance(user_result, pd.DataFrame):
                dest = validate_file_path(write_path, must_exist=False)
                user_result.to_excel(str(dest), index=False, engine="openpyxl")
                logger.info("Wrote DataFrame result to %s", dest)
                return {
                    "status": "success",
                    "message": (
                        f"Result DataFrame ({user_result.shape[0]} rows, "
                        f"{user_result.shape[1]} cols) written to {dest.name}"
                    ),
                    "result": _serialize_result(user_result),
                }
            serialized = _serialize_result(user_result)
            return {
                "status": "success",
                "message": "Code executed successfully",
                "result": serialized,
            }

        # No explicit result — check if df was modified (identity or content)
        df_changed = (
            id(modified_df) != original_id or not isinstance(modified_df, pd.DataFrame) or not df.equals(modified_df)
        )
        if isinstance(modified_df, pd.DataFrame) and df_changed:
            dest = validate_file_path(write_path, must_exist=False)
            modified_df.to_excel(str(dest), index=False, engine="openpyxl")
            logger.info("Wrote modified df to %s", dest)
            return {
                "status": "success",
                "message": (
                    f"Modified DataFrame ({modified_df.shape[0]} rows, "
                    f"{modified_df.shape[1]} cols) written to {dest.name}"
                ),
                "result": _serialize_result(modified_df),
            }

        return {
            "status": "success",
            "message": "Code executed successfully (no modifications to save)",
        }
    except Exception as e:
        logger.error("Failed to write result: %s", e)
        return {"status": "error", "message": f"Failed to write result: {e}"}
