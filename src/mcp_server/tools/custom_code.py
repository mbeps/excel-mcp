"""Sandboxed custom Python/pandas code execution against Excel files."""

from __future__ import annotations

import ast
import os
import pickle
import subprocess
import sys
import tempfile
from logging import Logger
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from mcp_server.utils.excel_helpers import read_sheet_df, validate_file_path, write_dataframe_to_workbook
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
        "read_xml",
        "to_xml",
        "to_latex",
        "to_string",
        "to_markdown",
        "info",
        "DataSource",
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

_BLOCKED_STRING_ATTRS: frozenset[str] = frozenset({"format", "format_map"})

_SANDBOX_TIMEOUT_SECONDS = 15
_SANDBOX_MAX_MEMORY_BYTES = 1024 * 1024 * 1024  # 1 GiB

_SANDBOX_RUNNER = """
import os
import pickle
import sys

import numpy as np
import pandas as pd

try:
    import resource
    max_mem = int(os.environ.get("EXCEL_MCP_SANDBOX_MAX_MEMORY_BYTES", "1073741824"))
    try:
        resource.setrlimit(resource.RLIMIT_AS, (max_mem, max_mem))
    except (ValueError, OSError):
        pass
except ImportError:
    pass

SAFE_BUILTINS = {
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

input_pkl, output_pkl, code_file = sys.argv[1], sys.argv[2], sys.argv[3]
df = pd.read_pickle(input_pkl)
with open(code_file, "r", encoding="utf-8") as fh:
    code = fh.read()

orig_df = df.copy()
namespace = {"df": df, "pd": pd, "np": np, "result": None}
orig_id = id(namespace["df"])

try:
    exec(code, {"__builtins__": SAFE_BUILTINS}, namespace)
except Exception as e:
    with open(output_pkl, "wb") as fh:
        pickle.dump({"status": "error", "message": f"Execution error: {e}"}, fh)
    sys.exit(0)

user_result = namespace["result"]
modified_df = namespace["df"]
df_changed = (
    id(modified_df) != orig_id
    or not isinstance(modified_df, pd.DataFrame)
    or not orig_df.equals(modified_df)
)

payload = {
    "status": "ok",
    "user_result": user_result,
    "modified_df": modified_df if (user_result is None and df_changed) else None,
    "df_changed": bool(df_changed),
}
with open(output_pkl, "wb") as fh:
    pickle.dump(payload, fh)
"""


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
            if node.attr in _BLOCKED_MODULE_ATTRS:
                raise ValueError(f"Blocked: attribute access '{node.attr}' is not allowed")
            if isinstance(node.value, ast.Name) and node.value.id in _BLOCKED_MODULE_ATTRS:
                raise ValueError(f"Blocked: attribute access on '{node.value.id}' is not allowed")
            if node.attr in _BLOCKED_IO_ATTRS:
                raise ValueError(f"Blocked: I/O method '{node.attr}' is not allowed in sandbox")
            if node.attr in _BLOCKED_STRING_ATTRS:
                raise ValueError(f"Blocked: string method '{node.attr}' is not allowed in sandbox")

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
        df = read_sheet_df(str(validated_path), sheet_name)
    except Exception as e:
        return {"status": "error", "message": f"Failed to read file: {e}"}

    # 4. Execute in isolated subprocess sandbox
    try:
        with tempfile.TemporaryDirectory(prefix="excel_mcp_sandbox_") as tmp_dir:
            tmp_root = Path(tmp_dir)
            input_pkl = tmp_root / "input.pkl"
            output_pkl = tmp_root / "output.pkl"
            code_file = tmp_root / "user_code.py"

            df.to_pickle(input_pkl)
            code_file.write_text(code, encoding="utf-8")

            proc_env = {
                "PATH": os.environ.get("PATH", ""),
                "OPENBLAS_NUM_THREADS": "1",
                "OMP_NUM_THREADS": "1",
                "MKL_NUM_THREADS": "1",
                "EXCEL_MCP_SANDBOX_MAX_MEMORY_BYTES": os.environ.get(
                    "EXCEL_MCP_SANDBOX_MAX_MEMORY_BYTES",
                    str(_SANDBOX_MAX_MEMORY_BYTES),
                ),
            }
            proc = subprocess.run(
                [sys.executable, "-I", "-c", _SANDBOX_RUNNER, str(input_pkl), str(output_pkl), str(code_file)],
                env=proc_env,
                timeout=_SANDBOX_TIMEOUT_SECONDS,
                capture_output=True,
                text=True,
                check=False,
            )

            if not output_pkl.exists():
                err_detail = proc.stderr.strip() or f"sandbox exited with code {proc.returncode}"
                logger.warning("Custom code subprocess failed: %s", err_detail)
                return {"status": "error", "message": f"Execution error: {err_detail}"}

            with open(output_pkl, "rb") as fh:
                sandbox_out = pickle.load(fh)
    except subprocess.TimeoutExpired:
        logger.warning("Custom code execution timed out after %ss", _SANDBOX_TIMEOUT_SECONDS)
        return {"status": "error", "message": f"Execution timed out after {_SANDBOX_TIMEOUT_SECONDS} seconds"}
    except Exception as e:
        logger.warning("Custom code execution failed: %s", e)
        return {"status": "error", "message": f"Execution error: {e}"}

    if sandbox_out.get("status") != "ok":
        msg = sandbox_out.get("message", "Execution error")
        logger.warning("Custom code execution failed: %s", msg)
        return {"status": "error", "message": msg}

    # 5. Process result
    user_result = sandbox_out.get("user_result")
    modified_df = sandbox_out.get("modified_df")
    write_path = output_file or str(validated_path)

    def _write_df(target_df: pd.DataFrame, label: str) -> dict[str, Any]:
        dest = validate_file_path(write_path, must_exist=False)
        if dest.suffix.lower() == ".csv":
            target_df.to_csv(str(dest), index=False)
        else:
            write_dataframe_to_workbook(
                target_df,
                str(validated_path),
                sheet_name=sheet,
                output_file=output_file,
            )
        logger.info("Wrote %s to %s", label, dest)
        resp: dict[str, Any] = {
            "status": "success",
            "message": f"{label} ({target_df.shape[0]} rows, {target_df.shape[1]} cols) written to {dest.name}",
            "result": _serialize_result(target_df),
        }
        if proc.stdout:
            resp["stdout"] = proc.stdout
        return resp

    try:
        if user_result is not None:
            if isinstance(user_result, pd.DataFrame):
                return _write_df(user_result, "Result DataFrame")
            resp: dict[str, Any] = {
                "status": "success",
                "message": "Code executed successfully",
                "result": _serialize_result(user_result),
            }
            if proc.stdout:
                resp["stdout"] = proc.stdout
            return resp

        if isinstance(modified_df, pd.DataFrame):
            return _write_df(modified_df, "Modified DataFrame")

        resp = {
            "status": "success",
            "message": "Code executed successfully (no modifications to save)",
        }
        if proc.stdout:
            resp["stdout"] = proc.stdout
        return resp
    except Exception as e:
        logger.error("Failed to write result: %s", e)
        return {"status": "error", "message": f"Failed to write result: {e}"}
