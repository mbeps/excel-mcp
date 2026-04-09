from __future__ import annotations

from typing import Literal

from mcp.types import ToolAnnotations

import mcp_server.tools.cell_ops as _cell_ops
import mcp_server.tools.formulas as _formulas
from mcp_server.models.named_ranges import FormulaErrorInfo, FormulaInfo

__all__ = [
    "formula_write",
    "formula_audit",
]


def formula_write(
    action: Literal["set", "batch", "fill", "auto_sum"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    formula: str | None = None,
    is_array: bool = False,
    target_range: str | None = None,
    formulas: dict[str, str] | None = None,
    source_range: str | None = None,
) -> str | dict:
    """Set, batch-set, fill or auto-sum formulas in cells.

    Args:
        action: "set", "batch", "fill", or "auto_sum".
            - "set": requires `cell_ref` and `formula` (is_array optional, target_range optional).
            - "batch": requires `formulas` dict mapping cell_ref→formula.
            - "fill": requires `cell_ref` (source) and `target_range`.
            - "auto_sum": requires `cell_ref` (destination) and optional `source_range`.
        file_path, sheet_name: Workbook and sheet to modify.
        is_array: Whether the formula is an array formula.
        target_range, formulas, source_range: Operation-specific params.

    Returns:
        str | dict: Success message or batch result.

    Notes:
        - Writing formulas mutates the workbook. `fill` uses formula translation utilities; verify absolute/relative reference behaviour.
    """
    if action == "set":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='set'.")
        if formula is None:
            raise ValueError("formula is required for action='set'.")
        return _formulas.set_formula(file_path, sheet_name, cell_ref, formula, is_array, target_range)
    if action == "batch":
        if formulas is None:
            raise ValueError("formulas is required for action='batch'.")
        return _formulas.set_formulas_batch(file_path, sheet_name, formulas)
    if action == "fill":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='fill'.")
        if target_range is None:
            raise ValueError("target_range is required for action='fill'.")
        return _cell_ops.fill_formula(file_path, sheet_name, cell_ref, target_range)
    if action == "auto_sum":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='auto_sum'.")
        return _cell_ops.auto_sum(file_path, sheet_name, cell_ref, source_range)
    raise ValueError(f"Unknown action: {action}")


def formula_audit(
    action: Literal["value", "errors", "precedents", "dependents", "list"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    cell_range: str | None = None,
) -> dict | list[FormulaErrorInfo] | list[str] | list[FormulaInfo]:
    """Inspect formulas: get cached value, list errors, find precedents/dependents, or list all formulas.

    Args:
        action: "value", "errors", "precedents", "dependents", or "list".
        file_path, sheet_name: Workbook and sheet.
        cell_ref: Required for "value", "precedents", and "dependents".
        cell_range: Optional range filter for "errors".

    Returns:
        dict or list: Action-dependent payload (e.g. value, list of FormulaErrorInfo, list of cell refs).

    Notes:
        - Read-only.
    """
    if action == "value":
        if not cell_ref:
            raise ValueError("cell_ref is required for action='value'.")
        return _formulas.get_formula_value(file_path, sheet_name, cell_ref)
    if action == "errors":
        return _formulas.get_formula_errors(file_path, sheet_name, cell_range)
    if action == "precedents":
        if not cell_ref:
            raise ValueError("cell_ref is required for action='precedents'.")
        return _formulas.get_formula_precedents(file_path, sheet_name, cell_ref)
    if action == "dependents":
        if not cell_ref:
            raise ValueError("cell_ref is required for action='dependents'.")
        return _formulas.get_formula_dependents(file_path, sheet_name, cell_ref)
    if action == "list":
        return _formulas.list_formulas(file_path, sheet_name)
    raise ValueError(f"Unknown action: {action}")


def register(mcp) -> None:
    """Register tools on *mcp*."""
    mcp.tool()(formula_write)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(formula_audit)
