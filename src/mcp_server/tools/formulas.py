"""Formula operations: set, validate, and batch-apply Excel formulas."""

from __future__ import annotations

from logging import Logger

from openpyxl.worksheet.formula import ArrayFormula

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def set_formula(file_path: str, sheet_name: str, cell_ref: str, formula: str) -> str:
    """Set an Excel formula on a cell. Prepends '=' if missing."""
    if not formula.startswith("="):
        formula = f"={formula}"

    wb = load_workbook_safe(file_path)
    ws = get_sheet(wb, sheet_name)
    ws[cell_ref] = formula
    save_workbook_safe(wb, file_path)
    logger.info("Set formula on %s!%s: %s", sheet_name, cell_ref, formula)
    return f"Formula {formula} set on {cell_ref} in '{sheet_name}'."


def set_array_formula(file_path: str, sheet_name: str, target_range: str, formula: str) -> str:
    """Apply an array formula to a range. Strips curly braces if present."""
    formula = formula.strip("{}")
    if not formula.startswith("="):
        formula = f"={formula}"

    wb = load_workbook_safe(file_path)
    ws = get_sheet(wb, sheet_name)

    anchor = target_range.split(":")[0]
    ws[anchor] = ArrayFormula(target_range, formula)

    save_workbook_safe(wb, file_path)
    logger.info("Set array formula on %s!%s: %s", sheet_name, target_range, formula)
    return f"Array formula {formula} applied to {target_range} in '{sheet_name}'."


def validate_formula_syntax(formula: str) -> dict:
    """Validate formula syntax and return token info."""
    from openpyxl.formula import Tokenizer

    if not formula.startswith("="):
        formula = f"={formula}"

    try:
        tok = Tokenizer(formula)
        tokens = [{"value": t.value, "type": t.type, "subtype": t.subtype} for t in tok.items]
        return {"valid": True, "tokens": tokens, "error": None}
    except Exception as exc:
        return {"valid": False, "tokens": [], "error": str(exc)}


def set_formulas_batch(file_path: str, sheet_name: str, formulas: dict[str, str]) -> str:
    """Set multiple formulas at once. Keys are cell refs, values are formulas."""
    wb = load_workbook_safe(file_path)
    ws = get_sheet(wb, sheet_name)

    for cell_ref, formula in formulas.items():
        if not formula.startswith("="):
            formula = f"={formula}"
        ws[cell_ref] = formula

    save_workbook_safe(wb, file_path)
    count = len(formulas)
    logger.info("Set %d formulas in %s!%s", count, file_path, sheet_name)
    return f"Set {count} formulas in '{sheet_name}'."
