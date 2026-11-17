"""Formula operations: set, validate, and batch-apply Excel formulas."""

from __future__ import annotations

from logging import Logger

from openpyxl.worksheet.formula import ArrayFormula

from mcp_server.models.named_ranges import FormulaErrorInfo, FormulaInfo
from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def set_formula(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
    formula: str,
    is_array: bool = False,
    target_range: str | None = None,
) -> str:
    """Set an Excel formula on a cell. Prepends '=' if missing.

    Set is_array=True to apply as an array (CSE) formula. For array formulas,
    target_range specifies the full range (e.g. 'A1:A10'); cell_ref is the anchor cell.
    Strips curly braces from array formulas if present.
    """
    formula = formula.strip("{}")
    if not formula.startswith("="):
        formula = f"={formula}"

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        if is_array:
            range_ref = target_range or cell_ref
            ws[cell_ref] = ArrayFormula(range_ref, formula)
            save_workbook_safe(wb, file_path)
            logger.info("Set array formula on %s!%s: %s", sheet_name, range_ref, formula)
            return f"Array formula {formula} applied to {range_ref} in '{sheet_name}'."
        else:
            ws[cell_ref] = formula
            save_workbook_safe(wb, file_path)
            logger.info("Set formula on %s!%s: %s", sheet_name, cell_ref, formula)
            return f"Formula {formula} set on {cell_ref} in '{sheet_name}'."
    finally:
        wb.close()


def set_formulas_batch(file_path: str, sheet_name: str, formulas: dict[str, str]) -> str:
    """Set multiple formulas at once. Keys are cell refs, values are formulas."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        for cell_ref, formula in formulas.items():
            formula = formula.strip("{}")
            if not formula.startswith("="):
                formula = f"={formula}"
            ws[cell_ref] = formula

        save_workbook_safe(wb, file_path)
        count = len(formulas)
        logger.info("Set %d formulas in %s!%s", count, file_path, sheet_name)
        return f"Set {count} formulas in '{sheet_name}'."
    finally:
        wb.close()


def get_formula_value(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
) -> dict[str, object]:
    """Read the cached calculated value of a formula cell (loads with data_only=True)."""
    wb = load_workbook_safe(file_path, data_only=True)
    try:
        ws = get_sheet(wb, sheet_name)
        val = getattr(ws[cell_ref].value, "text", ws[cell_ref].value)
        result: dict[str, object] = {"cell": cell_ref, "value": val, "type": type(val).__name__}
        if val is None:
            # Check whether the cell actually contains a formula (without data_only)
            wb2 = load_workbook_safe(file_path)
            try:
                raw_formula = get_sheet(wb2, sheet_name)[cell_ref].value
                formula = getattr(raw_formula, "text", raw_formula)
            finally:
                wb2.close()
            if isinstance(formula, str) and formula.startswith("="):
                result["formula"] = formula
                result["note"] = (
                    "openpyxl cannot evaluate formulas. The cached value is null because "
                    "this file was not saved by Excel. To get computed values, open the "
                    "file in Excel and save it."
                )
        return result
    finally:
        wb.close()


def get_formula_errors(
    file_path: str,
    sheet_name: str,
    cell_range: str | None = None,
) -> dict[str, list[FormulaErrorInfo] | int]:
    """Find all cells with Excel error values (#REF!, #VALUE!, #DIV/0!, etc.)."""
    EXCEL_ERRORS = {
        "#REF!",
        "#VALUE!",
        "#DIV/0!",
        "#N/A",
        "#NAME?",
        "#NUM!",
        "#NULL!",
        "#SPILL!",
        "#CALC!",
        "#GETTING_DATA",
        "#BLOCKED!",
        "#CONNECT!",
        "#FIELD!",
    }

    wb = load_workbook_safe(file_path, read_only=True, data_only=True)
    try:
        ws = get_sheet(wb, sheet_name)

        if cell_range:
            raw = ws[cell_range]
            if not isinstance(raw, tuple):
                rows_iter = ((raw,),)
            elif raw and not isinstance(raw[0], tuple):
                rows_iter = (raw,)
            else:
                rows_iter = raw
        else:
            rows_iter = ws.iter_rows()

        errors: list[FormulaErrorInfo] = []
        for row in rows_iter:
            for cell in row:
                if isinstance(cell.value, str) and cell.value in EXCEL_ERRORS:
                    errors.append({"cell": cell.coordinate, "error": cell.value})

        return {"errors": errors, "count": len(errors)}
    finally:
        wb.close()


_MAX_PRECEDENT_EXPANSION = 100_000
_STRING_LITERAL_RE = r'"(?:[^"]|"")*"'
_SHEET_PREFIX_RE = r"(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!"


def _strip_string_literals(formula: str) -> str:
    import re

    return re.sub(_STRING_LITERAL_RE, "", formula)


def _unquote_sheet_prefix(prefix_with_bang: str) -> str:
    name = prefix_with_bang[:-1]
    if len(name) >= 2 and name.startswith("'") and name.endswith("'"):
        return name[1:-1].replace("''", "'")
    return name


def get_formula_precedents(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
) -> dict[str, object]:
    """Parse a cell's formula to extract cell references it depends on (precedents)."""
    import re

    wb = load_workbook_safe(file_path, read_only=True)
    try:
        ws = get_sheet(wb, sheet_name)
        formula = getattr(ws[cell_ref].value, "text", ws[cell_ref].value)

        if not formula or not isinstance(formula, str) or not formula.startswith("="):
            return {"cell": cell_ref, "formula": formula, "precedents": []}

        formula_no_str = _strip_string_literals(formula)
        range_pattern = (
            rf"(?<![A-Za-z0-9_])(?:{_SHEET_PREFIX_RE})?"
            r"\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?(?![A-Za-z0-9_(])"
        )
        raw_refs = re.findall(range_pattern, formula_no_str, flags=re.IGNORECASE)

        from openpyxl.utils import column_index_from_string, get_column_letter
        from openpyxl.utils.cell import coordinate_from_string

        precedents: list[str] = []
        for ref in raw_refs:
            if ":" in ref:
                prefix = ""
                range_part = ref
                if "!" in ref:
                    prefix, range_part = ref.rsplit("!", 1)
                    prefix += "!"
                start, end = range_part.split(":")
                start_col_str, start_row = coordinate_from_string(start.replace("$", "").upper())
                end_col_str, end_row = coordinate_from_string(end.replace("$", "").upper())
                start_col = column_index_from_string(start_col_str)
                end_col = column_index_from_string(end_col_str)
                c_min, c_max = min(start_col, end_col), max(start_col, end_col)
                r_min, r_max = min(start_row, end_row), max(start_row, end_row)
                if (c_max - c_min + 1) * (r_max - r_min + 1) > _MAX_PRECEDENT_EXPANSION:
                    raise ValueError(
                        f"Range '{ref}' exceeds maximum precedent expansion of {_MAX_PRECEDENT_EXPANSION} cells."
                    )
                for c in range(c_min, c_max + 1):
                    for r in range(r_min, r_max + 1):
                        precedents.append(f"{prefix}{get_column_letter(c)}{r}")
            else:
                precedents.append(ref)

        return {"cell": cell_ref, "formula": formula, "precedents": list(dict.fromkeys(precedents))}
    finally:
        wb.close()


def _cell_in_range(cell_ref: str, range_ref: str) -> bool:
    """Check if a cell falls within an A1-style range (e.g. 'B2' in 'B1:B10')."""
    from openpyxl.utils import column_index_from_string
    from openpyxl.utils.cell import coordinate_from_string

    parts = range_ref.split(":")
    if len(parts) != 2:
        return False

    col_str, row = coordinate_from_string(cell_ref)
    col = column_index_from_string(col_str)

    start_col_str, start_row = coordinate_from_string(parts[0])
    start_col = column_index_from_string(start_col_str)

    end_col_str, end_row = coordinate_from_string(parts[1])
    end_col = column_index_from_string(end_col_str)

    return bool(
        min(start_col, end_col) <= col <= max(start_col, end_col)
        and min(start_row, end_row) <= row <= max(start_row, end_row)
    )


def get_formula_dependents(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
) -> dict[str, object]:
    """Find all cells that reference the given cell in their formula (dependents)."""
    import re

    wb = load_workbook_safe(file_path, read_only=True)
    try:
        ws = get_sheet(wb, sheet_name)
        cell_upper = cell_ref.upper().replace("$", "")
        target_sheet_lower = sheet_name.lower()

        cross_sheet_re = re.compile(
            rf"({_SHEET_PREFIX_RE})(\$?[A-Za-z]{{1,3}}\$?\d+(?::\$?[A-Za-z]{{1,3}}\$?\d+)?)",
            flags=re.IGNORECASE,
        )
        range_pattern = re.compile(r"(?<![A-Z0-9_])[A-Z]{1,3}\d+:[A-Z]{1,3}\d+(?![A-Z0-9_(])")

        def _filter_sheet_ref(match: re.Match[str]) -> str:
            ref_sheet = _unquote_sheet_prefix(match.group(1)).lower()
            return match.group(2) if ref_sheet == target_sheet_lower else ""

        dependents: list[dict[str, str]] = []
        for row in ws.iter_rows():
            for cell in row:
                raw_val = getattr(cell.value, "text", cell.value)
                if not isinstance(raw_val, str) or not raw_val.startswith("="):
                    continue
                no_str = _strip_string_literals(raw_val)
                filtered = cross_sheet_re.sub(_filter_sheet_ref, no_str)
                formula_clean = filtered.upper().replace("$", "")
                if re.search(r"(?<![A-Z0-9_])" + re.escape(cell_upper) + r"(?![A-Z0-9_(])", formula_clean):
                    dependents.append({"cell": cell.coordinate, "formula": raw_val})
                    continue
                for rng in range_pattern.findall(formula_clean):
                    if _cell_in_range(cell_upper, rng):
                        dependents.append({"cell": cell.coordinate, "formula": raw_val})
                        break

        return {"cell": cell_ref, "dependents": dependents, "count": len(dependents)}
    finally:
        wb.close()


def list_formulas(file_path: str, sheet_name: str) -> list[FormulaInfo]:
    """List all cells containing formulas in a sheet.

    Returns list of {cell_ref: str, formula: str}.
    """
    wb = load_workbook_safe(file_path, read_only=True)
    try:
        ws = get_sheet(wb, sheet_name)
        results: list[FormulaInfo] = []
        for row in ws.iter_rows():
            for cell in row:
                raw_val = getattr(cell.value, "text", cell.value)
                if isinstance(raw_val, str) and raw_val.startswith("="):
                    results.append({"cell_ref": cell.coordinate, "formula": raw_val})
        logger.info("Found %d formulas in %s!%s", len(results), sheet_name, file_path)
        return results
    finally:
        wb.close()
