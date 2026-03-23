"""Data validation operations for Excel worksheets."""

from __future__ import annotations

from logging import Logger
from typing import Any, cast

from openpyxl.worksheet.datavalidation import DataValidation

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

BETWEEN_OPERATORS = {"between", "notBetween"}


def add_dropdown_validation(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    options: list[str],
    allow_blank: bool = True,
) -> str:
    """Add a dropdown list data validation."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        formula = '"' + ",".join(options) + '"'
        dv = DataValidation(type="list", formula1=formula, allow_blank=allow_blank)
        dv.error = "Invalid entry. Please select from the dropdown list."
        dv.errorTitle = "Invalid Input"
        dv.prompt = "Select a value from the list."
        dv.promptTitle = "Dropdown"
        dv.add(cell_range)
        ws.add_data_validation(dv)

        save_workbook_safe(wb, file_path)
        logger.info("Added dropdown validation to %s in %s", cell_range, file_path)
        return f"Added dropdown validation to '{cell_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def add_numeric_validation(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    operator: str,
    value1: float,
    value2: float | None = None,
    allow_blank: bool = True,
) -> str:
    """Add numeric data validation."""
    if operator in BETWEEN_OPERATORS and value2 is None:
        raise ValueError(f"value2 is required for operator '{operator}'.")

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        val_type = "decimal" if isinstance(value1, float) and not value1.is_integer() else "whole"
        dv = DataValidation(
            type=val_type,
            operator=cast(Any, operator),
            formula1=str(value1),
            formula2=str(value2) if value2 is not None else None,
            allow_blank=allow_blank,
        )
        dv.error = f"Value must satisfy: {operator} {value1}" + (f" and {value2}" if value2 is not None else "")
        dv.errorTitle = "Invalid Number"
        dv.prompt = "Enter a valid number."
        dv.promptTitle = "Numeric Input"
        dv.add(cell_range)
        ws.add_data_validation(dv)

        save_workbook_safe(wb, file_path)
        logger.info("Added numeric validation (%s) to %s in %s", operator, cell_range, file_path)
        return f"Added numeric validation (operator='{operator}') to '{cell_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def add_formula_validation(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    formula: str,
    allow_blank: bool = True,
) -> str:
    """Add custom formula-based validation."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        dv = DataValidation(type="custom", formula1=formula, allow_blank=allow_blank)
        dv.error = "Value does not satisfy the custom formula."
        dv.errorTitle = "Validation Error"
        dv.prompt = "Enter a value that satisfies the formula."
        dv.promptTitle = "Custom Validation"
        dv.add(cell_range)
        ws.add_data_validation(dv)

        save_workbook_safe(wb, file_path)
        logger.info("Added formula validation to %s in %s", cell_range, file_path)
        return f"Added formula validation to '{cell_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def _resolve_range_values(ws: Any, formula: str) -> list | None:
    """Resolve a range reference formula to actual cell values."""
    from openpyxl.utils.cell import range_boundaries

    ref = formula.lstrip("=").strip()
    # Handle sheet-qualified references like Sheet1!$A$1:$A$5
    if "!" in ref:
        ref = ref.split("!", 1)[1]
    ref = ref.replace("$", "")

    try:
        min_col, min_row, max_col, max_row = range_boundaries(ref)
    except Exception:
        return None

    values = []
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            v = ws.cell(row=row, column=col).value
            if v is not None:
                values.append(v)
    return values


def list_validations(file_path: str, sheet_name: str, resolve_sources: bool = True) -> list[dict]:
    """List all data validations on a sheet, optionally resolving dropdown range sources."""
    wb = load_workbook_safe(file_path, read_only=False)
    try:
        ws = get_sheet(wb, sheet_name)
        result = []
        for dv in ws.data_validations.dataValidation:
            entry: dict[str, Any] = {
                "type": dv.type,
                "formula1": dv.formula1,
                "formula2": dv.formula2,
                "ranges": str(dv.sqref),
                "allow_blank": dv.allow_blank,
            }

            if (
                resolve_sources
                and dv.type == "list"
                and dv.formula1
                and (dv.formula1.startswith("=") or "$" in dv.formula1)
            ):
                resolved = _resolve_range_values(ws, dv.formula1)
                if resolved is not None:
                    entry["resolved_values"] = resolved

            result.append(entry)
        return result
    finally:
        wb.close()


def add_date_validation(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    operator: str = "greaterThan",
    date1: str = "",
    date2: str | None = None,
    allow_blank: bool = True,
) -> str:
    """Add date validation to a cell range."""
    if operator in BETWEEN_OPERATORS and not date2:
        raise ValueError(f"date2 is required for operator '{operator}'.")

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        dv = DataValidation(
            type="date",
            operator=cast(Any, operator),
            formula1=date1,
            formula2=date2,
            allow_blank=allow_blank,
        )
        dv.error = f"Date must satisfy: {operator} {date1}" + (f" and {date2}" if date2 else "")
        dv.errorTitle = "Invalid Date"
        dv.prompt = "Enter a valid date."
        dv.promptTitle = "Date Input"
        dv.add(cell_range)
        ws.add_data_validation(dv)

        save_workbook_safe(wb, file_path)
        logger.info("Added date validation (%s) to %s in %s", operator, cell_range, file_path)
        return f"Added date validation (operator='{operator}') to '{cell_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def add_text_length_validation(
    file_path: str,
    sheet_name: str,
    cell_range: str,
    operator: str = "lessThanOrEqual",
    length1: int = 100,
    length2: int | None = None,
    allow_blank: bool = True,
) -> str:
    """Add text length validation to a cell range."""
    if operator in BETWEEN_OPERATORS and length2 is None:
        raise ValueError(f"length2 is required for operator '{operator}'.")

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        dv = DataValidation(
            type="textLength",
            operator=cast(Any, operator),
            formula1=str(length1),
            formula2=str(length2) if length2 is not None else None,
            allow_blank=allow_blank,
        )
        msg = f"Text length must satisfy: {operator} {length1}"
        dv.error = msg + (f" and {length2}" if length2 is not None else "")
        dv.errorTitle = "Invalid Text Length"
        dv.prompt = "Enter text of valid length."
        dv.promptTitle = "Text Length"
        dv.add(cell_range)
        ws.add_data_validation(dv)

        save_workbook_safe(wb, file_path)
        logger.info("Added text length validation (%s) to %s in %s", operator, cell_range, file_path)
        return f"Added text length validation (operator='{operator}') to '{cell_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def remove_validation(file_path: str, sheet_name: str, cell_range: str) -> str:
    """Remove data validations from a specific range."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        to_remove = [dv for dv in ws.data_validations.dataValidation if str(dv.sqref) == cell_range]

        for dv in to_remove:
            ws.data_validations.dataValidation.remove(dv)

        save_workbook_safe(wb, file_path)
        count = len(to_remove)
        logger.info("Removed %d validation(s) from %s in %s", count, cell_range, file_path)
        return f"Removed {count} validation(s) from '{cell_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()
