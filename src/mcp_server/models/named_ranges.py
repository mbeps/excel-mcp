"""Schemas for named ranges and discovered formulas/errors in a workbook."""

from __future__ import annotations

from typing import TypedDict


class NamedRangeInfo(TypedDict):
    """A named range defined in the workbook.

    Keys:
        name (str): Name of the named range. Required.
        destination (str): A1-style destination (sheet!A1:B2). Required.
        scope (str): Workbook or sheet scope identifier.
    """

    name: str
    destination: str
    scope: str


class FormulaInfo(TypedDict):
    """A formula found in a cell.

    Keys:
        cell_ref (str): Cell reference containing the formula. Required.
        formula (str): The formula text (A1 style) as stored in the cell.
    """

    cell_ref: str
    formula: str


class FormulaErrorInfo(TypedDict):
    """A formula error found in a cell.

    Keys:
        cell (str): Cell reference with the error. Required.
        error (str): Error code or message (e.g., '#DIV/0!').
    """

    cell: str
    error: str
