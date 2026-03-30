from __future__ import annotations

from typing import TypedDict


class NamedRangeInfo(TypedDict):
    """A named range defined in the workbook."""

    name: str
    destination: str
    scope: str


class FormulaInfo(TypedDict):
    """A formula found in a cell."""

    cell_ref: str
    formula: str


class FormulaErrorInfo(TypedDict):
    """A formula error found in a cell."""

    cell: str
    error: str
