from __future__ import annotations

from typing import TypedDict

from .common import ValidationOperator


class TableInfo(TypedDict):
    """An Excel table defined on a worksheet."""

    name: str
    ref: str
    style: str | None


class ValidationRuleInfo(TypedDict, total=False):
    """A data validation rule applied to a range."""

    range: str
    type: str
    operator: ValidationOperator
    formula1: str | None
    formula2: str | None
    allow_blank: bool
    show_dropdown: bool
