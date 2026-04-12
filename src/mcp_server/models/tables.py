"""Schemas for table metadata and data validation rule descriptions."""

from __future__ import annotations

from typing import TypedDict

from .common import ValidationOperator


class TableInfo(TypedDict):
    """An Excel table defined on a worksheet.

    Keys:
        name (str): Table name. Required.
        ref (str): A1-style reference for the table. Required.
        style (str | None): Named table style, if set.
    """

    name: str
    ref: str
    style: str | None


class ValidationRuleInfo(TypedDict, total=False):
    """A data validation rule applied to a range.

    Keys (total=False):
        range (str): A1-style range string.
        type (str): Validation type (e.g., 'list', 'whole', 'decimal', 'date').
        operator (ValidationOperator): Operator used (see common.ValidationOperator).
        formula1, formula2 (str | None): Formula bounds.
        allow_blank (bool): Allow blank values when True.
        show_dropdown (bool): Show dropdown for list validations.
    """

    range: str
    type: str
    operator: ValidationOperator
    formula1: str | None
    formula2: str | None
    allow_blank: bool
    show_dropdown: bool
