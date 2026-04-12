"""Shared type aliases and Literal constraints used across model schemas.

This module centralises primitive union types and enumerated Literal values so model
definitions remain consistent across the codebase.

Attributes:
    CellScalar (TypeAlias): Union type for cell scalar values (str|int|float|bool|datetime|date|None).
    ScenarioCellValue (TypeAlias): Narrow scalar used by scenario storage (float|str|int|bool|None).
    BorderStyle, BorderSide, HorizontalAlignment, VerticalAlignment, ValidationOperator:
        Literal types enumerating allowed string values for styles, alignments and validation ops.
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Literal, TypeAlias

"""Type alias for a spreadsheet cell scalar value.

Attributes:
    CellScalar (TypeAlias): Allowed types: str | int | float | bool | datetime | date | None.

Purpose:
    Used across response and request models wherever a generic cell value is represented.

Constraints:
    - None represents an empty cell. Datetime/date objects are accepted as native Python types.
"""
CellScalar: TypeAlias = str | int | float | bool | datetime | date | None


"""Type alias for values stored in scenario definitions.

Attributes:
    ScenarioCellValue (TypeAlias): Allowed: float | str | int | bool | None.

Notes:
    - Designed to be slightly narrower than CellScalar: datetime values are not used by scenarios.
"""
ScenarioCellValue: TypeAlias = float | str | int | bool | None

# Border styles supported by openpyxl
"""Allowed border style names (maps to openpyxl's border style strings).

Constraints:
    - Must be one of the listed literal values. Used by formatting models.
"""
BorderStyle = Literal[
    "dashDot",
    "dashDotDot",
    "dashed",
    "dotted",
    "double",
    "hair",
    "medium",
    "mediumDashDot",
    "mediumDashDotDot",
    "mediumDashed",
    "slantDashDot",
    "thick",
    "thin",
]

# Side of a cell where borders can be applied
"""Sides of a cell where a border can be applied.

Keys:
    BorderSide (Literal): One of 'left', 'right', 'top', 'bottom', 'diagonal', 'outline'.
"""
BorderSide = Literal["left", "right", "top", "bottom", "diagonal", "outline"]

# Alignment options for cells
"""Horizontal alignment options supported for cells.

Constraints:
    - Must be one of the literal values: 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'.
"""
HorizontalAlignment = Literal["left", "center", "right", "fill", "justify", "centerContinuous", "distributed"]
"""Vertical alignment options supported for cells.

Constraints:
    - Must be one of: 'top', 'center', 'bottom', 'justify', 'distributed'.
"""
VerticalAlignment = Literal["top", "center", "bottom", "justify", "distributed"]

# Data validation operators
"""Data validation operator names used by data validation rules.

Constraints:
    - Must be one of the known operator Literal values (between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual).
"""
ValidationOperator = Literal[
    "between",
    "notBetween",
    "equal",
    "notEqual",
    "greaterThan",
    "lessThan",
    "greaterThanOrEqual",
    "lessThanOrEqual",
]
