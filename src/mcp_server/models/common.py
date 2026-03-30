from __future__ import annotations

from datetime import date, datetime
from typing import Literal, TypeAlias

CellScalar: TypeAlias = str | int | float | bool | datetime | date | None
ScenarioCellValue: TypeAlias = float | str | int | bool | None

# Border styles supported by openpyxl
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
BorderSide = Literal["left", "right", "top", "bottom", "diagonal", "outline"]

# Alignment options for cells
HorizontalAlignment = Literal["left", "center", "right", "fill", "justify", "centerContinuous", "distributed"]
VerticalAlignment = Literal["top", "center", "bottom", "justify", "distributed"]

# Data validation operators
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
