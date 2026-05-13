"""Schemas for pivot ETL operations: pivot results and paginated chunk reads."""

from typing import TypedDict

from pydantic import BaseModel, Field


class NativePivotValueField(BaseModel):
    field: str = Field(..., description="Source column name to aggregate.")
    aggfunc: str = Field(
        "sum",
        description=(
            "Aggregation function. One of: sum, count, average, max, min, "
            "product, countNums, stdDev, stdDevp, var, varp."
        ),
    )
    name: str | None = Field(
        None,
        description="Display name shown in the pivot header. Auto-generated if omitted.",
    )
    show_data_as: str | None = Field(
        None,
        description=(
            "How values are displayed. One of: normal, difference, percent, "
            "percentDiff, runTotal, percentOfRow, percentOfCol, percentOfTotal, index. "
            "Omit for default numeric display."
        ),
    )

class NativePivotResult(TypedDict):
    file: str
    output_sheet: str
    pivot_table_name: str
    source_range: str
    pivot_location: str
    row_fields: list[str]
    col_fields: list[str]
    value_fields: list[str]
    filter_fields: list[str]
