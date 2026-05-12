"""Schemas for data analysis and profiling results.

Used by analysis tools to return structured results (filter hits, aggregates, column stats).
"""

from __future__ import annotations

from pydantic import BaseModel, Field

from .common import CellScalar


class SortDescriptor(BaseModel):
    """A single sort criterion for multi-column sorting.

    Attributes:
        column (str): Column name or letter to sort by.
        ascending (bool): Sort ascending (True) or descending (False). Defaults to True.
    """

    column: str = Field(..., description="Column name or letter to sort by.")
    ascending: bool = Field(True, description="Sort ascending (True) or descending (False).")


class FilterCondition(BaseModel):
    """A single filter condition for advanced data filtering.

    Attributes:
        column (str): Column name to filter on.
        operator (str): Comparison operator: '==', '!=', '>', '<', '>=', '<=', 'contains', 'startswith', 'endswith'.
        value (str | int | float | bool): Value to compare against.
    """

    column: str = Field(..., description="Column name to filter on.")
    operator: str = Field(
        ...,
        description="Comparison operator: '==', '!=', '>', '<', '>=', '<=', 'contains', 'startswith', 'endswith'.",
    )
    value: str | int | float | bool = Field(..., description="Value to compare against.")


class FilterResult(BaseModel):
    """Rows that matched a filter condition.

    Attributes:
        matched_rows (list[list[CellScalar]]): Rows matching the filter criteria. Required.
        total_rows (int): Total number of rows evaluated. Required.
        matched_count (int): Number of rows that matched. Required.
        headers (list[str]): Column headers for returned rows. Required.
    """

    matched_rows: list[list[CellScalar]] = Field(..., description="Rows matching the filter criteria.")
    total_rows: int = Field(..., description="Total number of rows evaluated.")
    matched_count: int = Field(..., description="Number of rows that matched.")
    headers: list[str] = Field(..., description="Column headers for the returned rows.")


class ColumnStats(BaseModel):
    """Descriptive statistics for a single numeric column.

    Attributes:
        column (str): Column letter or header name. Required.
        count (int): Number of non-empty numeric values. Required.
        mean (float | None): Arithmetic mean, or None if unavailable.
        median (float | None): Median value, or None if unavailable.
        min_val, max_val, std, sum_val (float | None): Summary values; None if unavailable.
        skewness, kurtosis (float | None): Higher-order moments; may be None when insufficient data.
        message (str | None): Informational message (e.g., non-numeric column).
    """

    column: str = Field(..., description="Column letter or header name.")
    count: int = Field(..., description="Number of non-empty numeric values.")
    mean: float | None = Field(None, description="Arithmetic mean, or None if unavailable.")
    median: float | None = Field(None, description="Median value, or None if unavailable.")
    min_val: float | None = Field(None, description="Minimum value, or None if unavailable.")
    max_val: float | None = Field(None, description="Maximum value, or None if unavailable.")
    std: float | None = Field(None, description="Standard deviation, or None if unavailable.")
    sum_val: float | None = Field(None, description="Sum of values, or None if unavailable.")
    skewness: float | None = Field(None, description="Skewness of the distribution, or None if insufficient data.")
    kurtosis: float | None = Field(None, description="Kurtosis of the distribution, or None if insufficient data.")
    message: str | None = Field(None, description="Informational message, e.g. when column is non-numeric.")
