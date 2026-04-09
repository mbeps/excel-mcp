from __future__ import annotations

"""Schemas for data analysis and profiling results.

Used by analysis tools to return structured results (filter hits, aggregates, column stats).
"""

from pydantic import BaseModel, Field

from .common import CellScalar


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


class AggregateResult(BaseModel):
    """Result of a group-by aggregation operation.

    Attributes:
        groups (list[dict[str, CellScalar]]): Group dictionaries with keys and aggregated values. Required.
        group_by (str): Column used for grouping. Required.
        operation (str): Aggregation operation applied (sum, mean, count, etc.). Required.
    """

    groups: list[dict[str, CellScalar]] = Field(
        ..., description="List of group dictionaries with keys and aggregated values."
    )
    group_by: str = Field(..., description="Column used for grouping.")
    operation: str = Field(..., description="Aggregation operation applied (sum, mean, count, etc.).")


class DuplicateResult(BaseModel):
    """Rows identified as duplicates.

    Attributes:
        duplicates (list[list[CellScalar]]): Duplicate rows found. Required.
        count (int): Number of duplicate rows. Required.
        headers (list[str]): Column headers for returned rows. Required.
    """

    duplicates: list[list[CellScalar]] = Field(..., description="Duplicate rows found.")
    count: int = Field(..., description="Number of duplicate rows.")
    headers: list[str] = Field(..., description="Column headers for the returned rows.")


class ValidationResult(BaseModel):
    """Result of expression/formula validation.

    Attributes:
        valid (bool): Whether the expression is valid. Required.
        tokens (list[dict[str, str]]): Parsed tokens (structure: list of {'type': str, 'value': str}).
        error (str | None): Error message if validation failed.

    Notes:
        - Consumers should not assume token shapes beyond a simple dict with string keys.
    """

    valid: bool = Field(..., description="Whether the expression is valid.")
    tokens: list[dict[str, str]] = Field(..., description="Parsed tokens from the expression.")
    error: str | None = Field(None, description="Error message if validation failed, or None.")
