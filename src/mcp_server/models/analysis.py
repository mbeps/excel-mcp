from __future__ import annotations

from pydantic import BaseModel, Field

from .common import CellScalar


class FilterResult(BaseModel):
    """Rows that matched a filter condition."""

    matched_rows: list[list[CellScalar]] = Field(..., description="Rows matching the filter criteria.")
    total_rows: int = Field(..., description="Total number of rows evaluated.")
    matched_count: int = Field(..., description="Number of rows that matched.")
    headers: list[str] = Field(..., description="Column headers for the returned rows.")


class ColumnStats(BaseModel):
    """Descriptive statistics for a single numeric column."""

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
    """Result of a group-by aggregation operation."""

    groups: list[dict[str, CellScalar]] = Field(
        ..., description="List of group dictionaries with keys and aggregated values."
    )
    group_by: str = Field(..., description="Column used for grouping.")
    operation: str = Field(..., description="Aggregation operation applied (sum, mean, count, etc.).")


class DuplicateResult(BaseModel):
    """Rows identified as duplicates."""

    duplicates: list[list[CellScalar]] = Field(..., description="Duplicate rows found.")
    count: int = Field(..., description="Number of duplicate rows.")
    headers: list[str] = Field(..., description="Column headers for the returned rows.")


class ValidationResult(BaseModel):
    """Result of a formula or expression validation."""

    valid: bool = Field(..., description="Whether the expression is valid.")
    tokens: list[dict[str, str]] = Field(..., description="Parsed tokens from the expression.")
    error: str | None = Field(None, description="Error message if validation failed, or None.")
