from __future__ import annotations

from pydantic import BaseModel, Field

from .common import CellScalar


class CsvPreview(BaseModel):
    """Preview of a CSV file's contents."""

    headers: list[str] = Field(..., description="Column headers from the first row.")
    rows: list[list[CellScalar]] = Field(..., description="Preview rows of data.")
    total_rows: int = Field(..., description="Total number of rows in the CSV file.")


class DataProfile(BaseModel):
    """Statistical profile of a dataset."""

    columns: list[str] = Field(..., description="Column names in the dataset.")
    row_count: int = Field(..., description="Total number of data rows.")
    missing_values: dict[str, int] = Field(..., description="Count of missing values per column.")
    duplicates: int = Field(..., description="Number of fully duplicate rows.")
    summary_statistics: dict[str, dict[str, float | int | None]] = Field(
        ..., description="Per-column summary statistics (count, mean, std, min, max, etc.)."
    )
