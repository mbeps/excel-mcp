from __future__ import annotations

"""CSV preview and dataset profiling schemas returned by CSV and cleaning tools.
"""

from pydantic import BaseModel, Field

from .common import CellScalar


class CsvPreview(BaseModel):
    """Preview of a CSV file's contents.

    Attributes:
        headers (list[str]): Column headers from the first row. Required.
        rows (list[list[CellScalar]]): Preview rows. Required.
        total_rows (int): Total number of rows in the file. Required.
    """

    headers: list[str] = Field(..., description="Column headers from the first row.")
    rows: list[list[CellScalar]] = Field(..., description="Preview rows of data.")
    total_rows: int = Field(..., description="Total number of rows in the CSV file.")


class DataProfile(BaseModel):
    """Statistical profile of a dataset.

    Attributes:
        columns (list[str]): Column names. Required.
        row_count (int): Number of data rows. Required.
        missing_values (dict[str, int]): Missing value counts per column. Required.
        duplicates (int): Number of fully duplicate rows. Required.
        summary_statistics (dict[str, dict[str, float | int | None]]): Per-column stats.
    """

    columns: list[str] = Field(..., description="Column names in the dataset.")
    row_count: int = Field(..., description="Total number of data rows.")
    missing_values: dict[str, int] = Field(..., description="Count of missing values per column.")
    duplicates: int = Field(..., description="Number of fully duplicate rows.")
    summary_statistics: dict[str, dict[str, float | int | None]] = Field(
        ..., description="Per-column summary statistics (count, mean, std, min, max, etc.)."
    )
