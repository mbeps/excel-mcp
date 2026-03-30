from __future__ import annotations

from pydantic import BaseModel, Field

from .common import CellScalar


class PivotResult(BaseModel):
    """Result of a pivot table operation."""

    data: list[dict[str, CellScalar]] = Field(..., description="Pivoted data rows as dictionaries.")
    index_columns: list[str] = Field(..., description="Columns used as the pivot index.")
    value_columns: list[str] = Field(..., description="Columns aggregated in the pivot.")
    operation: str = Field(..., description="Aggregation function applied (sum, mean, count, etc.).")


class ChunkReadResult(BaseModel):
    """A paginated chunk of rows from a large dataset."""

    rows: list[dict[str, CellScalar]] = Field(..., description="Rows in this chunk as column-keyed dicts.")
    chunk_start: int = Field(..., description="0-based starting row index of this chunk.")
    chunk_size: int = Field(..., description="Number of rows in this chunk.")
    has_more: bool = Field(..., description="Whether more rows remain after this chunk.")
    next_start_row: int | None = Field(
        None,
        description="Row number to pass as start_row to get the next chunk; None if has_more=False.",
    )
