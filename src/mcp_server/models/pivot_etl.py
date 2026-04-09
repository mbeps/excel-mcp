from __future__ import annotations

"""Schemas for pivot ETL operations: pivot results and paginated chunk reads.
"""

from pydantic import BaseModel, Field

from .common import CellScalar


class PivotResult(BaseModel):
    """Result of a pivot table operation.

    Attributes:
        data (list[dict[str, CellScalar]]): Pivoted rows as dictionaries (required).
        index_columns (list[str]): Columns used as the pivot index (required).
        value_columns (list[str]): Columns aggregated in the pivot (required).
        operation (str): Aggregation function applied (sum, mean, count, etc.). Required.
    """

    data: list[dict[str, CellScalar]] = Field(..., description="Pivoted data rows as dictionaries.")
    index_columns: list[str] = Field(..., description="Columns used as the pivot index.")
    value_columns: list[str] = Field(..., description="Columns aggregated in the pivot.")
    operation: str = Field(..., description="Aggregation function applied (sum, mean, count, etc.).")


class ChunkReadResult(BaseModel):
    """A paginated chunk of rows from a large dataset.

    Attributes:
        rows (list[dict[str, CellScalar]]): Rows in this chunk (required).
        chunk_start (int): 0-based starting row index (required).
        chunk_size (int): Number of rows returned (required).
        has_more (bool): True if additional rows remain (required).
        next_start_row (int | None): Next start row to request, or None if has_more is False.
    """

    rows: list[dict[str, CellScalar]] = Field(..., description="Rows in this chunk as column-keyed dicts.")
    chunk_start: int = Field(..., description="0-based starting row index of this chunk.")
    chunk_size: int = Field(..., description="Number of rows in this chunk.")
    has_more: bool = Field(..., description="Whether more rows remain after this chunk.")
    next_start_row: int | None = Field(
        None,
        description="Row number to pass as start_row to get the next chunk; None if has_more=False.",
    )
