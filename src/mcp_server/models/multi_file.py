from __future__ import annotations

from typing import TypedDict

from .common import CellScalar


class MultiFilePerFileResult(TypedDict):
    """Per-file result row from bulk_aggregate_multi_files."""

    file: str
    row_count: int
    value: float


class MultiFileAggResult(TypedDict):
    """Aggregate result from bulk_aggregate_multi_files."""

    column: str
    operation: str
    aggregate: float
    per_file: list[MultiFilePerFileResult]


class MultiFileFilterPerFileResult(TypedDict):
    """Per-file filter result from bulk_filter_multi_files."""

    file: str
    total_rows: int
    matched_rows: int
    data: list[list[CellScalar]]


class WorkbookDiff(TypedDict):
    """A single differing cell between two compared workbooks."""

    sheet: str
    cell: str
    value_a: CellScalar
    value_b: CellScalar
