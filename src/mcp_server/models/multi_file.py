from __future__ import annotations

"""Per-file aggregated/filtering schemas and diff row shapes used by multi-file tools.
"""

from typing import TypedDict

from .common import CellScalar


class MultiFilePerFileResult(TypedDict):
    """Per-file result row from multi-file aggregation operations.

    Keys:
        file (str): File path or identifier.
        row_count (int): Number of rows read from the file.
        value (float): Aggregated value for the file.
    """

    file: str
    row_count: int
    value: float


class MultiFileAggResult(TypedDict):
    """Aggregate result from bulk aggregation across multiple files.

    Keys:
        column (str): Column aggregated.
        operation (str): Aggregation operation (sum, mean, etc.).
        aggregate (float): Global aggregate.
        per_file (list[MultiFilePerFileResult]): Per-file breakdown.
    """

    column: str
    operation: str
    aggregate: float
    per_file: list[MultiFilePerFileResult]


class MultiFileFilterPerFileResult(TypedDict):
    """Per-file filter result used by bulk filter operations.

    Keys:
        file (str): File path or identifier.
        total_rows (int): Total rows examined.
        matched_rows (int): Rows matching criteria.
        data (list[list[CellScalar]]): Matched row data preview.
    """

    file: str
    total_rows: int
    matched_rows: int
    data: list[list[CellScalar]]


class WorkbookDiff(TypedDict):
    """A single differing cell between two compared workbooks.

    Keys:
        sheet (str): Sheet name where difference was found.
        cell (str): Cell reference in A1 notation.
        value_a, value_b (CellScalar): Values from workbook A and B.
    """

    sheet: str
    cell: str
    value_a: CellScalar
    value_b: CellScalar
