from __future__ import annotations

from mcp_server.tools.row_col import delete_cols, delete_rows, insert_cols, insert_rows
from mcp_server.tools.workbook import get_sheet_summary


def test_insert_rows(sample_xlsx: str) -> None:
    before = get_sheet_summary(sample_xlsx, "Sheet1")
    insert_rows(sample_xlsx, "Sheet1", row_index=2, count=3)
    after = get_sheet_summary(sample_xlsx, "Sheet1")
    assert after["row_count"] == before["row_count"] + 3


def test_delete_rows(sample_xlsx: str) -> None:
    before = get_sheet_summary(sample_xlsx, "Sheet1")
    delete_rows(sample_xlsx, "Sheet1", row_index=2, count=2)
    after = get_sheet_summary(sample_xlsx, "Sheet1")
    assert after["row_count"] == before["row_count"] - 2


def test_insert_cols(sample_xlsx: str) -> None:
    before = get_sheet_summary(sample_xlsx, "Sheet1")
    insert_cols(sample_xlsx, "Sheet1", col_index=2, count=2)
    after = get_sheet_summary(sample_xlsx, "Sheet1")
    assert after["col_count"] == before["col_count"] + 2


def test_delete_cols(sample_xlsx: str) -> None:
    before = get_sheet_summary(sample_xlsx, "Sheet1")
    delete_cols(sample_xlsx, "Sheet1", col_index=1, count=1)
    after = get_sheet_summary(sample_xlsx, "Sheet1")
    assert after["col_count"] == before["col_count"] - 1
