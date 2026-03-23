from __future__ import annotations

import openpyxl

import pytest

from mcp_server.tools.comments import (
    add_comment,
    add_comments_bulk,
    delete_comment,
    delete_comments_bulk,
    list_comments,
    read_comment,
    update_comment,
)


def test_add_comment(sample_xlsx: str) -> None:
    add_comment(sample_xlsx, "Sheet1", "A1", "This is a header", author="Tester")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws["A1"].comment is not None
    assert ws["A1"].comment.text == "This is a header"
    assert ws["A1"].comment.author == "Tester"
    wb.close()


def test_read_comment(sample_xlsx: str) -> None:
    add_comment(sample_xlsx, "Sheet1", "B2", "Age value", author="Bot")
    result = read_comment(sample_xlsx, "Sheet1", "B2")
    assert result is not None
    assert result["text"] == "Age value"
    assert result["author"] == "Bot"


def test_delete_comment(sample_xlsx: str) -> None:
    add_comment(sample_xlsx, "Sheet1", "C1", "Temporary note")
    delete_comment(sample_xlsx, "Sheet1", "C1")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws["C1"].comment is None
    wb.close()


def test_list_comments(sample_xlsx: str) -> None:
    add_comment(sample_xlsx, "Sheet1", "A1", "Comment 1", author="A1")
    add_comment(sample_xlsx, "Sheet1", "B1", "Comment 2", author="B1")
    result = list_comments(sample_xlsx, "Sheet1")
    refs = [c["cell_ref"] for c in result]
    assert "A1" in refs
    assert "B1" in refs
    assert len(result) >= 2


def test_update_comment_text(sample_xlsx: str) -> None:
    add_comment(sample_xlsx, "Sheet1", "A2", "Original text", author="Author")
    result = update_comment(sample_xlsx, "Sheet1", "A2", "Updated text")
    assert "updated" in result.lower()
    updated = read_comment(sample_xlsx, "Sheet1", "A2")
    assert updated is not None
    assert updated["text"] == "Updated text"
    assert updated["author"] == "Author"  # author preserved when not provided


def test_update_comment_with_author(sample_xlsx: str) -> None:
    add_comment(sample_xlsx, "Sheet1", "B2", "Hello", author="OldAuthor")
    update_comment(sample_xlsx, "Sheet1", "B2", "New text", author="NewAuthor")
    updated = read_comment(sample_xlsx, "Sheet1", "B2")
    assert updated is not None
    assert updated["text"] == "New text"
    assert updated["author"] == "NewAuthor"


def test_update_comment_missing_raises(sample_xlsx: str) -> None:
    with pytest.raises(ValueError, match="No comment found"):
        update_comment(sample_xlsx, "Sheet1", "Z99", "Should fail")


def test_add_comments_bulk(sample_xlsx: str) -> None:
    entries = [
        {"cell": "A1", "text": "Bulk comment 1", "author": "Tester"},
        {"cell": "B2", "text": "Bulk comment 2", "author": "Tester"},
        {"cell": "C3", "text": "Bulk comment 3"},
    ]
    result = add_comments_bulk(sample_xlsx, "Sheet1", entries)
    assert result["added"] == 3
    comments = list_comments(sample_xlsx, "Sheet1")
    refs = [c["cell_ref"] for c in comments]
    assert "A1" in refs
    assert "B2" in refs
    assert "C3" in refs


def test_delete_comments_bulk(sample_xlsx: str) -> None:
    add_comment(sample_xlsx, "Sheet1", "A1", "Keep me")
    add_comment(sample_xlsx, "Sheet1", "B1", "Delete me 1")
    add_comment(sample_xlsx, "Sheet1", "C1", "Delete me 2")
    result = delete_comments_bulk(sample_xlsx, "Sheet1", ["B1", "C1"])
    assert result["deleted"] == 2
    comments = list_comments(sample_xlsx, "Sheet1")
    refs = [c["cell_ref"] for c in comments]
    assert "A1" in refs
    assert "B1" not in refs
    assert "C1" not in refs
