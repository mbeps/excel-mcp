from __future__ import annotations

import openpyxl

from mcp_server.tools.hyperlinks import (
    add_hyperlink,
    delete_hyperlink,
    list_hyperlinks,
    read_hyperlink,
)


def test_add_hyperlink(sample_xlsx: str) -> None:
    add_hyperlink(sample_xlsx, "Sheet1", "A1", "https://example.com", display_text="Example")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws["A1"].hyperlink is not None
    assert ws["A1"].hyperlink.target == "https://example.com"
    assert ws["A1"].value == "Example"
    wb.close()


def test_read_hyperlink(sample_xlsx: str) -> None:
    add_hyperlink(sample_xlsx, "Sheet1", "B1", "https://test.org", tooltip="Test site")
    result = read_hyperlink(sample_xlsx, "Sheet1", "B1")
    assert result is not None
    assert result["target"] == "https://test.org"
    assert result["tooltip"] == "Test site"


def test_delete_hyperlink(sample_xlsx: str) -> None:
    add_hyperlink(sample_xlsx, "Sheet1", "C1", "https://remove.me")
    delete_hyperlink(sample_xlsx, "Sheet1", "C1")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws["C1"].hyperlink is None
    wb.close()


def test_list_hyperlinks(sample_xlsx: str) -> None:
    add_hyperlink(sample_xlsx, "Sheet1", "A1", "https://link1.com")
    add_hyperlink(sample_xlsx, "Sheet1", "B1", "https://link2.com")
    result = list_hyperlinks(sample_xlsx, "Sheet1")
    targets = [h["target"] for h in result]
    assert "https://link1.com" in targets
    assert "https://link2.com" in targets
    assert len(result) >= 2
