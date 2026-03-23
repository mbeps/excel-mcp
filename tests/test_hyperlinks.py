from __future__ import annotations

import openpyxl
import pytest

from mcp_server.tools.hyperlinks import (
    add_hyperlink,
    add_internal_hyperlink,
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


def test_add_internal_hyperlink(tmp_path: str) -> None:
    fp = str(tmp_path / "internal.xlsx")
    wb = openpyxl.Workbook()
    wb.active.title = "Sheet1"
    wb.create_sheet("Sheet2")
    wb.save(fp)
    wb.close()
    result = add_internal_hyperlink(fp, "Sheet1", "A1", "Sheet2", "B2")
    assert "Sheet2" in result and "B2" in result
    wb2 = openpyxl.load_workbook(fp)
    ws = wb2["Sheet1"]
    assert ws["A1"].hyperlink is not None
    assert ws["A1"].hyperlink.target == "#Sheet2!B2"
    wb2.close()


def test_add_internal_hyperlink_invalid_sheet(tmp_path: str) -> None:
    fp = str(tmp_path / "internal.xlsx")
    wb = openpyxl.Workbook()
    wb.save(fp)
    wb.close()
    with pytest.raises(ValueError, match="Sheet 'NonExistent' not found"):
        add_internal_hyperlink(fp, "Sheet", "A1", "NonExistent")
