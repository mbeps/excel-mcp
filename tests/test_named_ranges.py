from __future__ import annotations

import openpyxl

import pytest

from mcp_server.tools.named_ranges import (
    create_named_range,
    delete_named_range,
    list_named_ranges,
    rename_named_range,
    update_named_range,
)


def test_create_named_range(sample_xlsx: str) -> None:
    create_named_range(sample_xlsx, "SalesData", "Sheet1!$A$1:$D$6")
    wb = openpyxl.load_workbook(sample_xlsx)
    assert "SalesData" in wb.defined_names
    assert wb.defined_names["SalesData"].attr_text == "Sheet1!$A$1:$D$6"
    wb.close()


def test_list_named_ranges(sample_xlsx: str) -> None:
    create_named_range(sample_xlsx, "MyRange", "Sheet1!$B$1:$B$6")
    result = list_named_ranges(sample_xlsx)
    assert len(result) == 1
    assert result[0]["name"] == "MyRange"
    assert result[0]["destination"] == "Sheet1!$B$1:$B$6"
    assert result[0]["scope"] == "workbook"


def test_delete_named_range(sample_xlsx: str) -> None:
    create_named_range(sample_xlsx, "TempRange", "Sheet1!$A$1:$A$6")
    delete_named_range(sample_xlsx, "TempRange")
    wb = openpyxl.load_workbook(sample_xlsx)
    assert "TempRange" not in wb.defined_names
    wb.close()


def test_update_named_range(sample_xlsx: str) -> None:
    create_named_range(sample_xlsx, "UpdateMe", "Sheet1!$A$1:$A$3")
    update_named_range(sample_xlsx, "UpdateMe", "Sheet1!$A$1:$D$6")
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb.defined_names["UpdateMe"].attr_text == "Sheet1!$A$1:$D$6"
    wb.close()


def test_rename_named_range(sample_xlsx: str) -> None:
    create_named_range(sample_xlsx, "OldRangeName", "Sheet1!$A$1:$B$3")
    result = rename_named_range(sample_xlsx, "OldRangeName", "NewRangeName")
    assert "NewRangeName" in result
    wb = openpyxl.load_workbook(sample_xlsx)
    assert "NewRangeName" in wb.defined_names
    assert "OldRangeName" not in wb.defined_names
    assert wb.defined_names["NewRangeName"].attr_text == "Sheet1!$A$1:$B$3"
    wb.close()


def test_rename_named_range_not_found(sample_xlsx: str) -> None:
    with pytest.raises(ValueError, match="not found"):
        rename_named_range(sample_xlsx, "DoesNotExist", "AnyName")


def test_rename_named_range_target_exists(sample_xlsx: str) -> None:
    create_named_range(sample_xlsx, "Alpha", "Sheet1!$A$1")
    create_named_range(sample_xlsx, "Beta", "Sheet1!$B$1")
    with pytest.raises(ValueError, match="already exists"):
        rename_named_range(sample_xlsx, "Alpha", "Beta")
