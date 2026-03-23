from __future__ import annotations

import openpyxl

from mcp_server.tools.formulas import (
    list_formulas,
    set_array_formula,
    set_formula,
    set_formulas_batch,
    validate_formula_syntax,
)


def test_set_formula(sample_xlsx: str) -> None:
    set_formula(sample_xlsx, "Sheet1", "E2", "=SUM(B2:B6)")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws["E2"].value == "=SUM(B2:B6)"
    wb.close()


def test_set_array_formula(sample_xlsx: str) -> None:
    set_array_formula(sample_xlsx, "Sheet1", "E2:E6", "=B2:B6*2")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    cell = ws["E2"]
    assert cell.value is not None
    wb.close()


def test_validate_formula_syntax_valid() -> None:
    result = validate_formula_syntax("=SUM(A1:A10)")
    assert result["valid"] is True
    assert result["error"] is None
    assert len(result["tokens"]) > 0


def test_validate_formula_syntax_prepends_equals() -> None:
    result = validate_formula_syntax("SUM(A1:A10)")
    assert result["valid"] is True


def test_set_formulas_batch(sample_xlsx: str) -> None:
    formulas = {
        "E2": "=SUM(B2:D2)",
        "E3": "=AVERAGE(B3:D3)",
    }
    set_formulas_batch(sample_xlsx, "Sheet1", formulas)
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws["E2"].value == "=SUM(B2:D2)"
    assert ws["E3"].value == "=AVERAGE(B3:D3)"
    wb.close()


def test_list_formulas(sample_xlsx: str) -> None:
    set_formula(sample_xlsx, "Sheet1", "F1", "=SUM(A1:A5)")
    result = list_formulas(sample_xlsx, "Sheet1")
    assert isinstance(result, list)
    assert any(r["cell_ref"] == "F1" for r in result)
    assert any("SUM" in r["formula"] for r in result)


def test_list_formulas_empty(empty_xlsx: str) -> None:
    result = list_formulas(empty_xlsx, "Sheet1")
    assert result == []
