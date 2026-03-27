from __future__ import annotations

import openpyxl

from mcp_server.tools.formulas import (
    set_formula,
    set_formulas_batch,
)


def test_set_formula(sample_xlsx: str) -> None:
    set_formula(sample_xlsx, "Sheet1", "E2", "=SUM(B2:B6)")
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws["E2"].value == "=SUM(B2:B6)"
    wb.close()


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


def test_set_formula_array(sample_xlsx: str) -> None:
    """Test set_formula with is_array=True sets an ArrayFormula on the cell."""
    from openpyxl.worksheet.formula import ArrayFormula

    result = set_formula(sample_xlsx, "Sheet1", "E2", "=SUM(B2:B6)", is_array=True, target_range="E2:E6")
    assert "Array formula" in result
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert isinstance(ws["E2"].value, ArrayFormula)
    wb.close()


from mcp_server.tools.formulas import (
    get_formula_value,
    get_formula_errors,
    get_formula_precedents,
    get_formula_dependents,
    list_formulas,
)


def test_get_formula_value(sample_xlsx: str) -> None:
    # Set a formula first
    set_formula(sample_xlsx, "Sheet1", "E2", "=SUM(B2:B3)")
    # Note: openpyxl doesn't evaluate formulas, so value will be None in a new workbook
    # But get_formula_value uses data_only=True which returns the cached value
    result = get_formula_value(sample_xlsx, "Sheet1", "E2")
    assert result["cell"] == "E2"
    # value might be None because we just wrote it and it's not cached by Excel yet
    assert "value" in result


def test_get_formula_errors(sample_xlsx: str) -> None:
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    ws["E10"] = "#DIV/0!"
    ws["E11"] = "#REF!"
    wb.save(sample_xlsx)
    wb.close()

    result = get_formula_errors(sample_xlsx, "Sheet1")
    assert result["count"] >= 2
    cells = [err["cell"] for err in result["errors"]]
    assert "E10" in cells
    assert "E11" in cells

    # Test with range
    result_range = get_formula_errors(sample_xlsx, "Sheet1", cell_range="E10:E10")
    assert result_range["count"] == 1
    assert result_range["errors"][0]["cell"] == "E10"


def test_get_formula_precedents(sample_xlsx: str) -> None:
    set_formula(sample_xlsx, "Sheet1", "F1", "=SUM(A1:B2) + Sheet1!C3 + $D$4")
    result = get_formula_precedents(sample_xlsx, "Sheet1", "F1")
    assert "A1" in str(result["precedents"])
    assert "B2" in str(result["precedents"])
    assert "Sheet1!C3" in result["precedents"]
    assert "$D$4" in result["precedents"]


def test_get_formula_dependents(sample_xlsx: str) -> None:
    set_formula(sample_xlsx, "Sheet1", "G1", "=A1*2")
    set_formula(sample_xlsx, "Sheet1", "G2", "=SUM(A1:A10)")
    # G1 depends on A1
    result = get_formula_dependents(sample_xlsx, "Sheet1", "A1")
    cells = [dep["cell"] for dep in result["dependents"]]
    assert "G1" in cells
    # G2 might not be caught by the current simple regex if it's in a range
    # but the regex "A1" in "SUM(A1:A10)" might match if not careful.
    # The tool uses: re.search(r"(?<![A-Z0-9])" + re.escape(cell_upper) + r"(?![A-Z0-9])", formula_clean)
    # "A1" in "SUM(A1:A10)":
    # before A1 is "(" -> OK
    # after A1 is ":" -> NOT OK (it's part of A1:A10)
    # So G2 should NOT be in dependents if it's looking for exact "A1" tokens.


def test_list_formulas(sample_xlsx: str) -> None:
    set_formula(sample_xlsx, "Sheet1", "H1", "=1+1")
    set_formula(sample_xlsx, "Sheet1", "H2", "=2+2")
    results = list_formulas(sample_xlsx, "Sheet1")
    cells = [f["cell_ref"] for f in results]
    assert "H1" in cells
    assert "H2" in cells


# ── additional coverage ────────────────────────────────────────────────────────


def test_set_formula_auto_prepend_equals(sample_xlsx: str) -> None:
    """Formula string without leading '=' has it automatically prepended."""
    set_formula(sample_xlsx, "Sheet1", "E2", "SUM(B2:B6)")
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"]["E2"].value == "=SUM(B2:B6)"
    wb.close()


def test_set_formulas_batch_auto_prepend(sample_xlsx: str) -> None:
    """set_formulas_batch auto-prepends '=' where missing."""
    set_formulas_batch(sample_xlsx, "Sheet1", {"F1": "A1+1", "F2": "=A2+1"})
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"]["F1"].value == "=A1+1"
    assert wb["Sheet1"]["F2"].value == "=A2+1"
    wb.close()


def test_set_formulas_batch_returns_count(sample_xlsx: str) -> None:
    """set_formulas_batch return message contains the number of formulas set."""
    result = set_formulas_batch(sample_xlsx, "Sheet1", {"G1": "=1", "G2": "=2", "G3": "=3"})
    assert "3" in result


def test_get_formula_value_plain_cell(sample_xlsx: str) -> None:
    """Plain-value cell returns its value with no 'note' key."""
    result = get_formula_value(sample_xlsx, "Sheet1", "A2")
    assert result["cell"] == "A2"
    assert result["value"] == "Alice"
    assert "note" not in result


def test_get_formula_value_uncached_formula_has_note(sample_xlsx: str) -> None:
    """Formula never evaluated by Excel returns None value with a hint note."""
    set_formula(sample_xlsx, "Sheet1", "E2", "=SUM(B2:B6)")
    result = get_formula_value(sample_xlsx, "Sheet1", "E2")
    assert result["cell"] == "E2"
    assert result["value"] is None
    assert "note" in result
    assert "formula" in result


def test_get_formula_errors_name_and_num_errors(sample_xlsx: str) -> None:
    """#NAME? and #NUM! strings are detected as formula errors."""
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    ws["F1"] = "#NAME?"
    ws["F2"] = "#NUM!"
    wb.save(sample_xlsx)
    wb.close()

    result = get_formula_errors(sample_xlsx, "Sheet1")
    cells = [e["cell"] for e in result["errors"]]
    assert "F1" in cells
    assert "F2" in cells


def test_get_formula_errors_empty_sheet(empty_xlsx: str) -> None:
    """Empty sheet returns zero formula errors."""
    result = get_formula_errors(empty_xlsx, "Sheet1")
    assert result["count"] == 0
    assert result["errors"] == []


def test_get_formula_precedents_no_formula(sample_xlsx: str) -> None:
    """Cell with no formula returns empty precedents list."""
    result = get_formula_precedents(sample_xlsx, "Sheet1", "A2")
    assert result["cell"] == "A2"
    assert result["precedents"] == []


def test_get_formula_precedents_captures_refs(sample_xlsx: str) -> None:
    """All cell references in a formula appear in the precedents output."""
    set_formula(sample_xlsx, "Sheet1", "I1", "=A1+B2")
    result = get_formula_precedents(sample_xlsx, "Sheet1", "I1")
    precedents_str = str(result["precedents"])
    assert "A1" in precedents_str
    assert "B2" in precedents_str


def test_get_formula_dependents_no_dependents(sample_xlsx: str) -> None:
    """Cell that nothing depends on returns empty dependents."""
    result = get_formula_dependents(sample_xlsx, "Sheet1", "Z99")
    assert result["dependents"] == []
    assert result["count"] == 0


def test_get_formula_dependents_multiple_cells(sample_xlsx: str) -> None:
    """All cells referencing the given cell are returned as dependents."""
    set_formula(sample_xlsx, "Sheet1", "J1", "=A1*2")
    set_formula(sample_xlsx, "Sheet1", "K1", "=A1+100")
    result = get_formula_dependents(sample_xlsx, "Sheet1", "A1")
    cells = [d["cell"] for d in result["dependents"]]
    assert "J1" in cells
    assert "K1" in cells
    assert result["count"] >= 2


def test_list_formulas_empty_sheet(empty_xlsx: str) -> None:
    """Sheet with no formulas returns an empty list."""
    result = list_formulas(empty_xlsx, "Sheet1")
    assert result == []


def test_list_formulas_returns_all(sample_xlsx: str) -> None:
    """list_formulas returns every formula cell with correct keys."""
    set_formula(sample_xlsx, "Sheet1", "L1", "=1+1")
    set_formula(sample_xlsx, "Sheet1", "L2", "=2+2")
    result = list_formulas(sample_xlsx, "Sheet1")
    cell_refs = [f["cell_ref"] for f in result]
    assert "L1" in cell_refs
    assert "L2" in cell_refs
    assert all(f["formula"].startswith("=") for f in result)


def test_set_formula_strips_curly_braces(sample_xlsx: str) -> None:
    """Curly braces in the formula string are stripped before storing."""
    set_formula(sample_xlsx, "Sheet1", "M1", "{=SUM(B2:B6)}", is_array=False)
    wb = openpyxl.load_workbook(sample_xlsx)
    assert wb["Sheet1"]["M1"].value == "=SUM(B2:B6)"
    wb.close()
