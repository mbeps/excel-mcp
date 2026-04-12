"""Workflow tests for formula & calculation operations.

Tests chain multiple MCP tool calls simulating real user workflows and verify
results independently using openpyxl — never via MCP tools.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from mcp_server.tools.cell_ops import auto_sum, fill_formula, write_cell, write_range
from mcp_server.tools.formulas import (
    get_formula_dependents,
    get_formula_errors,
    get_formula_precedents,
    get_formula_value,
    list_formulas,
    set_formula,
    set_formulas_batch,
)
from mcp_server.tools.workbook import create_workbook


# ---------------------------------------------------------------------------
# 1. Set formula and verify
# ---------------------------------------------------------------------------


class TestSetFormulaAndVerify:
    def test_set_formula_and_verify(self, tmp_path: Path) -> None:
        """set_formula with =A1+B1 → verify formula string with openpyxl."""
        fp = str(tmp_path / "formulas.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", 10)
        write_cell(fp, "Sheet1", "B1", 20)
        set_formula(fp, "Sheet1", "C1", "=A1+B1")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["C1"].value == "=A1+B1"
        wb.close()


# ---------------------------------------------------------------------------
# 2. Batch set formulas
# ---------------------------------------------------------------------------


class TestBatchSetFormulas:
    def test_batch_set_formulas(self, tmp_path: Path) -> None:
        """batch_set_formulas with 5+ formulas → verify all with openpyxl."""
        fp = str(tmp_path / "batch.xlsx")
        create_workbook(fp, sheet_names=["Data"])

        write_range(
            fp,
            "Data",
            "A1",
            [
                [10, 20],
                [30, 40],
                [50, 60],
                [70, 80],
                [90, 100],
            ],
        )

        formulas = {
            "C1": "=A1+B1",
            "C2": "=A2+B2",
            "C3": "=A3*B3",
            "C4": "=A4-B4",
            "C5": "=A5/B5",
        }
        result = set_formulas_batch(fp, "Data", formulas)
        assert "5" in result

        wb = openpyxl.load_workbook(fp)
        ws = wb["Data"]
        assert ws["C1"].value == "=A1+B1"
        assert ws["C2"].value == "=A2+B2"
        assert ws["C3"].value == "=A3*B3"
        assert ws["C4"].value == "=A4-B4"
        assert ws["C5"].value == "=A5/B5"
        wb.close()


# ---------------------------------------------------------------------------
# 3. Fill formula down column
# ---------------------------------------------------------------------------


class TestFillFormulaDownColumn:
    def test_fill_formula_down_column(self, tmp_path: Path) -> None:
        """Set formula in one cell → fill_formula down → verify translated formulas."""
        fp = str(tmp_path / "fill_down.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_range(
            fp,
            "Sheet1",
            "A1",
            [
                [1, 10],
                [2, 20],
                [3, 30],
                [4, 40],
            ],
        )

        set_formula(fp, "Sheet1", "C1", "=A1+B1")
        fill_formula(fp, "Sheet1", "C1", "C2:C4")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["C1"].value == "=A1+B1"
        assert ws["C2"].value == "=A2+B2"
        assert ws["C3"].value == "=A3+B3"
        assert ws["C4"].value == "=A4+B4"
        wb.close()


# ---------------------------------------------------------------------------
# 4. Fill formula across row
# ---------------------------------------------------------------------------


class TestFillFormulaAcrossRow:
    def test_fill_formula_across_row(self, tmp_path: Path) -> None:
        """Fill formula right → verify translation across columns."""
        fp = str(tmp_path / "fill_right.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_range(
            fp,
            "Sheet1",
            "A1",
            [
                [10, 20, 30, 40],
                [1, 2, 3, 4],
            ],
        )

        set_formula(fp, "Sheet1", "A3", "=A1+A2")
        fill_formula(fp, "Sheet1", "A3", "B3:D3")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A3"].value == "=A1+A2"
        assert ws["B3"].value == "=B1+B2"
        assert ws["C3"].value == "=C1+C2"
        assert ws["D3"].value == "=D1+D2"
        wb.close()


# ---------------------------------------------------------------------------
# 5. Formula with absolute references
# ---------------------------------------------------------------------------


class TestFormulaWithAbsoluteRefs:
    def test_formula_with_absolute_refs(self, tmp_path: Path) -> None:
        """Set formula with $A$1 → fill → verify $ refs preserved."""
        fp = str(tmp_path / "abs_ref.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", 100)
        write_range(fp, "Sheet1", "B1", [[10], [20], [30]])

        set_formula(fp, "Sheet1", "C1", "=$A$1*B1")
        fill_formula(fp, "Sheet1", "C1", "C2:C3")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["C1"].value == "=$A$1*B1"
        assert ws["C2"].value == "=$A$1*B2"
        assert ws["C3"].value == "=$A$1*B3"
        wb.close()


# ---------------------------------------------------------------------------
# 6. Audit formula value
# ---------------------------------------------------------------------------


class TestAuditFormulaValue:
    def test_audit_formula_value(self, tmp_path: Path) -> None:
        """set_formula → write data → audit_formula_value → verify result structure."""
        fp = str(tmp_path / "audit_val.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", 5)
        write_cell(fp, "Sheet1", "B1", 15)
        set_formula(fp, "Sheet1", "C1", "=A1+B1")

        result = get_formula_value(fp, "Sheet1", "C1")
        assert result["cell"] == "C1"
        # openpyxl can't evaluate — cached value is None since not saved by Excel
        assert "value" in result

        # Verify the formula is actually stored
        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["C1"].value == "=A1+B1"
        wb.close()


# ---------------------------------------------------------------------------
# 7. Audit formula errors
# ---------------------------------------------------------------------------


class TestAuditFormulaErrors:
    def test_audit_formula_errors(self, tmp_path: Path) -> None:
        """Write error values → audit_formula_errors → verify errors detected."""
        fp = str(tmp_path / "audit_err.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        # Write cells with Excel error strings
        write_cell(fp, "Sheet1", "A1", "#REF!")
        write_cell(fp, "Sheet1", "B1", "#DIV/0!")
        write_cell(fp, "Sheet1", "C1", "normal")

        result = get_formula_errors(fp, "Sheet1")
        assert result["count"] == 2
        errors = result["errors"]
        error_cells = {e["cell"] for e in errors}
        assert "A1" in error_cells
        assert "B1" in error_cells

    def test_audit_formula_errors_in_range(self, tmp_path: Path) -> None:
        """Audit errors scoped to a specific cell range."""
        fp = str(tmp_path / "audit_range.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", "#VALUE!")
        write_cell(fp, "Sheet1", "A2", 42)
        write_cell(fp, "Sheet1", "B1", "#N/A")

        result = get_formula_errors(fp, "Sheet1", cell_range="A1:A2")
        assert result["count"] == 1
        assert result["errors"][0]["cell"] == "A1"


# ---------------------------------------------------------------------------
# 8. Find precedents — single ref
# ---------------------------------------------------------------------------


class TestFindPrecedentsSingleRef:
    def test_find_precedents_single_ref(self, tmp_path: Path) -> None:
        """Formula =A1+B1 → find_precedents → verify [A1, B1]."""
        fp = str(tmp_path / "prec_single.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", 10)
        write_cell(fp, "Sheet1", "B1", 20)
        set_formula(fp, "Sheet1", "C1", "=A1+B1")

        result = get_formula_precedents(fp, "Sheet1", "C1")
        assert result["cell"] == "C1"
        assert result["formula"] == "=A1+B1"
        precs = result["precedents"]
        assert "A1" in precs
        assert "B1" in precs


# ---------------------------------------------------------------------------
# 9. Find precedents — range ref
# ---------------------------------------------------------------------------


class TestFindPrecedentsRangeRef:
    def test_find_precedents_range_ref(self, tmp_path: Path) -> None:
        """Formula =SUM(A1:A5) → find_precedents → verify A1..A5 all present."""
        fp = str(tmp_path / "prec_range.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        for i in range(1, 6):
            write_cell(fp, "Sheet1", f"A{i}", i * 10)
        set_formula(fp, "Sheet1", "B1", "=SUM(A1:A5)")

        result = get_formula_precedents(fp, "Sheet1", "B1")
        assert result["formula"] == "=SUM(A1:A5)"
        precs = result["precedents"]
        for i in range(1, 6):
            assert f"A{i}" in precs


# ---------------------------------------------------------------------------
# 10. Find dependents
# ---------------------------------------------------------------------------


class TestFindDependents:
    def test_find_dependents(self, tmp_path: Path) -> None:
        """Set formula referencing cell → find_dependents on that cell → verify."""
        fp = str(tmp_path / "deps.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", 100)
        set_formula(fp, "Sheet1", "B1", "=A1*2")
        set_formula(fp, "Sheet1", "C1", "=A1+10")
        set_formula(fp, "Sheet1", "D1", "=B1+5")  # depends on B1 not A1

        result = get_formula_dependents(fp, "Sheet1", "A1")
        assert result["cell"] == "A1"
        dep_cells = {d["cell"] for d in result["dependents"]}
        assert "B1" in dep_cells
        assert "C1" in dep_cells
        assert "D1" not in dep_cells  # D1 references B1, not A1 directly


# ---------------------------------------------------------------------------
# 11. List formulas
# ---------------------------------------------------------------------------


class TestListFormulas:
    def test_list_formulas(self, tmp_path: Path) -> None:
        """Set multiple formulas → list_formulas → verify all listed."""
        fp = str(tmp_path / "list_f.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", 10)
        write_cell(fp, "Sheet1", "A2", 20)
        set_formula(fp, "Sheet1", "B1", "=A1*2")
        set_formula(fp, "Sheet1", "B2", "=A2*3")
        set_formula(fp, "Sheet1", "C1", "=SUM(A1:A2)")

        result = list_formulas(fp, "Sheet1")
        assert len(result) == 3
        refs = {f["cell_ref"] for f in result}
        assert refs == {"B1", "B2", "C1"}
        formulas_map = {f["cell_ref"]: f["formula"] for f in result}
        assert formulas_map["B1"] == "=A1*2"
        assert formulas_map["B2"] == "=A2*3"
        assert formulas_map["C1"] == "=SUM(A1:A2)"


# ---------------------------------------------------------------------------
# 12. Auto-sum with formulas
# ---------------------------------------------------------------------------


class TestAutoSumWithFormulas:
    def test_auto_sum_with_formulas(self, tmp_path: Path) -> None:
        """Write data → auto_sum → verify formula text."""
        fp = str(tmp_path / "autosum.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        for i in range(1, 6):
            write_cell(fp, "Sheet1", f"A{i}", i * 10)

        result = auto_sum(fp, "Sheet1", "A6")
        assert result["cell"] == "A6"
        assert "SUM" in result["formula"]

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        formula = ws["A6"].value
        assert formula.startswith("=SUM(")
        assert "A1" in formula
        assert "A5" in formula
        wb.close()

    def test_auto_sum_explicit_range(self, tmp_path: Path) -> None:
        """auto_sum with explicit source_range → verify exact formula."""
        fp = str(tmp_path / "autosum_exp.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_range(fp, "Sheet1", "B1", [[100], [200], [300]])

        result = auto_sum(fp, "Sheet1", "B4", source_range="B1:B3")
        assert result["formula"] == "=SUM(B1:B3)"

        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["B4"].value == "=SUM(B1:B3)"
        wb.close()


# ---------------------------------------------------------------------------
# 13. Formula chain
# ---------------------------------------------------------------------------


class TestFormulaChain:
    def test_formula_chain(self, tmp_path: Path) -> None:
        """A1=10, B1=A1*2, C1=B1+5 → verify chain with openpyxl."""
        fp = str(tmp_path / "chain.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", 10)
        set_formula(fp, "Sheet1", "B1", "=A1*2")
        set_formula(fp, "Sheet1", "C1", "=B1+5")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].value == 10
        assert ws["B1"].value == "=A1*2"
        assert ws["C1"].value == "=B1+5"
        wb.close()

    def test_formula_chain_dependents(self, tmp_path: Path) -> None:
        """Verify dependents of each cell in the chain."""
        fp = str(tmp_path / "chain_dep.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", 10)
        set_formula(fp, "Sheet1", "B1", "=A1*2")
        set_formula(fp, "Sheet1", "C1", "=B1+5")

        # A1 has dependent B1
        deps_a = get_formula_dependents(fp, "Sheet1", "A1")
        dep_cells_a = {d["cell"] for d in deps_a["dependents"]}
        assert "B1" in dep_cells_a

        # B1 has dependent C1
        deps_b = get_formula_dependents(fp, "Sheet1", "B1")
        dep_cells_b = {d["cell"] for d in deps_b["dependents"]}
        assert "C1" in dep_cells_b


# ---------------------------------------------------------------------------
# 14. Formula overwrite
# ---------------------------------------------------------------------------


class TestFormulaOverwrite:
    def test_formula_overwrite(self, tmp_path: Path) -> None:
        """Set formula → set different formula → verify new one."""
        fp = str(tmp_path / "overwrite.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", 10)
        set_formula(fp, "Sheet1", "B1", "=A1*2")

        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["B1"].value == "=A1*2"
        wb.close()

        set_formula(fp, "Sheet1", "B1", "=A1+100")

        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["B1"].value == "=A1+100"
        wb.close()


# ---------------------------------------------------------------------------
# 15. Complex nested formulas
# ---------------------------------------------------------------------------


class TestComplexNestedFormulas:
    def test_complex_nested_formulas(self, tmp_path: Path) -> None:
        """IF, nested SUM, nested formula strings → verify formula strings."""
        fp = str(tmp_path / "complex.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_range(
            fp,
            "Sheet1",
            "A1",
            [
                ["Score", "Grade"],
                [90, None],
                [75, None],
                [50, None],
            ],
        )

        set_formula(fp, "Sheet1", "B2", '=IF(A2>=80,"Pass","Fail")')
        set_formula(fp, "Sheet1", "B3", '=IF(A3>=80,"Pass","Fail")')
        set_formula(fp, "Sheet1", "B4", '=IF(A4>=80,"Pass","Fail")')
        set_formula(fp, "Sheet1", "C1", "=SUM(A2:A4)")
        set_formula(fp, "Sheet1", "D1", "=AVERAGE(A2:A4)")
        set_formula(fp, "Sheet1", "E1", "=IF(C1>200,SUM(A2:A4)*2,C1)")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["B2"].value == '=IF(A2>=80,"Pass","Fail")'
        assert ws["B3"].value == '=IF(A3>=80,"Pass","Fail")'
        assert ws["C1"].value == "=SUM(A2:A4)"
        assert ws["D1"].value == "=AVERAGE(A2:A4)"
        assert ws["E1"].value == "=IF(C1>200,SUM(A2:A4)*2,C1)"
        wb.close()

    def test_nested_sum_product(self, tmp_path: Path) -> None:
        """SUMPRODUCT and other complex built-in functions."""
        fp = str(tmp_path / "sumproduct.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_range(fp, "Sheet1", "A1", [[2, 3], [4, 5], [6, 7]])
        set_formula(fp, "Sheet1", "C1", "=SUMPRODUCT(A1:A3,B1:B3)")
        set_formula(fp, "Sheet1", "C2", '=COUNTIF(A1:A3,">3")')

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["C1"].value == "=SUMPRODUCT(A1:A3,B1:B3)"
        assert ws["C2"].value == '=COUNTIF(A1:A3,">3")'
        wb.close()


# ---------------------------------------------------------------------------
# 16. Set formula auto-prepends = sign
# ---------------------------------------------------------------------------


class TestFormulaAutoEquals:
    def test_formula_auto_prepends_equals(self, tmp_path: Path) -> None:
        """set_formula without leading '=' should auto-prepend it."""
        fp = str(tmp_path / "auto_eq.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", 5)
        set_formula(fp, "Sheet1", "B1", "A1*10")

        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["B1"].value == "=A1*10"
        wb.close()


# ---------------------------------------------------------------------------
# 17. Fill formula — single cell target
# ---------------------------------------------------------------------------


class TestFillFormulaSingleCell:
    def test_fill_formula_single_cell(self, tmp_path: Path) -> None:
        """fill_formula to a single cell target (no colon range)."""
        fp = str(tmp_path / "fill_single.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_range(fp, "Sheet1", "A1", [[10, 20], [30, 40]])
        set_formula(fp, "Sheet1", "C1", "=A1+B1")
        result = fill_formula(fp, "Sheet1", "C1", "C2")
        assert result["cells_filled"] == 1

        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["C2"].value == "=A2+B2"
        wb.close()


# ---------------------------------------------------------------------------
# 18. Precedents on non-formula cell
# ---------------------------------------------------------------------------


class TestPrecedentsNonFormula:
    def test_precedents_non_formula_cell(self, tmp_path: Path) -> None:
        """find_precedents on a plain-value cell returns empty list."""
        fp = str(tmp_path / "no_prec.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", 42)

        result = get_formula_precedents(fp, "Sheet1", "A1")
        assert result["precedents"] == []


# ---------------------------------------------------------------------------
# 19. Array formula
# ---------------------------------------------------------------------------


class TestArrayFormula:
    def test_set_array_formula(self, tmp_path: Path) -> None:
        """set_formula with is_array=True → verify ArrayFormula stored."""
        fp = str(tmp_path / "array.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_range(fp, "Sheet1", "A1", [[1, 2], [3, 4], [5, 6]])
        result = set_formula(
            fp,
            "Sheet1",
            "C1",
            "=A1:A3*B1:B3",
            is_array=True,
            target_range="C1:C3",
        )
        assert "Array formula" in result

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        # openpyxl stores array formulas — the cell should have a value
        cell_val = ws["C1"].value
        assert cell_val is not None
        wb.close()
