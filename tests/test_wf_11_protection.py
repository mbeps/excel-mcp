"""Workflow tests for protection and data validation.

Tests call tool functions directly and verify results independently using
openpyxl — never via MCP tools.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl

from mcp_server.tools.cell_ops import write_range
from mcp_server.tools.data_validation import (
    add_date_validation,
    add_dropdown_validation,
    add_formula_validation,
    add_numeric_validation,
    remove_validation,
)
from mcp_server.tools.doc_properties import protect_workbook, unprotect_workbook
from mcp_server.tools.protection import protect_cells, protect_sheet, unprotect_sheet
from mcp_server.tools.workbook import create_workbook

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _create_data_file(fp: str) -> None:
    """Create a workbook with sample data for protection tests."""
    create_workbook(fp, sheet_names=["Sheet1"])
    write_range(
        fp,
        "Sheet1",
        "A1",
        [
            ["Name", "Age", "City", "Salary"],
            ["Alice", 30, "New York", 70000],
            ["Bob", 25, "Chicago", 55000],
            ["Charlie", 35, "Boston", 90000],
        ],
    )


# ---------------------------------------------------------------------------
# 1. Protect sheet basic
# ---------------------------------------------------------------------------


class TestProtectSheetBasic:
    def test_protect_sheet_basic(self, tmp_path: Path) -> None:
        """protect_sheet → verify sheet protection with openpyxl."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        result = protect_sheet(fp, "Sheet1")
        assert "protected" in result.lower()

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws.protection.sheet is True
        wb.close()


# ---------------------------------------------------------------------------
# 2. Protect sheet with password
# ---------------------------------------------------------------------------


class TestProtectSheetWithPassword:
    def test_protect_sheet_with_password(self, tmp_path: Path) -> None:
        """Protect with password → verify protection is set."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        protect_sheet(fp, "Sheet1", password="secret123")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws.protection.sheet is True
        # Password should be hashed — not None
        assert ws.protection.password is not None
        wb.close()


# ---------------------------------------------------------------------------
# 3. Unprotect sheet
# ---------------------------------------------------------------------------


class TestUnprotectSheet:
    def test_unprotect_sheet(self, tmp_path: Path) -> None:
        """protect → unprotect → verify unprotected."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        protect_sheet(fp, "Sheet1")
        unprotect_sheet(fp, "Sheet1")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws.protection.sheet is False
        wb.close()


# ---------------------------------------------------------------------------
# 4. Protect cells range
# ---------------------------------------------------------------------------


class TestProtectCellsRange:
    def test_protect_cells_range(self, tmp_path: Path) -> None:
        """protect_cells with unlocked_ranges → verify lock attributes."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        protect_cells(
            fp,
            "Sheet1",
            locked_range="A1:D4",
            unlocked_ranges=["B2:B4"],
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        # A1 should be locked
        assert ws["A1"].protection.locked is True
        # B2 should be unlocked
        assert ws["B2"].protection.locked is False
        assert ws["B3"].protection.locked is False
        wb.close()


# ---------------------------------------------------------------------------
# 5. Protect cells single cell (BUG-17 regression)
# ---------------------------------------------------------------------------


class TestProtectCellsSingleCell:
    def test_protect_cells_single_cell(self, tmp_path: Path) -> None:
        """Unlock single cell (e.g., 'B2') → verify it works (BUG-17)."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        protect_cells(
            fp,
            "Sheet1",
            locked_range="A1:D4",
            unlocked_ranges=["B2"],
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        # B2 unlocked, A1 locked
        assert ws["B2"].protection.locked is False
        assert ws["A1"].protection.locked is True
        # Adjacent cell B3 should still be locked
        assert ws["B3"].protection.locked is True
        wb.close()


# ---------------------------------------------------------------------------
# 6. Protect workbook
# ---------------------------------------------------------------------------


class TestProtectWorkbook:
    def test_protect_workbook(self, tmp_path: Path) -> None:
        """protect_workbook → verify with openpyxl."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        result = protect_workbook(fp, lock_structure=True, lock_windows=True)
        assert "protected" in result.lower()

        wb = openpyxl.load_workbook(fp)
        assert wb.security.lockStructure is True
        assert wb.security.lockWindows is True
        wb.close()


# ---------------------------------------------------------------------------
# 7. Unprotect workbook
# ---------------------------------------------------------------------------


class TestUnprotectWorkbook:
    def test_unprotect_workbook(self, tmp_path: Path) -> None:
        """protect → unprotect → verify unprotected."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        protect_workbook(fp, lock_structure=True)
        unprotect_workbook(fp)

        wb = openpyxl.load_workbook(fp)
        assert wb.security.lockStructure is False
        assert wb.security.lockWindows is False
        wb.close()


# ---------------------------------------------------------------------------
# 8. Add dropdown validation
# ---------------------------------------------------------------------------


class TestAddDropdownValidation:
    def test_add_dropdown_validation(self, tmp_path: Path) -> None:
        """add_dropdown_validation → verify data validation with openpyxl."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        add_dropdown_validation(
            fp,
            "Sheet1",
            cell_range="C2:C4",
            options=["New York", "Chicago", "Boston"],
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        validations = ws.data_validations.dataValidation
        assert len(validations) == 1
        dv = validations[0]
        assert dv.type == "list"
        assert "C2:C4" in str(dv.sqref)
        wb.close()


# ---------------------------------------------------------------------------
# 9. Add numeric validation
# ---------------------------------------------------------------------------


class TestAddNumericValidation:
    def test_add_numeric_validation(self, tmp_path: Path) -> None:
        """add_numeric_validation → verify validation rule."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        add_numeric_validation(
            fp,
            "Sheet1",
            cell_range="D2:D4",
            operator="between",
            value1=0,
            value2=100000,
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        validations = ws.data_validations.dataValidation
        assert len(validations) == 1
        dv = validations[0]
        assert dv.type in ("whole", "decimal")
        assert dv.operator == "between"
        assert "D2:D4" in str(dv.sqref)
        wb.close()


# ---------------------------------------------------------------------------
# 10. Add date validation
# ---------------------------------------------------------------------------


class TestAddDateValidation:
    def test_add_date_validation(self, tmp_path: Path) -> None:
        """add_date_validation → verify validation rule."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        add_date_validation(
            fp,
            "Sheet1",
            cell_range="E2:E4",
            operator="greaterThan",
            date1="2020-01-01",
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        validations = ws.data_validations.dataValidation
        assert len(validations) == 1
        dv = validations[0]
        assert dv.type == "date"
        assert dv.operator == "greaterThan"
        wb.close()


# ---------------------------------------------------------------------------
# 11. Add formula validation
# ---------------------------------------------------------------------------


class TestAddFormulaValidation:
    def test_add_formula_validation(self, tmp_path: Path) -> None:
        """add_formula_validation → verify custom validation."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        result = add_formula_validation(
            fp,
            "Sheet1",
            range_str="B2:B4",
            formula="=B2>0",
        )

        assert result["validation_type"] == "custom"
        assert result["formula"] == "=B2>0"

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        validations = ws.data_validations.dataValidation
        assert len(validations) == 1
        dv = validations[0]
        assert dv.type == "custom"
        wb.close()


# ---------------------------------------------------------------------------
# 12. Remove validation
# ---------------------------------------------------------------------------


class TestRemoveValidation:
    def test_remove_validation(self, tmp_path: Path) -> None:
        """Add → remove_validation → verify removed."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        add_dropdown_validation(
            fp,
            "Sheet1",
            cell_range="C2:C4",
            options=["A", "B", "C"],
        )

        # Verify it was added
        wb = openpyxl.load_workbook(fp)
        assert len(wb["Sheet1"].data_validations.dataValidation) == 1
        wb.close()

        # Remove it
        remove_validation(fp, "Sheet1", cell_range="C2:C4")

        wb = openpyxl.load_workbook(fp)
        assert len(wb["Sheet1"].data_validations.dataValidation) == 0
        wb.close()


# ---------------------------------------------------------------------------
# 13. Combined protection workflow
# ---------------------------------------------------------------------------


class TestProtectionWorkflow:
    def test_protection_workflow(self, tmp_path: Path) -> None:
        """protect sheet → protect cells → add validation → verify all together."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        # Step 1: protect cells — lock all, unlock C2:C4
        protect_cells(fp, "Sheet1", locked_range="A1:D4", unlocked_ranges=["C2:C4"])

        # Step 2: add dropdown to the unlocked range
        add_dropdown_validation(
            fp,
            "Sheet1",
            cell_range="C2:C4",
            options=["New York", "Chicago", "Boston"],
        )

        # Step 3: protect the sheet
        protect_sheet(fp, "Sheet1", password="pass123")

        # Verify everything
        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]

        # Sheet protection enabled
        assert ws.protection.sheet is True

        # A1 is locked, C2 is unlocked
        assert ws["A1"].protection.locked is True
        assert ws["C2"].protection.locked is False

        # Dropdown validation present
        validations = ws.data_validations.dataValidation
        assert len(validations) == 1
        assert validations[0].type == "list"
        wb.close()


# ---------------------------------------------------------------------------
# 14. Protect cells with mixed ranges
# ---------------------------------------------------------------------------


class TestProtectCellsMixedRanges:
    def test_protect_cells_mixed_ranges(self, tmp_path: Path) -> None:
        """Mix of single cells and ranges → verify all unlocked correctly."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        protect_cells(
            fp,
            "Sheet1",
            locked_range="A1:D4",
            unlocked_ranges=["B2", "C3:C4", "D2"],
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]

        # Locked cells
        assert ws["A1"].protection.locked is True
        assert ws["A2"].protection.locked is True
        assert ws["D3"].protection.locked is True

        # Unlocked single cells
        assert ws["B2"].protection.locked is False
        assert ws["D2"].protection.locked is False

        # Unlocked range
        assert ws["C3"].protection.locked is False
        assert ws["C4"].protection.locked is False

        # Adjacent cells should remain locked
        assert ws["B3"].protection.locked is True
        assert ws["C2"].protection.locked is True
        wb.close()


# ---------------------------------------------------------------------------
# 15. Protect sheet with permissions
# ---------------------------------------------------------------------------


class TestProtectSheetWithPermissions:
    def test_protect_sheet_with_permissions(self, tmp_path: Path) -> None:
        """Protect sheet with specific permissions → verify flags."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        protect_sheet(
            fp,
            "Sheet1",
            allow_sort=True,
            allow_filter=True,
            allow_formatting_cells=True,
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws.protection.sheet is True
        assert ws.protection.sort is True
        assert ws.protection.autoFilter is True
        assert ws.protection.formatCells is True
        # Defaults should be False
        assert ws.protection.insertRows is False
        assert ws.protection.deleteRows is False
        wb.close()


# ---------------------------------------------------------------------------
# 16. Protect workbook with password
# ---------------------------------------------------------------------------


class TestProtectWorkbookWithPassword:
    def test_protect_workbook_with_password(self, tmp_path: Path) -> None:
        """protect_workbook with password → verify password is set."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        protect_workbook(fp, password="wb_pass", lock_structure=True)

        wb = openpyxl.load_workbook(fp)
        assert wb.security.lockStructure is True
        # Password should be hashed
        assert wb.security.workbookPassword is not None
        wb.close()


# ---------------------------------------------------------------------------
# 17. Numeric validation with between operator
# ---------------------------------------------------------------------------


class TestNumericValidationBetween:
    def test_numeric_validation_between(self, tmp_path: Path) -> None:
        """Between operator requires value2 → verify both bounds set."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        add_numeric_validation(
            fp,
            "Sheet1",
            cell_range="B2:B4",
            operator="between",
            value1=18,
            value2=65,
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        dv = ws.data_validations.dataValidation[0]
        assert dv.formula1 == "18"
        assert dv.formula2 == "65"
        wb.close()


# ---------------------------------------------------------------------------
# 18. Date validation between dates
# ---------------------------------------------------------------------------


class TestDateValidationBetween:
    def test_date_validation_between(self, tmp_path: Path) -> None:
        """Date validation with between operator → verify both dates."""
        fp = str(tmp_path / "test.xlsx")
        _create_data_file(fp)

        add_date_validation(
            fp,
            "Sheet1",
            cell_range="E2:E4",
            operator="between",
            date1="2020-01-01",
            date2="2025-12-31",
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        dv = ws.data_validations.dataValidation[0]
        assert dv.type == "date"
        assert dv.operator == "between"
        assert dv.formula1 == "2020-01-01"
        assert dv.formula2 == "2025-12-31"
        wb.close()
