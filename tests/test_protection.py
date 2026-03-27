from __future__ import annotations

from openpyxl import load_workbook

from mcp_server.tools.doc_properties import protect_workbook, unprotect_workbook
from mcp_server.tools.protection import protect_cells, protect_sheet, unprotect_sheet

# ── protect_sheet ───────────────────────────────────────────────


def test_protect_sheet_basic(sample_xlsx: str) -> None:
    result = protect_sheet(sample_xlsx, "Sheet1")
    assert "protected" in result.lower()

    wb = load_workbook(sample_xlsx)
    assert wb["Sheet1"].protection.sheet is True
    wb.close()


def test_protect_sheet_with_password(sample_xlsx: str) -> None:
    result = protect_sheet(sample_xlsx, "Sheet1", password="secret")
    assert "protected" in result.lower()

    wb = load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws.protection.sheet is True
    assert ws.protection.password is not None
    wb.close()


def test_protect_sheet_with_permissions(sample_xlsx: str) -> None:
    protect_sheet(
        sample_xlsx,
        "Sheet1",
        allow_sort=True,
        allow_filter=True,
        allow_insert_rows=True,
    )
    wb = load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws.protection.sheet is True
    assert ws.protection.sort is True
    assert ws.protection.autoFilter is True
    assert ws.protection.insertRows is True
    assert ws.protection.insertColumns is False
    wb.close()


def test_protect_sheet_all_permissions(sample_xlsx: str) -> None:
    protect_sheet(
        sample_xlsx,
        "Sheet1",
        allow_formatting_cells=True,
        allow_formatting_columns=True,
        allow_formatting_rows=True,
        allow_insert_columns=True,
        allow_insert_rows=True,
        allow_delete_columns=True,
        allow_delete_rows=True,
        allow_sort=True,
        allow_filter=True,
    )
    wb = load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws.protection.formatCells is True
    assert ws.protection.formatColumns is True
    assert ws.protection.formatRows is True
    assert ws.protection.insertColumns is True
    assert ws.protection.insertRows is True
    assert ws.protection.deleteColumns is True
    assert ws.protection.deleteRows is True
    assert ws.protection.sort is True
    assert ws.protection.autoFilter is True
    wb.close()


# ── unprotect_sheet ─────────────────────────────────────────────


def test_unprotect_sheet(sample_xlsx: str) -> None:
    protect_sheet(sample_xlsx, "Sheet1", password="secret")
    result = unprotect_sheet(sample_xlsx, "Sheet1")
    assert "removed" in result.lower()

    wb = load_workbook(sample_xlsx)
    assert wb["Sheet1"].protection.sheet is False
    wb.close()


def test_unprotect_sheet_without_password(sample_xlsx: str) -> None:
    """Unprotecting without providing the password still clears protection."""
    protect_sheet(sample_xlsx, "Sheet1", password="secret")
    result = unprotect_sheet(sample_xlsx, "Sheet1")
    assert "removed" in result.lower()

    wb = load_workbook(sample_xlsx)
    assert wb["Sheet1"].protection.sheet is False
    wb.close()


def test_unprotect_sheet_not_protected(sample_xlsx: str) -> None:
    """Unprotecting an unprotected sheet is a no-op, not an error."""
    result = unprotect_sheet(sample_xlsx, "Sheet1")
    assert "removed" in result.lower()


# ── protect_workbook ────────────────────────────────────────────


def test_protect_workbook_basic(sample_xlsx: str) -> None:
    result = protect_workbook(sample_xlsx)
    assert "protected" in result.lower()

    wb = load_workbook(sample_xlsx)
    assert wb.security.lockStructure is True
    wb.close()


def test_protect_workbook_with_password(sample_xlsx: str) -> None:
    result = protect_workbook(sample_xlsx, password="wbpass")
    assert "protected" in result.lower()

    wb = load_workbook(sample_xlsx)
    assert wb.security.lockStructure is True
    wb.close()


def test_protect_workbook_lock_windows(sample_xlsx: str) -> None:
    protect_workbook(sample_xlsx, lock_structure=False, lock_windows=True)
    wb = load_workbook(sample_xlsx)
    assert wb.security.lockStructure is False
    assert wb.security.lockWindows is True
    wb.close()


def test_protect_workbook_both_locks(sample_xlsx: str) -> None:
    protect_workbook(sample_xlsx, lock_structure=True, lock_windows=True, password="p")
    wb = load_workbook(sample_xlsx)
    assert wb.security.lockStructure is True
    assert wb.security.lockWindows is True
    wb.close()


# ── unprotect_workbook ──────────────────────────────────────────


def test_unprotect_workbook(sample_xlsx: str) -> None:
    protect_workbook(sample_xlsx, password="wbpass")
    result = unprotect_workbook(sample_xlsx)
    assert "removed" in result.lower()

    wb = load_workbook(sample_xlsx)
    assert wb.security.lockStructure is False
    assert wb.security.lockWindows is False
    wb.close()


def test_unprotect_workbook_not_protected(sample_xlsx: str) -> None:
    """Unprotecting an unprotected workbook is a no-op."""
    result = unprotect_workbook(sample_xlsx)
    assert "removed" in result.lower()


# ── protect_cells ───────────────────────────────────────────────


def test_protect_cells_locked_range(sample_xlsx: str) -> None:
    result = protect_cells(sample_xlsx, "Sheet1", locked_range="A1:D1")
    assert "Locked" in result

    wb = load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws["A1"].protection.locked is True
    assert ws["D1"].protection.locked is True
    wb.close()


def test_protect_cells_with_unlocked_ranges(sample_xlsx: str) -> None:
    result = protect_cells(
        sample_xlsx,
        "Sheet1",
        locked_range="A1:D6",
        unlocked_ranges=["B2:B6"],
    )
    assert "Locked" in result
    assert "unlocked" in result.lower()

    wb = load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    # Locked range cell
    assert ws["A1"].protection.locked is True
    # Unlocked range cell
    assert ws["B2"].protection.locked is False
    assert ws["B6"].protection.locked is False
    wb.close()


def test_protect_cells_multiple_unlocked_ranges(sample_xlsx: str) -> None:
    protect_cells(
        sample_xlsx,
        "Sheet1",
        locked_range="A1:D6",
        unlocked_ranges=["B2:B6", "D2:D6"],
    )
    wb = load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws["A1"].protection.locked is True
    assert ws["B3"].protection.locked is False
    assert ws["D4"].protection.locked is False
    wb.close()


def test_protect_cells_no_unlocked(sample_xlsx: str) -> None:
    result = protect_cells(sample_xlsx, "Sheet1", locked_range="A1:B2")
    assert "unlocked" not in result.lower()


def test_protect_cells_enable_sheet_protection_message(sample_xlsx: str) -> None:
    """Return message should mention enabling sheet protection."""
    result = protect_cells(sample_xlsx, "Sheet1", locked_range="A1:A1")
    assert "sheet protection" in result.lower()


# ── lifecycle / integration ─────────────────────────────────────


def test_protect_then_unprotect_sheet_roundtrip(sample_xlsx: str) -> None:
    protect_sheet(sample_xlsx, "Sheet1", password="abc")
    wb = load_workbook(sample_xlsx)
    assert wb["Sheet1"].protection.sheet is True
    wb.close()

    unprotect_sheet(sample_xlsx, "Sheet1")
    wb = load_workbook(sample_xlsx)
    assert wb["Sheet1"].protection.sheet is False
    wb.close()


def test_protect_cells_and_sheet_together(sample_xlsx: str) -> None:
    """Cells + sheet protection work in sequence."""
    protect_cells(sample_xlsx, "Sheet1", locked_range="A1:D1", unlocked_ranges=["B1:B1"])
    protect_sheet(sample_xlsx, "Sheet1")

    wb = load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert ws.protection.sheet is True
    assert ws["A1"].protection.locked is True
    assert ws["B1"].protection.locked is False
    wb.close()


# ── additional coverage ────────────────────────────────────────────────────────


def test_unprotect_sheet_wrong_password_still_succeeds(sample_xlsx: str) -> None:
    """openpyxl does not validate passwords on read, so passing the wrong
    password to unprotect_sheet still clears the protection flag."""
    protect_sheet(sample_xlsx, "Sheet1", password="correct_password")
    result = unprotect_sheet(sample_xlsx, "Sheet1", password="wrong_password")
    assert "removed" in result.lower()
    wb = load_workbook(sample_xlsx)
    assert wb["Sheet1"].protection.sheet is False
    wb.close()


def test_protect_sheet_return_message_contains_sheet_name(sample_xlsx: str) -> None:
    """protect_sheet return message contains the sheet name."""
    result = protect_sheet(sample_xlsx, "Sheet1")
    assert "Sheet1" in result


def test_protect_cells_single_cell(sample_xlsx: str) -> None:
    """protect_cells can lock a single-cell range."""
    result = protect_cells(sample_xlsx, "Sheet1", locked_range="C3:C3")
    assert "Locked" in result
    wb = load_workbook(sample_xlsx)
    assert wb["Sheet1"]["C3"].protection.locked is True
    wb.close()


def test_protect_cells_return_message_mentions_sheet(sample_xlsx: str) -> None:
    """protect_cells return message includes the sheet name."""
    result = protect_cells(sample_xlsx, "Sheet1", locked_range="A1:B2")
    assert "Sheet1" in result
