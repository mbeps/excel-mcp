"""Tests verifying atomic workbook saving and concurrent access resilience.

Guarantees:
1. Concurrent saves to the same workbook do not corrupt the ZIP archive (BadZipFile / Bad CRC-32).
2. Save failure leaves original target file intact on disk (no truncation or corruption).
3. read_sheet_df wraps corrupt or unreadable files in ValueError.
"""

from __future__ import annotations

import threading
from pathlib import Path

import openpyxl
import pytest

from mcp_server.tools.formatting import format_cells
from mcp_server.utils.excel_helpers import (
    load_workbook_safe,
    read_sheet_df,
    save_workbook_safe,
)


def _create_sample_file(path: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Car Details"
    for r in range(1, 30):
        for c in range(1, 10):
            ws.cell(row=r, column=c, value=f"Val_{r}_{c}")
    wb.save(path)
    wb.close()


def test_concurrent_formatting_no_corruption(tmp_path: Path) -> None:
    """Simulate rapid/concurrent format_cells calls on the same file without BadZipFile corruption."""
    fp = str(tmp_path / "car_comparison.xlsx")
    _create_sample_file(fp)

    ranges = ["A1:E1", "A2:A20", "A1:E20", "B2:D15", "A10:E25"]
    exceptions: list[Exception] = []

    def run_format(cell_range: str) -> None:
        try:
            format_cells(
                file_path=fp,
                sheet_name="Car Details",
                cell_range=cell_range,
                bold=True,
            )
        except Exception as e:
            exceptions.append(e)

    threads = [threading.Thread(target=run_format, args=(rng,)) for rng in ranges]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    # Verify no unhandled thread crashes occurred
    assert not exceptions, f"Exceptions occurred during formatting: {exceptions}"

    # Verify workbook is not corrupt and can be read cleanly
    wb = load_workbook_safe(fp)
    assert "Car Details" in wb.sheetnames
    wb.close()

    # Verify read_sheet_df reads it cleanly without BadZipFile
    df = read_sheet_df(fp, "Car Details")
    assert not df.empty


def test_save_workbook_failure_preserves_original(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """If wb.save fails midway, original workbook on disk must remain untouched."""
    fp = str(tmp_path / "original.xlsx")
    _create_sample_file(fp)

    original_size = Path(fp).stat().st_size
    wb = load_workbook_safe(fp)

    def mock_save_fail(self: openpyxl.Workbook, filename: str) -> None:
        # Simulate partial write then failure
        with open(filename, "wb") as f:
            f.write(b"CORRUPTED_PARTIAL_BYTES")
        raise IOError("Disk full simulation")

    monkeypatch.setattr(openpyxl.Workbook, "save", mock_save_fail)

    with pytest.raises(ValueError, match="Failed to save workbook"):
        save_workbook_safe(wb, fp)

    # Original file must still be intact with original size and loadable
    assert Path(fp).stat().st_size == original_size
    wb_orig = openpyxl.load_workbook(fp)
    assert "Car Details" in wb_orig.sheetnames
    wb_orig.close()


def test_read_sheet_df_wraps_corrupt_file_in_value_error(tmp_path: Path) -> None:
    """Corrupt file read by read_sheet_df raises ValueError instead of unhandled BadZipFile."""
    fp = str(tmp_path / "corrupt.xlsx")
    with open(fp, "wb") as f:
        f.write(b"PK\x03\x04not a valid zip content")

    with pytest.raises(ValueError, match="Failed to read sheet"):
        read_sheet_df(fp, "Sheet1")
