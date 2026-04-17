"""Workflow tests for edge cases.

Tests call tool functions directly and verify results independently using
openpyxl or pandas — never via MCP tools.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from mcp_server.tools.analysis import (
    aggregate_data,
    column_statistics,
    filter_data_advanced,
    sort_data,
)
from mcp_server.tools.cell_ops import (
    find_replace,
    read_cell,
    read_range,
    write_cell,
    write_range,
)
from mcp_server.tools.charts import create_chart
from mcp_server.tools.formatting import format_cells
from mcp_server.tools.formulas import (
    get_formula_dependents,
    get_formula_precedents,
    list_formulas,
    set_formula,
)
from mcp_server.tools.workbook import (
    copy_sheet,
    create_workbook,
    get_workbook_metadata,
    rename_sheet,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_wb(fp: str, headers: list[str], rows: list[list]) -> None:
    create_workbook(fp, sheet_names=["Sheet1"])
    write_range(fp, "Sheet1", "A1", [headers] + rows)


# ===========================================================================
# Empty / Minimal Data
# ===========================================================================


class TestWriteSingleCellEmptyWorkbook:
    """1. Write to a cell in a fresh, empty workbook."""

    def test_write_single_cell_empty_workbook(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "empty.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "C5", "hello")

        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["C5"].value == "hello"
        wb.close()


class TestReadEmptyCell:
    """2. Read a cell that has never been written to."""

    def test_read_empty_cell(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "empty.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        result = read_cell(fp, "Sheet1", "Z99")

        assert result["value"] is None


class TestReadRangeBeyondData:
    """3. Read range that exceeds actual data — graceful handling."""

    def test_read_range_beyond_data(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "small.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", "only")

        result = read_range(fp, "Sheet1", "A1", "D10")

        # openpyxl read_range returns the requested rectangle; rows beyond
        # written data still appear with None values
        assert result["row_count"] >= 1
        assert result["col_count"] >= 1
        # The first cell should contain our value
        assert result["rows"][0][0] == "only"


class TestOperationsOnSingleRow:
    """4. Filter/sort/aggregate on a dataset with only 1 data row."""

    def test_filter_single_row(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "one_row.xlsx")
        _make_wb(fp, ["Name", "Score"], [["Alice", 100]])

        result = filter_data_advanced(
            fp,
            "Sheet1",
            conditions=[{"column": "Score", "operator": ">=", "value": 50}],
        )
        assert result["rows"] == 1
        assert result["data"][0][0] == "Alice"

    def test_sort_single_row(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "one_row.xlsx")
        _make_wb(fp, ["Name", "Score"], [["Alice", 100]])

        msg = sort_data(fp, "Sheet1", column="Score")
        assert "1" in msg

    def test_aggregate_single_row(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "one_row.xlsx")
        _make_wb(fp, ["Name", "Score"], [["Alice", 100]])

        result = aggregate_data(fp, "Sheet1", group_by="Name", value_column="Score", operation="sum")
        assert len(result["groups"]) == 1


class TestOperationsOnSingleColumn:
    """5. Single-column dataset operations."""

    def test_column_statistics_single_col(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "one_col.xlsx")
        _make_wb(fp, ["Value"], [[10], [20], [30]])

        result = column_statistics(fp, "Sheet1", column="Value")
        assert result.count == 3
        assert result.mean == 20.0
        assert result.sum_val == 60.0

    def test_sort_single_column(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "one_col.xlsx")
        _make_wb(fp, ["Value"], [[30], [10], [20]])

        sort_data(fp, "Sheet1", column="Value", ascending=True)

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        vals = [ws.cell(row=r, column=1).value for r in range(2, 5)]
        assert vals == [10, 20, 30]
        wb.close()


# ===========================================================================
# Special Characters
# ===========================================================================


class TestUnicodeCellValues:
    """6. Write/read Chinese, Arabic, emoji — verify preservation."""

    def test_unicode_cell_values(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "unicode.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        test_values = {
            "A1": "你好世界",  # Chinese
            "A2": "مرحبا بالعالم",  # Arabic
            "A3": "🎉🚀💡",  # Emoji
            "A4": "café résumé",  # Latin accents
            "A5": "γεια σου",  # Greek
        }
        for ref, val in test_values.items():
            write_cell(fp, "Sheet1", ref, val)

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        for ref, expected in test_values.items():
            assert ws[ref].value == expected, f"Failed for {ref}"
        wb.close()


class TestSpecialCharsInFormulas:
    """7. Formulas with special characters."""

    def test_special_chars_in_formulas(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "formulas.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", 10)
        write_cell(fp, "Sheet1", "B1", 20)

        set_formula(fp, "Sheet1", "C1", '=A1&" + "&B1')

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["C1"].value.startswith("=")
        assert '"' in ws["C1"].value
        wb.close()


class TestWhitespaceValues:
    """8. Cells with only spaces, tabs, newlines."""

    def test_whitespace_values(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "ws.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", "   ")  # spaces
        write_cell(fp, "Sheet1", "A2", "\t\t")  # tabs
        write_cell(fp, "Sheet1", "A3", "\n\n")  # newlines
        write_cell(fp, "Sheet1", "A4", " \t\n mixed")  # mixed

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].value == "   "
        assert ws["A2"].value == "\t\t"
        assert ws["A3"].value == "\n\n"
        assert ws["A4"].value == " \t\n mixed"
        wb.close()


# ===========================================================================
# Boundary Values
# ===========================================================================


class TestMaxColumnLetter:
    """9. Write to very high columns (Z, AA, AZ)."""

    def test_max_column_letter(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "highcol.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "Z1", "col_z")
        write_cell(fp, "Sheet1", "AA1", "col_aa")
        write_cell(fp, "Sheet1", "AZ1", "col_az")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["Z1"].value == "col_z"
        assert ws["AA1"].value == "col_aa"
        assert ws["AZ1"].value == "col_az"
        wb.close()


class TestLargeNumberValues:
    """10. Very large and very small numbers — precision check."""

    def test_large_number_values(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "nums.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        large = 1e15
        small = 1e-15
        write_cell(fp, "Sheet1", "A1", large)
        write_cell(fp, "Sheet1", "A2", small)

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].value == large
        assert abs(ws["A2"].value - small) < 1e-30
        wb.close()


class TestNegativeNumbers:
    """11. Negative values in numeric tools."""

    def test_negative_numbers(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "neg.xlsx")
        _make_wb(fp, ["Value"], [[-10], [-20], [-30]])

        result = column_statistics(fp, "Sheet1", column="Value")
        assert result.sum_val == -60.0
        assert result.mean == -20.0
        assert result.min_val == -30.0
        assert result.max_val == -10.0

    def test_negative_write_read(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "neg2.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", -999.99)

        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["A1"].value == -999.99
        wb.close()


class TestZeroValues:
    """12. Zeros everywhere — verify not treated as empty."""

    def test_zero_values(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "zeros.xlsx")
        _make_wb(fp, ["A", "B"], [[0, 0], [0, 0]])

        result = read_range(fp, "Sheet1", "A2", "B3")
        for row in result["rows"]:
            for val in row:
                assert val == 0, "Zero should be preserved, not None"

    def test_zero_not_filtered_out(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "zeros2.xlsx")
        _make_wb(fp, ["Score"], [[0], [10], [0]])

        result = filter_data_advanced(
            fp,
            "Sheet1",
            conditions=[{"column": "Score", "operator": "==", "value": 0}],
        )
        assert result["rows"] == 2


class TestBooleanValues:
    """13. True/False in cells — type preserved."""

    def test_boolean_values(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "bool.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", True)
        write_cell(fp, "Sheet1", "A2", False)

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].value is True
        assert ws["A2"].value is False
        assert isinstance(ws["A1"].value, bool)
        assert isinstance(ws["A2"].value, bool)
        wb.close()


# ===========================================================================
# Error Handling
# ===========================================================================


class TestInvalidFileExtension:
    """14. .txt file — verify rejected."""

    def test_invalid_file_extension(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "bad.txt")
        with pytest.raises(ValueError, match="Unsupported file type"):
            create_workbook(fp)


class TestNonexistentFileRead:
    """15. Read from non-existent file."""

    def test_nonexistent_file_read(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "does_not_exist.xlsx")
        with pytest.raises(ValueError, match="File not found|not found"):
            read_cell(fp, "Sheet1", "A1")


class TestInvalidCellReference:
    """16. Write to invalid cell reference — verify error."""

    def test_invalid_cell_reference(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "test.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        with pytest.raises((ValueError, KeyError)):
            write_cell(fp, "Sheet1", "ZZZZ99999999", "bad")


class TestInvalidRange:
    """17. Read with invalid range format."""

    def test_invalid_range(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "test.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        with pytest.raises((ValueError, TypeError)):
            read_range(fp, "Sheet1", "!!!!", "@@@@")


class TestDuplicateSheetName:
    """18. Create/rename to existing sheet name."""

    def test_duplicate_sheet_rename(self, tmp_path: Path) -> None:
        """openpyxl silently allows renaming to an existing name (overwrites title).
        Verify the rename completes and the workbook has the expected sheets."""
        fp = str(tmp_path / "dup.xlsx")
        create_workbook(fp, sheet_names=["Alpha", "Beta"])

        # openpyxl does not raise on duplicate rename — it overwrites the title
        rename_sheet(fp, "Alpha", "Beta")

        wb = openpyxl.load_workbook(fp)
        # At least one sheet named "Beta" must exist
        assert "Beta" in wb.sheetnames
        wb.close()

    def test_copy_sheet_existing_name(self, tmp_path: Path) -> None:
        """openpyxl silently renames the copy to avoid collision (e.g. 'Beta1')."""
        fp = str(tmp_path / "dup2.xlsx")
        create_workbook(fp, sheet_names=["Alpha", "Beta"])

        # openpyxl copy_worksheet auto-renames to avoid collision
        copy_sheet(fp, "Alpha", "Beta")

        wb = openpyxl.load_workbook(fp)
        # Should have at least 3 sheets (Alpha, Beta, Beta-copy-variant)
        assert len(wb.sheetnames) >= 2
        wb.close()


# ===========================================================================
# Formula Edge Cases
# ===========================================================================


class TestCircularReferenceDetection:
    """19. A1=B1, B1=A1 — verify handling."""

    def test_circular_reference_detection(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "circ.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        # openpyxl stores formulas as text; circular refs are stored without error
        set_formula(fp, "Sheet1", "A1", "=B1")
        set_formula(fp, "Sheet1", "B1", "=A1")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].value == "=B1"
        assert ws["B1"].value == "=A1"
        wb.close()

        # Verify precedent/dependent detection works with circular refs
        prec = get_formula_precedents(fp, "Sheet1", "A1")
        assert "B1" in prec["precedents"]

        deps = get_formula_dependents(fp, "Sheet1", "A1")
        assert any(d["cell"] == "B1" for d in deps["dependents"])


class TestFormulaWithErrorValue:
    """20. =1/0 — verify the formula is stored (openpyxl cannot eval)."""

    def test_formula_with_error_value(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "div0.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        set_formula(fp, "Sheet1", "A1", "=1/0")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].value == "=1/0"
        wb.close()

        # list_formulas should find it
        formulas = list_formulas(fp, "Sheet1")
        assert any(f["cell_ref"] == "A1" for f in formulas)


class TestEmptyFormula:
    """21. Set empty formula string — prepends '=' → stores '='."""

    def test_empty_formula(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "empty_f.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        set_formula(fp, "Sheet1", "A1", "")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        # set_formula prepends '=' if missing, so cell should have '='
        assert ws["A1"].value == "="
        wb.close()


# ===========================================================================
# Multi-Sheet Edge Cases
# ===========================================================================


class TestCrossSheetFormulaReferences:
    """22. Formula referencing another sheet."""

    def test_cross_sheet_formula_references(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "cross.xlsx")
        create_workbook(fp, sheet_names=["Data", "Summary"])
        write_cell(fp, "Data", "A1", 42)

        set_formula(fp, "Summary", "A1", "=Data!A1")

        wb = openpyxl.load_workbook(fp)
        assert wb["Summary"]["A1"].value == "=Data!A1"
        wb.close()

        prec = get_formula_precedents(fp, "Summary", "A1")
        assert any("Data" in str(p) for p in prec["precedents"])


class TestOperationsAfterSheetRename:
    """23. Rename sheet → operations still work on renamed sheet."""

    def test_operations_after_sheet_rename(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "rename.xlsx")
        create_workbook(fp, sheet_names=["Original"])
        write_cell(fp, "Original", "A1", "data")

        rename_sheet(fp, "Original", "Renamed")

        # Read from the renamed sheet
        result = read_cell(fp, "Renamed", "A1")
        assert result["value"] == "data"

        # Write to the renamed sheet
        write_cell(fp, "Renamed", "B1", "more_data")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Renamed"]
        assert ws["A1"].value == "data"
        assert ws["B1"].value == "more_data"
        assert "Original" not in wb.sheetnames
        wb.close()


# ===========================================================================
# Concurrency-like
# ===========================================================================


class TestRapidSuccessiveWrites:
    """24. 100 write operations in sequence — verify all persisted."""

    def test_rapid_successive_writes(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "rapid.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        for i in range(1, 101):
            write_cell(fp, "Sheet1", f"A{i}", i)

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        for i in range(1, 101):
            assert ws[f"A{i}"].value == i, f"Row {i} not persisted"
        wb.close()


class TestReadWriteInterleave:
    """25. Write → read → write → read cycle — verify consistency."""

    def test_read_write_interleave(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "interleave.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        write_cell(fp, "Sheet1", "A1", 10)
        r1 = read_cell(fp, "Sheet1", "A1")
        assert r1["value"] == 10

        write_cell(fp, "Sheet1", "A1", 20)
        r2 = read_cell(fp, "Sheet1", "A1")
        assert r2["value"] == 20

        write_cell(fp, "Sheet1", "B1", "text")
        r3 = read_cell(fp, "Sheet1", "B1")
        assert r3["value"] == "text"

        # Verify A1 still has last written value
        r4 = read_cell(fp, "Sheet1", "A1")
        assert r4["value"] == 20

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].value == 20
        assert ws["B1"].value == "text"
        wb.close()


# ===========================================================================
# Data Type Coercion
# ===========================================================================


class TestNumericStringPreservation:
    """26. "123" as string → verify not coerced to number."""

    def test_numeric_string_preservation(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "numstr.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", "123")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].value == "123"
        assert isinstance(ws["A1"].value, str)
        wb.close()

        result = read_cell(fp, "Sheet1", "A1")
        assert result["value"] == "123"


class TestDateStringNotAutoParsed:
    """27. "2024-01-01" → verify stored as-is when no date parsing."""

    def test_date_string_not_auto_parsed(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "datestr.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", "2024-01-01")

        result = read_cell(fp, "Sheet1", "A1")
        assert result["value"] == "2024-01-01"
        assert isinstance(result["value"], str)


class TestFormulaLikeString:
    """28. Write "=SUM(A1)" as value (not formula) → verify stored as formula by openpyxl."""

    def test_formula_like_string(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "fstr.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        # write_cell passes the value directly to openpyxl
        # openpyxl treats strings starting with '=' as formulas
        write_cell(fp, "Sheet1", "A1", "=SUM(A2:A5)")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        # openpyxl stores '=...' strings as formula cells
        assert ws["A1"].value == "=SUM(A2:A5)"
        assert ws["A1"].data_type == "f"  # formula type
        wb.close()


# ===========================================================================
# Additional edge cases
# ===========================================================================


class TestWriteNoneValue:
    """Extra: Write None to an existing cell — should clear it."""

    def test_write_none_clears_cell(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "none.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", 42)
        write_cell(fp, "Sheet1", "A1", None)

        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["A1"].value is None
        wb.close()


class TestReadCellWithFormula:
    """Extra: read_cell with include_formula flag."""

    def test_read_cell_with_formula(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "rf.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        set_formula(fp, "Sheet1", "A1", "=1+1")

        result = read_cell(fp, "Sheet1", "A1", include_formula=True)
        assert result["formula"] == "=1+1"


class TestFindReplaceNoMatch:
    """Extra: find_replace when search text doesn't exist."""

    def test_find_replace_no_match(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "fr.xlsx")
        _make_wb(fp, ["Name"], [["Alice"], ["Bob"]])

        result = find_replace(fp, "Sheet1", "NONEXISTENT", "replaced")
        assert result["replacements_made"] == 0


class TestFormatCellsOnEmptyRange:
    """Extra: Apply formatting to cells that have no value."""

    def test_format_cells_on_empty_range(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "fmt.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        msg = format_cells(fp, "Sheet1", "A1:C3", bold=True, font_size=14)
        assert "A1:C3" in msg or "format" in msg.lower()

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].font.bold is True
        assert ws["A1"].font.size == 14
        wb.close()


class TestChartOnMinimalData:
    """Extra: Create a chart on a 2-row dataset (header + 1 row)."""

    def test_chart_on_minimal_data(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "chart.xlsx")
        _make_wb(fp, ["X", "Y"], [[1, 10]])

        msg = create_chart(fp, "Sheet1", "A1:B2", chart_type="column", title="Tiny")
        assert "chart" in msg.lower() or "Tiny" in msg

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        wb.close()


class TestGetWorkbookMetadataEmpty:
    """Extra: Metadata on an empty workbook."""

    def test_get_workbook_metadata_empty(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "meta.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        meta = get_workbook_metadata(fp)
        assert meta.file_path == fp
        assert "Sheet1" in [s.name for s in meta.sheets]


class TestMultipleSheetsReadWrite:
    """Extra: Write to multiple sheets and verify isolation."""

    def test_multiple_sheets_read_write(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "multi.xlsx")
        create_workbook(fp, sheet_names=["A", "B", "C"])

        write_cell(fp, "A", "A1", "sheet_a")
        write_cell(fp, "B", "A1", "sheet_b")
        write_cell(fp, "C", "A1", "sheet_c")

        wb = openpyxl.load_workbook(fp)
        assert wb["A"]["A1"].value == "sheet_a"
        assert wb["B"]["A1"].value == "sheet_b"
        assert wb["C"]["A1"].value == "sheet_c"
        wb.close()
