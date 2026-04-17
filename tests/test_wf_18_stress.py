"""Stress tests and regression tests for edge cases and cross-feature interactions.

Tests call tool functions directly and verify results independently using
openpyxl or pandas — never via MCP tools.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from mcp_server.tools.analysis import (
    aggregate_data,
    filter_data_advanced,
    sort_data,
)
from mcp_server.tools.cell_ops import (
    read_cell,
    read_file_chunked,
    read_range,
    transpose_range,
    write_cell,
    write_range,
)
from mcp_server.tools.charts import create_chart
from mcp_server.tools.conditional_formatting import (
    add_highlight_rule,
    apply_conditional_formatting,
)
from mcp_server.tools.custom_code import execute_custom_code
from mcp_server.tools.data_validation import (
    add_dropdown_validation,
    add_numeric_validation,
)
from mcp_server.tools.formatting import format_cells
from mcp_server.tools.formulas import (
    list_formulas,
    set_formula,
    set_formulas_batch,
)
from mcp_server.tools.images import insert_image
from mcp_server.tools.named_ranges import (
    create_named_range,
    list_named_ranges,
)
from mcp_server.tools.pivot_etl import create_pivot_table
from mcp_server.tools.scenarios import add_scenario, apply_scenario, list_scenarios
from mcp_server.tools.tables import create_table, list_tables
from mcp_server.tools.workbook import (
    copy_sheet,
    create_workbook,
    delete_sheet,
    get_workbook_metadata,
)
from mcp_server.tools.worksheet_ops import (
    copy_range_across_sheets,
    copy_sheet_across_workbooks,
    delete_rows,
    insert_rows,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_wb(fp: str, headers: list[str], rows: list[list]) -> None:
    create_workbook(fp, sheet_names=["Sheet1"])
    write_range(fp, "Sheet1", "A1", [headers] + rows)


def _large_dataset(n: int) -> tuple[list[str], list[list]]:
    """Generate a dataset with n rows: Name, Value, Category."""
    headers = ["Name", "Value", "Category"]
    cats = ["A", "B", "C", "D", "E"]
    rows = [[f"Item_{i}", i * 10, cats[i % len(cats)]] for i in range(1, n + 1)]
    return headers, rows


# ===========================================================================
# Large Data Stress
# ===========================================================================


class TestWrite5000Rows:
    """1. Write 5000 rows and verify with pandas."""

    def test_write_5000_rows(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "big.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        headers = ["ID", "Value"]
        rows = [[i, i * 100] for i in range(1, 5001)]
        write_range(fp, "Sheet1", "A1", [headers] + rows)

        df = pd.read_excel(fp, sheet_name="Sheet1", engine="openpyxl")
        assert len(df) == 5000
        assert df.iloc[0]["ID"] == 1
        assert df.iloc[4999]["ID"] == 5000
        assert df.iloc[4999]["Value"] == 500000


class TestReadChunkedAllChunks:
    """2. Read 2000 rows in 4 chunks of 500."""

    def test_read_chunked_all_chunks(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "chunked.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        headers = ["ID", "Val"]
        rows = [[i, i * 2] for i in range(1, 2001)]
        write_range(fp, "Sheet1", "A1", [headers] + rows)

        all_rows: list = []
        start = 0
        for chunk_num in range(4):
            result = read_file_chunked(fp, "Sheet1", start_row=start, chunk_size=500)
            all_rows.extend(result["rows"])
            if chunk_num < 3:
                assert result["has_more"] is True
            start = result["next_start_row"]

        assert len(all_rows) == 2000
        assert all_rows[0]["ID"] == 1
        assert all_rows[1999]["ID"] == 2000


class TestFilterLargeDataset:
    """3. Filter a 1000-row dataset."""

    def test_filter_large_dataset(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "filter_large.xlsx")
        headers, rows = _large_dataset(1000)
        _make_wb(fp, headers, rows)

        result = filter_data_advanced(
            fp,
            "Sheet1",
            conditions=[{"column": "Category", "operator": "==", "value": "A"}],
        )
        # Items 1,6,11,...,996  → i%5==1 → 200 items
        assert result["rows"] == 200


class TestSortLargeDataset:
    """4. Sort a 1000-row dataset."""

    def test_sort_large_dataset(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "sort_large.xlsx")
        headers, rows = _large_dataset(1000)
        _make_wb(fp, headers, rows)

        sort_data(fp, "Sheet1", column="Value", ascending=False)

        df = pd.read_excel(fp, sheet_name="Sheet1", engine="openpyxl")
        assert df.iloc[0]["Value"] == 10000  # 1000*10
        assert df.iloc[999]["Value"] == 10  # 1*10


class TestAggregateLargeDataset:
    """5. Aggregate 1000 rows."""

    def test_aggregate_large_dataset(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "agg_large.xlsx")
        headers, rows = _large_dataset(1000)
        _make_wb(fp, headers, rows)

        result = aggregate_data(
            fp,
            "Sheet1",
            group_by="Category",
            value_column="Value",
            operation="sum",
        )
        assert len(result["groups"]) == 5  # A,B,C,D,E
        total = sum(g["Value_sum"] for g in result["groups"])
        expected = sum(i * 10 for i in range(1, 1001))
        assert total == expected


# ===========================================================================
# Multiple Operations on Same File
# ===========================================================================


class TestSequentialWrites:
    """6. 50 sequential write_cell calls."""

    def test_50_sequential_writes(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "seq.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        for i in range(1, 51):
            write_cell(fp, "Sheet1", f"A{i}", i * 100)

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        for i in range(1, 51):
            assert ws[f"A{i}"].value == i * 100
        wb.close()


class TestInsertDeleteCycle:
    """7. Insert 5 rows → delete 3 → insert 2 → net +4."""

    def test_insert_delete_cycle(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "cycle.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        # Write 10 rows of data
        data = [["Row", "Val"]] + [[i, i] for i in range(1, 11)]
        write_range(fp, "Sheet1", "A1", data)

        initial_rows = 11  # 1 header + 10 data

        insert_rows(fp, "Sheet1", row=3, count=5)  # +5
        delete_rows(fp, "Sheet1", row=6, count=3)  # -3
        insert_rows(fp, "Sheet1", row=2, count=2)  # +2

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws.max_row == initial_rows + 4  # net +4
        wb.close()


class TestFormatReformatCycle:
    """8. Format → reformat with different styles → verify final."""

    def test_format_reformat_cycle(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "reformat.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", "Test")

        format_cells(fp, "Sheet1", "A1", bold=True, font_size=14, font_color="FF0000")
        format_cells(fp, "Sheet1", "A1", bold=False, italic=True, font_size=10, font_color="0000FF")

        wb = openpyxl.load_workbook(fp)
        cell = wb["Sheet1"]["A1"]
        assert cell.font.italic is True
        assert cell.font.bold is not True  # overwritten
        assert cell.font.size == 10
        wb.close()


class TestCreateDeleteMultipleSheets:
    """9. Create 10 sheets → delete 5 → verify remaining 5."""

    def test_create_delete_multiple_sheets(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "sheets.xlsx")
        names = [f"S{i}" for i in range(1, 11)]
        create_workbook(fp, sheet_names=names)

        for name in names[:5]:
            delete_sheet(fp, name)

        wb = openpyxl.load_workbook(fp)
        assert len(wb.sheetnames) == 5
        assert set(wb.sheetnames) == {f"S{i}" for i in range(6, 11)}
        wb.close()


class TestMultipleTablesInWorkbook:
    """10. Create 3 tables in different ranges."""

    def test_multiple_tables_in_workbook(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "multi_tbl.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        # Table 1: A1:B4
        write_range(fp, "Sheet1", "A1", [["H1", "H2"], [1, 2], [3, 4], [5, 6]])
        create_table(fp, "Sheet1", "A1:B4", "Table1")

        # Table 2: D1:E4
        write_range(fp, "Sheet1", "D1", [["X", "Y"], [10, 20], [30, 40], [50, 60]])
        create_table(fp, "Sheet1", "D1:E4", "Table2")

        # Table 3: G1:H4
        write_range(fp, "Sheet1", "G1", [["P", "Q"], [100, 200], [300, 400], [500, 600]])
        create_table(fp, "Sheet1", "G1:H4", "Table3")

        tables = list_tables(fp, "Sheet1")
        assert len(tables) == 3
        table_names = {t["name"] for t in tables}
        assert table_names == {"Table1", "Table2", "Table3"}


# ===========================================================================
# Cross-Feature Interactions
# ===========================================================================


class TestFormulaSurvivesInsertRows:
    """11. Set formula → insert rows above → verify formula shifted."""

    def test_formula_survives_insert_rows(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "formula_shift.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [[10], [20], [30]])

        # SUM in A4
        set_formula(fp, "Sheet1", "A4", "=SUM(A1:A3)")

        # Insert 2 rows at row 2 → formula should shift from A4 to A6
        insert_rows(fp, "Sheet1", row=2, count=2)

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        # After inserting 2 rows at row 2, the formula was in A4 → now A6
        formula_cell = ws["A6"].value
        assert formula_cell is not None
        assert "SUM" in str(formula_cell).upper()
        wb.close()


class TestTableAfterFormatting:
    """12. Format range → create table on formatted range → verify both."""

    def test_table_after_formatting(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "fmt_table.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [["Name", "Score"], ["Alice", 90], ["Bob", 85]])

        format_cells(fp, "Sheet1", "A1:B1", bold=True, bg_color="FFFF00")
        create_table(fp, "Sheet1", "A1:B3", "ScoreTable")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        # Table exists
        assert "ScoreTable" in ws.tables
        # Formatting preserved: header bold
        assert ws["A1"].font.bold is True
        wb.close()


class TestChartAfterFilter:
    """13. Write data → create chart → verify chart exists."""

    def test_chart_after_data(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "chart_data.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(
            fp,
            "Sheet1",
            "A1",
            [["Month", "Sales"], ["Jan", 100], ["Feb", 200], ["Mar", 150]],
        )

        create_chart(
            fp,
            "Sheet1",
            data_range="B1:B4",
            chart_type="column",
            target_cell="D1",
            title="Sales Chart",
            categories_range="A2:A4",
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        # Chart title is a Title object; extract text from rich text paragraphs
        chart_title = ws._charts[0].title
        title_text = chart_title.tx.rich.paragraphs[0].r[0].t
        assert title_text == "Sales Chart"
        wb.close()


class TestPivotFromTable:
    """14. Create table → create pivot from table range → verify."""

    def test_pivot_from_table(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "pivot_tbl.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(
            fp,
            "Sheet1",
            "A1",
            [
                ["Region", "Product", "Sales"],
                ["East", "A", 100],
                ["West", "B", 200],
                ["East", "A", 150],
                ["West", "B", 250],
            ],
        )
        create_table(fp, "Sheet1", "A1:C5", "SalesTable")

        create_pivot_table(
            fp,
            "Sheet1",
            index_cols=["Region"],
            value_cols=["Sales"],
            aggfunc="sum",
            output_sheet="PivotResult",
        )

        wb = openpyxl.load_workbook(fp)
        assert "PivotResult" in wb.sheetnames
        ws = wb["PivotResult"]
        # Should have header + 2 region rows
        assert ws.max_row >= 3
        wb.close()


class TestConditionalFormatWithTable:
    """15. Create table → add conditional formatting to table range."""

    def test_conditional_format_with_table(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "cf_table.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(
            fp,
            "Sheet1",
            "A1",
            [["Name", "Score"], ["Alice", 90], ["Bob", 60], ["Carol", 75]],
        )
        create_table(fp, "Sheet1", "A1:B4", "Scores")

        add_highlight_rule(
            fp,
            "Sheet1",
            "B2:B4",
            operator="lessThan",
            formula="70",
            font_color="FF0000",
            bg_color="FFCCCC",
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert "Scores" in ws.tables
        assert len(ws.conditional_formatting) >= 1
        wb.close()


# ===========================================================================
# Error Recovery
# ===========================================================================


class TestOperationsAfterError:
    """16. Trigger error → perform normal operation → verify works."""

    def test_operations_after_error(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "recovery.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", "original")

        # Trigger error: reading non-existent sheet
        with pytest.raises(Exception):
            read_cell(fp, "NonExistent", "A1")

        # Normal operation should still work
        write_cell(fp, "Sheet1", "B1", "after_error")

        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["A1"].value == "original"
        assert wb["Sheet1"]["B1"].value == "after_error"
        wb.close()


class TestSaveAfterMultipleErrors:
    """17. Multiple failed ops → successful op → verify file valid."""

    def test_save_after_multiple_errors(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "multi_err.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", 42)

        # Multiple errors
        for _ in range(3):
            with pytest.raises(Exception):
                read_cell(fp, "BadSheet", "A1")

        # Successful write
        write_cell(fp, "Sheet1", "A2", 99)

        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["A1"].value == 42
        assert wb["Sheet1"]["A2"].value == 99
        wb.close()


class TestReadAfterFailedWrite:
    """18. Failed write (bad sheet) → successful read → data intact."""

    def test_read_after_failed_write(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "failed_write.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", "safe_data")

        with pytest.raises(Exception):
            write_cell(fp, "NoSheet", "A1", "bad")

        result = read_cell(fp, "Sheet1", "A1")
        assert result["value"] == "safe_data"


# ===========================================================================
# Custom Code
# ===========================================================================


class TestExecuteCustomCodeBasic:
    """19. Basic pandas operation via custom code."""

    def test_execute_custom_code_basic(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "custom.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [["X", "Y"], [1, 10], [2, 20], [3, 30]])

        result = execute_custom_code(
            fp,
            code="result = df['Y'].sum()",
            sheet="Sheet1",
        )
        assert result["result"] == 60


class TestExecuteCustomCodeWithNumpy:
    """20. Numpy calculation via custom code."""

    def test_execute_custom_code_with_numpy(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "custom_np.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [["Val"], [4], [9], [16]])

        result = execute_custom_code(
            fp,
            code="result = float(np.mean(df['Val']))",
            sheet="Sheet1",
        )
        # mean of 4,9,16 = 29/3 ≈ 9.666...
        assert abs(result["result"] - 9.666666) < 0.01


class TestExecuteCustomCodeForbiddenImport:
    """21. Importing os is rejected."""

    def test_execute_custom_code_forbidden_import(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "bad_import.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [["X"], [1]])

        result = execute_custom_code(fp, code="import os\nresult = os.getcwd()", sheet="Sheet1")
        assert result["status"] == "error"
        assert "import" in result["message"].lower() or "blocked" in result["message"].lower()


class TestExecuteCustomCodeForbiddenEval:
    """22. Calling eval() is rejected."""

    def test_execute_custom_code_forbidden_eval(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "bad_eval.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [["X"], [1]])

        result = execute_custom_code(fp, code="result = eval('1+1')", sheet="Sheet1")
        assert result["status"] == "error"
        assert "eval" in result["message"].lower() or "blocked" in result["message"].lower()


# ===========================================================================
# Image Operations
# ===========================================================================


class TestInsertImage:
    """23. Create test PNG → insert into worksheet → verify."""

    def test_insert_image(self, tmp_path: Path) -> None:
        # Create a minimal valid PNG
        from PIL import Image as PILImage

        img_path = str(tmp_path / "test.png")
        img = PILImage.new("RGB", (50, 50), color="red")
        img.save(img_path)

        fp = str(tmp_path / "img.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        insert_image(fp, "Sheet1", img_path, "B2", width=100, height=100)

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._images) == 1
        wb.close()


# ===========================================================================
# Copy/Paste Intensive
# ===========================================================================


class TestCopyRangeAcrossMultipleSheets:
    """24. Copy data across 3 sheets → verify all."""

    def test_copy_range_across_multiple_sheets(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "copy_multi.xlsx")
        create_workbook(fp, sheet_names=["Src", "Dst1", "Dst2", "Dst3"])
        write_range(fp, "Src", "A1", [["A", "B"], [1, 2], [3, 4]])

        for dst in ["Dst1", "Dst2", "Dst3"]:
            copy_range_across_sheets(fp, "Src", "A1:B3", dst, "A1")

        wb = openpyxl.load_workbook(fp)
        for dst in ["Dst1", "Dst2", "Dst3"]:
            assert wb[dst]["A1"].value == "A"
            assert wb[dst]["B2"].value == 2
            assert wb[dst]["A3"].value == 3
        wb.close()


class TestCopySheetAcrossWorkbooks:
    """25. Copy sheet between 2 workbooks."""

    def test_copy_sheet_across_workbooks(self, tmp_path: Path) -> None:
        src = str(tmp_path / "src_wb.xlsx")
        dst = str(tmp_path / "dst_wb.xlsx")

        create_workbook(src, sheet_names=["Data"])
        write_range(src, "Data", "A1", [["Col1", "Col2"], [10, 20], [30, 40]])

        create_workbook(dst, sheet_names=["Existing"])

        copy_sheet_across_workbooks(src, "Data", dst, dest_sheet_name="Imported")

        wb = openpyxl.load_workbook(dst)
        assert "Imported" in wb.sheetnames
        assert wb["Imported"]["A1"].value == "Col1"
        assert wb["Imported"]["B3"].value == 40
        wb.close()


class TestTransposeThenCopy:
    """26. Transpose → copy → verify."""

    def test_transpose_then_copy(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "trans_copy.xlsx")
        create_workbook(fp, sheet_names=["Sheet1", "Sheet2"])
        write_range(fp, "Sheet1", "A1", [[1, 2, 3], [4, 5, 6]])

        transpose_range(fp, "Sheet1", "A1:C2", "E1")

        # Copy the transposed data to Sheet2
        copy_range_across_sheets(fp, "Sheet1", "E1:F3", "Sheet2", "A1")

        wb = openpyxl.load_workbook(fp)
        # Transposed: 2x3 → 3x2
        # Original: [[1,2,3],[4,5,6]]
        # Transposed at E1: [[1,4],[2,5],[3,6]]
        assert wb["Sheet1"]["E1"].value == 1
        assert wb["Sheet1"]["F1"].value == 4
        assert wb["Sheet1"]["E3"].value == 3
        assert wb["Sheet1"]["F3"].value == 6
        # Copied to Sheet2
        assert wb["Sheet2"]["A1"].value == 1
        assert wb["Sheet2"]["B1"].value == 4
        wb.close()


# ===========================================================================
# Named Range Interactions
# ===========================================================================


class TestNamedRangeInFormula:
    """27. Create named range → use in formula → verify."""

    def test_named_range_in_formula(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "named_formula.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [[10], [20], [30]])

        create_named_range(fp, "MyRange", "Sheet1!$A$1:$A$3")
        set_formula(fp, "Sheet1", "B1", "=SUM(MyRange)")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["B1"].value == "=SUM(MyRange)"
        # Verify named range exists
        names = [nr.name for nr in wb.defined_names.values()]
        assert "MyRange" in names
        wb.close()


class TestNamedRangeAfterInsertRows:
    """28. Create range → insert rows → verify range still in defined names."""

    def test_named_range_after_insert_rows(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "named_insert.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [[10], [20], [30]])

        create_named_range(fp, "DataBlock", "Sheet1!$A$1:$A$3")
        insert_rows(fp, "Sheet1", row=1, count=2)

        # Named range should still exist (openpyxl preserves defined names)
        ranges = list_named_ranges(fp)
        range_names = [r["name"] for r in ranges]
        assert "DataBlock" in range_names


# ===========================================================================
# Scenario Interactions
# ===========================================================================


class TestMultipleScenarios:
    """29. Add 3 scenarios → apply each → verify values change."""

    def test_multiple_scenarios(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "scenarios.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", 0)
        write_cell(fp, "Sheet1", "A2", 0)

        add_scenario(fp, "Low", {"Sheet1": {"A1": 10, "A2": 20}}, description="Low scenario")
        add_scenario(fp, "Mid", {"Sheet1": {"A1": 50, "A2": 60}}, description="Mid scenario")
        add_scenario(fp, "High", {"Sheet1": {"A1": 100, "A2": 200}}, description="High scenario")

        scenarios = list_scenarios(fp)
        assert len(scenarios) == 3

        # Apply Low
        apply_scenario(fp, "Low")
        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["A1"].value == 10
        assert wb["Sheet1"]["A2"].value == 20
        wb.close()

        # Apply High
        apply_scenario(fp, "High")
        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["A1"].value == 100
        assert wb["Sheet1"]["A2"].value == 200
        wb.close()

        # Apply Mid
        apply_scenario(fp, "Mid")
        wb = openpyxl.load_workbook(fp)
        assert wb["Sheet1"]["A1"].value == 50
        assert wb["Sheet1"]["A2"].value == 60
        wb.close()


class TestScenarioWithFormulas:
    """30. Scenario changes formula inputs → verify formula cell after apply."""

    def test_scenario_with_formulas(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "scen_formula.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", 10)
        write_cell(fp, "Sheet1", "A2", 20)
        set_formula(fp, "Sheet1", "A3", "=A1+A2")

        add_scenario(fp, "Double", {"Sheet1": {"A1": 100, "A2": 200}})
        apply_scenario(fp, "Double")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert ws["A1"].value == 100
        assert ws["A2"].value == 200
        # Formula should still be intact
        assert ws["A3"].value == "=A1+A2"
        wb.close()


# ===========================================================================
# Data Validation Combinations
# ===========================================================================


class TestValidationOnTableColumn:
    """31. Create table → add validation to column → verify."""

    def test_validation_on_table_column(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "val_table.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(
            fp,
            "Sheet1",
            "A1",
            [["Status", "Value"], ["Open", 10], ["Closed", 20]],
        )
        create_table(fp, "Sheet1", "A1:B3", "StatusTable")

        add_dropdown_validation(
            fp,
            "Sheet1",
            "A2:A100",
            options=["Open", "Closed", "Pending"],
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert "StatusTable" in ws.tables
        assert len(ws.data_validations.dataValidation) >= 1
        wb.close()


class TestMultipleValidationsSameSheet:
    """32. 3 different validation types on same sheet."""

    def test_multiple_validations_same_sheet(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "multi_val.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(
            fp,
            "Sheet1",
            "A1",
            [["Choice", "Score", "Date"], ["X", 50, "2024-01-01"]],
        )

        add_dropdown_validation(fp, "Sheet1", "A2:A100", options=["X", "Y", "Z"])
        add_numeric_validation(fp, "Sheet1", "B2:B100", operator="between", value1=0, value2=100)
        add_dropdown_validation(fp, "Sheet1", "C2:C100", options=["2024-01-01", "2024-06-01", "2024-12-31"])

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws.data_validations.dataValidation) == 3
        wb.close()


# ===========================================================================
# Additional Regression Tests
# ===========================================================================


class TestWriteRangeThenReadRange:
    """33. Write a range and immediately read it back — verify round-trip."""

    def test_write_range_then_read_range(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "roundtrip.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        data = [["A", "B", "C"], [1, 2, 3], [4, 5, 6], [7, 8, 9]]
        write_range(fp, "Sheet1", "A1", data)

        result = read_range(fp, "Sheet1", "A1", "C4")
        assert result["row_count"] == 4
        assert result["col_count"] == 3
        assert result["rows"][0] == ["A", "B", "C"]
        assert result["rows"][3] == [7, 8, 9]


class TestBatchFormulasWithListFormulas:
    """34. Batch set formulas → list all formulas → verify count."""

    def test_batch_formulas_with_list(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "batch_form.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [[10], [20], [30], [40], [50]])

        formulas = {
            "B1": "=A1*2",
            "B2": "=A2*2",
            "B3": "=A3*2",
            "B4": "=A4*2",
            "B5": "=A5*2",
        }
        set_formulas_batch(fp, "Sheet1", formulas)

        found = list_formulas(fp, "Sheet1")
        assert len(found) >= 5


class TestCopySheetPreservesData:
    """35. Copy sheet within workbook → verify data integrity."""

    def test_copy_sheet_preserves_data(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "copy_internal.xlsx")
        create_workbook(fp, sheet_names=["Original"])
        write_range(fp, "Original", "A1", [["ID", "Name"], [1, "Alice"], [2, "Bob"]])

        copy_sheet(fp, "Original", "Backup")

        wb = openpyxl.load_workbook(fp)
        assert "Backup" in wb.sheetnames
        assert wb["Backup"]["A1"].value == "ID"
        assert wb["Backup"]["B3"].value == "Bob"
        wb.close()


class TestConditionalFormattingColorScale:
    """36. Apply color scale CF → verify rule persists."""

    def test_conditional_formatting_color_scale(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "cf_scale.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [["Score"]] + [[i * 10] for i in range(1, 11)])

        apply_conditional_formatting(
            fp,
            "Sheet1",
            "A2:A11",
            format_type="color_scale",
            start_color="FF0000",
            mid_color="FFFF00",
            end_color="00FF00",
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws.conditional_formatting) >= 1
        wb.close()


class TestWorkbookMetadataAfterManyChanges:
    """37. Many operations → get metadata → verify consistency."""

    def test_metadata_after_changes(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "meta.xlsx")
        create_workbook(fp, sheet_names=["Data", "Summary"])
        write_range(fp, "Data", "A1", [["X", "Y"], [1, 2], [3, 4]])
        write_cell(fp, "Summary", "A1", "Total")
        set_formula(fp, "Summary", "B1", "=Data!A2+Data!A3")
        create_named_range(fp, "AllData", "Data!$A$1:$B$3")

        meta = get_workbook_metadata(fp)
        sheet_names = [s.name for s in meta.sheets]
        assert "Data" in sheet_names
        assert "Summary" in sheet_names
        assert any(nr["name"] == "AllData" for nr in meta.named_ranges)


class TestCustomCodeForbiddenDunderAccess:
    """38. Accessing dunder attributes is rejected."""

    def test_custom_code_forbidden_dunder(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "dunder.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [["X"], [1]])

        result = execute_custom_code(fp, code="result = df.__class__.__name__", sheet="Sheet1")
        assert result["status"] == "error"
        assert "dunder" in result["message"].lower() or "blocked" in result["message"].lower()


class TestCustomCodeForbiddenSystemAttr:
    """39. Accessing os.system via attribute is rejected."""

    def test_custom_code_forbidden_system_attr(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "sys_attr.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [["X"], [1]])

        result = execute_custom_code(fp, code="result = os.getcwd()", sheet="Sheet1")
        assert result["status"] == "error"
        assert "os" in result["message"].lower() or "blocked" in result["message"].lower()


class TestCustomCodeReturnDataFrame:
    """40. Custom code returning a DataFrame → written back to file."""

    def test_custom_code_return_dataframe(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "cc_df.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(fp, "Sheet1", "A1", [["A", "B"], [1, 10], [2, 20], [3, 30]])

        out = str(tmp_path / "cc_out.xlsx")
        execute_custom_code(
            fp,
            code="result = df[df['B'] > 15]",
            sheet="Sheet1",
            output_file=out,
        )

        df = pd.read_excel(out, engine="openpyxl")
        assert len(df) == 2  # rows where B > 15: [20, 30]


class TestFormatPreserveExisting:
    """41. Format with preserve_existing → only specified attrs change."""

    def test_format_preserve_existing(self, tmp_path: Path) -> None:
        fp = str(tmp_path / "preserve.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_cell(fp, "Sheet1", "A1", "Hello")

        # First format: bold + red
        format_cells(fp, "Sheet1", "A1", bold=True, font_color="FF0000")
        # Second format: italic only, preserve existing
        format_cells(fp, "Sheet1", "A1", italic=True, preserve_existing=True)

        wb = openpyxl.load_workbook(fp)
        cell = wb["Sheet1"]["A1"]
        assert cell.font.bold is True  # preserved
        assert cell.font.italic is True  # newly applied
        wb.close()
