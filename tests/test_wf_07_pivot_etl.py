"""Workflow tests for pivot table, ETL, and data transformation operations.

Tests chain multiple MCP tool calls simulating real pivot/ETL workflows and verify
results independently using openpyxl and pandas — never via MCP tools.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from mcp_server.tools.cell_ops import write_range
from mcp_server.tools.pivot_etl import (
    add_computed_column,
    create_pivot_table,
    deduplicate_data,
    merge_datasets,
    refresh_pivot_table,
    unpivot_data,
)
from mcp_server.tools.workbook import create_workbook


def _create_sales_data(fp: str, sheet: str = "Sales") -> None:
    """Create a workbook with sample sales data."""
    create_workbook(fp, sheet_names=[sheet])
    write_range(
        fp,
        sheet,
        "A1",
        [
            ["Region", "Product", "Amount", "Qty"],
            ["North", "Widget", 100, 10],
            ["North", "Gadget", 200, 5],
            ["South", "Widget", 150, 8],
            ["South", "Gadget", 300, 12],
            ["North", "Widget", 120, 6],
            ["South", "Gadget", 250, 9],
        ],
    )


def _create_wide_data(fp: str, sheet: str = "Wide") -> None:
    """Create a workbook with wide-format data for unpivot tests."""
    create_workbook(fp, sheet_names=[sheet])
    write_range(
        fp,
        sheet,
        "A1",
        [
            ["Name", "Q1", "Q2", "Q3"],
            ["Alice", 100, 200, 300],
            ["Bob", 150, 250, 350],
            ["Charlie", 120, 220, 320],
        ],
    )


def _create_two_sheets(fp: str) -> None:
    """Create a workbook with two sheets for merge tests."""
    create_workbook(fp, sheet_names=["Employees", "Departments"])
    write_range(
        fp,
        "Employees",
        "A1",
        [
            ["EmpID", "Name", "DeptID"],
            [1, "Alice", 10],
            [2, "Bob", 20],
            [3, "Charlie", 10],
            [4, "Diana", 30],
        ],
    )
    write_range(
        fp,
        "Departments",
        "A1",
        [
            ["DeptID", "DeptName", "Budget"],
            [10, "Engineering", 500000],
            [20, "Marketing", 300000],
            [40, "HR", 200000],
        ],
    )


# ---------------------------------------------------------------------------
# 1. Create pivot — basic
# ---------------------------------------------------------------------------


class TestCreatePivotBasic:
    def test_create_pivot_basic(self, tmp_path: Path) -> None:
        """Write sales data → create_pivot_table → verify output sheet with pandas."""
        fp = str(tmp_path / "pivot.xlsx")
        _create_sales_data(fp)

        result = create_pivot_table(
            fp,
            "Sales",
            index_cols=["Region"],
            value_cols=["Amount"],
            aggfunc="sum",
            output_sheet="PivotOut",
        )

        assert "data" in result
        df = pd.read_excel(fp, sheet_name="PivotOut")
        assert "Region" in df.columns
        assert "Amount" in df.columns
        assert set(df["Region"].tolist()) == {"North", "South"}
        north_total = df.loc[df["Region"] == "North", "Amount"].iloc[0]
        assert north_total == 420  # 100 + 200 + 120


# ---------------------------------------------------------------------------
# 2. Create pivot — multiple value columns
# ---------------------------------------------------------------------------


class TestCreatePivotMultipleValues:
    def test_create_pivot_multiple_values(self, tmp_path: Path) -> None:
        """Pivot with 2 value columns → verify both present."""
        fp = str(tmp_path / "pivot_multi.xlsx")
        _create_sales_data(fp)

        result = create_pivot_table(
            fp,
            "Sales",
            index_cols=["Region"],
            value_cols=["Amount", "Qty"],
            aggfunc="sum",
            output_sheet="PivotMulti",
        )

        df = pd.read_excel(fp, sheet_name="PivotMulti")
        assert "Amount" in df.columns
        assert "Qty" in df.columns
        assert len(df) == 2  # North and South


# ---------------------------------------------------------------------------
# 3. Create pivot — with margins
# ---------------------------------------------------------------------------


class TestCreatePivotWithMargins:
    def test_create_pivot_with_margins(self, tmp_path: Path) -> None:
        """include_margins=True → verify total row exists."""
        fp = str(tmp_path / "pivot_margins.xlsx")
        _create_sales_data(fp)

        create_pivot_table(
            fp,
            "Sales",
            index_cols=["Region"],
            value_cols=["Amount"],
            aggfunc="sum",
            output_sheet="WithMargins",
            include_margins=True,
        )

        df = pd.read_excel(fp, sheet_name="WithMargins")
        regions = df["Region"].tolist()
        assert "Total" in regions
        total_row = df.loc[df["Region"] == "Total"]
        assert total_row["Amount"].iloc[0] == 1120  # sum of all amounts


# ---------------------------------------------------------------------------
# 4. Refresh pivot
# ---------------------------------------------------------------------------


class TestRefreshPivot:
    def test_refresh_pivot(self, tmp_path: Path) -> None:
        """Create pivot → modify source → refresh → verify updated values."""
        fp = str(tmp_path / "pivot_refresh.xlsx")
        _create_sales_data(fp)

        create_pivot_table(
            fp,
            "Sales",
            index_cols=["Region"],
            value_cols=["Amount"],
            aggfunc="sum",
            output_sheet="Refreshable",
        )

        # Modify source data — increase North Widget from 100 to 500
        wb = openpyxl.load_workbook(fp)
        ws = wb["Sales"]
        ws["C2"] = 500  # was 100
        wb.save(fp)
        wb.close()

        refresh_pivot_table(fp, "Refreshable")

        df = pd.read_excel(fp, sheet_name="Refreshable")
        north_total = df.loc[df["Region"] == "North", "Amount"].iloc[0]
        assert north_total == 820  # 500 + 200 + 120


# ---------------------------------------------------------------------------
# 5. Unpivot data
# ---------------------------------------------------------------------------


class TestUnpivotData:
    def test_unpivot_data(self, tmp_path: Path) -> None:
        """Write wide data → unpivot_data → verify melted format."""
        fp = str(tmp_path / "unpivot.xlsx")
        _create_wide_data(fp)

        result = unpivot_data(
            fp,
            "Wide",
            id_vars=["Name"],
            value_vars=["Q1", "Q2", "Q3"],
            var_name="Quarter",
            value_name="Revenue",
        )

        assert result["row_count"] == 9  # 3 names × 3 quarters
        data = result["data"]
        assert all("Name" in row and "Quarter" in row and "Revenue" in row for row in data)
        alice_q1 = [r for r in data if r["Name"] == "Alice" and r["Quarter"] == "Q1"]
        assert alice_q1[0]["Revenue"] == 100


# ---------------------------------------------------------------------------
# 6. Merge datasets — inner join
# ---------------------------------------------------------------------------


class TestMergeDatasetsInner:
    def test_merge_datasets_inner(self, tmp_path: Path) -> None:
        """Merge 2 datasets with inner join → verify only matching keys."""
        fp = str(tmp_path / "merge_inner.xlsx")
        _create_two_sheets(fp)

        result = merge_datasets(
            fp,
            "Employees",
            "Departments",
            join_key="DeptID",
            how="inner",
            output_sheet="InnerResult",
        )

        df = pd.read_excel(fp, sheet_name="InnerResult")
        # Only DeptID 10 and 20 exist in both sheets
        assert len(df) == 3  # Alice(10), Bob(20), Charlie(10)
        assert set(df["DeptID"].tolist()) == {10, 20}
        assert "DeptName" in df.columns


# ---------------------------------------------------------------------------
# 7. Merge datasets — left join
# ---------------------------------------------------------------------------


class TestMergeDatasetsLeft:
    def test_merge_datasets_left(self, tmp_path: Path) -> None:
        """Left join → verify nulls for non-matching keys."""
        fp = str(tmp_path / "merge_left.xlsx")
        _create_two_sheets(fp)

        result = merge_datasets(
            fp,
            "Employees",
            "Departments",
            join_key="DeptID",
            how="left",
            output_sheet="LeftResult",
        )

        df = pd.read_excel(fp, sheet_name="LeftResult")
        assert len(df) == 4  # all employees kept
        diana = df.loc[df["Name"] == "Diana"]
        assert diana["DeptName"].isna().iloc[0]  # DeptID=30 has no match


# ---------------------------------------------------------------------------
# 8. Merge datasets — left_on / right_on
# ---------------------------------------------------------------------------


class TestMergeDatasetsLeftOnRightOn:
    def test_merge_datasets_left_on_right_on(self, tmp_path: Path) -> None:
        """Different key column names → verify joined correctly."""
        fp = str(tmp_path / "merge_diff_keys.xlsx")
        create_workbook(fp, sheet_names=["Orders", "Products"])
        write_range(
            fp,
            "Orders",
            "A1",
            [
                ["OrderID", "ProdCode", "Qty"],
                [1, "A", 10],
                [2, "B", 5],
                [3, "C", 8],
            ],
        )
        write_range(
            fp,
            "Products",
            "A1",
            [
                ["SKU", "ProductName", "UnitPrice"],
                ["A", "Widget", 9.99],
                ["B", "Gadget", 19.99],
            ],
        )

        result = merge_datasets(
            fp,
            "Orders",
            "Products",
            left_on="ProdCode",
            right_on="SKU",
            how="left",
            output_sheet="Merged",
        )

        df = pd.read_excel(fp, sheet_name="Merged")
        assert len(df) == 3
        widget_row = df.loc[df["ProdCode"] == "A"]
        assert widget_row["ProductName"].iloc[0] == "Widget"
        c_row = df.loc[df["ProdCode"] == "C"]
        assert c_row["ProductName"].isna().iloc[0]


# ---------------------------------------------------------------------------
# 9. Add computed column — formula
# ---------------------------------------------------------------------------


class TestAddComputedColumnFormula:
    def test_add_computed_column_formula(self, tmp_path: Path) -> None:
        """Add column with formula expression → verify live Excel formulas written."""
        from openpyxl import load_workbook as lw

        fp = str(tmp_path / "computed.xlsx")
        _create_sales_data(fp)

        result = add_computed_column(
            fp,
            "Sales",
            new_column_name="Total",
            expression="Amount * Qty",
            column_type="formula",
        )
        assert "mode=formula" in str(result)

        # Verify numeric values written via openpyxl (formula mode now writes computed values)
        wb = lw(fp)
        ws = wb["Sales"]
        # Sales cols: Region(A=1), Product(B=2), Amount(C=3), Qty(D=4), Total(E=5)
        total_col = ws.max_column
        for row_idx in range(2, ws.max_row + 1):
            amount = ws.cell(row=row_idx, column=3).value  # Amount (C)
            qty = ws.cell(row=row_idx, column=4).value      # Qty (D)
            val = ws.cell(row=row_idx, column=total_col).value
            assert isinstance(val, (int, float))
            assert val == amount * qty
        wb.close()


# ---------------------------------------------------------------------------
# 10. Add computed column — cumsum
# ---------------------------------------------------------------------------


class TestAddComputedColumnCumsum:
    def test_add_computed_column_cumsum(self, tmp_path: Path) -> None:
        """Add cumsum column → verify running total."""
        fp = str(tmp_path / "cumsum.xlsx")
        _create_sales_data(fp)

        add_computed_column(
            fp,
            "Sales",
            new_column_name="RunningTotal",
            expression="",
            column_type="cumsum",
            source_col="Amount",
        )

        df = pd.read_excel(fp, sheet_name="Sales")
        assert "RunningTotal" in df.columns
        expected_cumsum = df["Amount"].cumsum()
        pd.testing.assert_series_equal(
            df["RunningTotal"],
            expected_cumsum,
            check_names=False,
        )


# ---------------------------------------------------------------------------
# 11. Deduplicate data
# ---------------------------------------------------------------------------


class TestDeduplicateData:
    def test_deduplicate_data(self, tmp_path: Path) -> None:
        """Insert duplicates → deduplicate_data → verify unique rows."""
        fp = str(tmp_path / "dedup.xlsx")
        create_workbook(fp, sheet_names=["Data"])
        write_range(
            fp,
            "Data",
            "A1",
            [
                ["Name", "Score"],
                ["Alice", 90],
                ["Bob", 85],
                ["Alice", 90],
                ["Charlie", 92],
                ["Bob", 85],
            ],
        )

        result = deduplicate_data(fp, "Data", columns=["Name", "Score"])

        df = pd.read_excel(fp, sheet_name="Data")
        assert len(df) == 3  # Alice, Bob, Charlie
        assert "Removed 2" in result


# ---------------------------------------------------------------------------
# 12. Pivot then unpivot — round-trip shape
# ---------------------------------------------------------------------------


class TestPivotThenUnpivot:
    def test_pivot_then_unpivot(self, tmp_path: Path) -> None:
        """Pivot → unpivot result → verify shape is consistent."""
        fp = str(tmp_path / "roundtrip.xlsx")
        _create_sales_data(fp)

        create_pivot_table(
            fp,
            "Sales",
            index_cols=["Region"],
            value_cols=["Amount"],
            aggfunc="sum",
            output_sheet="Pivoted",
        )

        # Read pivoted sheet to get columns
        df_pivot = pd.read_excel(fp, sheet_name="Pivoted")
        assert len(df_pivot) == 2  # North, South

        result = unpivot_data(
            fp,
            "Pivoted",
            id_vars=["Region"],
            value_vars=["Amount"],
            var_name="Metric",
            value_name="Value",
        )

        assert result["row_count"] == 2  # 2 regions × 1 value var
        for row in result["data"]:
            assert row["Metric"] == "Amount"


# ---------------------------------------------------------------------------
# 13. Merge three datasets
# ---------------------------------------------------------------------------


class TestMergeThreeDatasets:
    def test_merge_three_datasets(self, tmp_path: Path) -> None:
        """Merge A+B then result+C → verify chained merge."""
        fp = str(tmp_path / "triple_merge.xlsx")
        create_workbook(fp, sheet_names=["Orders", "Products", "Categories"])
        write_range(
            fp,
            "Orders",
            "A1",
            [
                ["OrderID", "ProdID", "Qty"],
                [1, 100, 2],
                [2, 101, 5],
            ],
        )
        write_range(
            fp,
            "Products",
            "A1",
            [
                ["ProdID", "ProdName", "CatID"],
                [100, "Widget", "C1"],
                [101, "Gadget", "C2"],
            ],
        )
        write_range(
            fp,
            "Categories",
            "A1",
            [
                ["CatID", "CatName"],
                ["C1", "Hardware"],
                ["C2", "Software"],
            ],
        )

        # First merge: Orders + Products
        merge_datasets(
            fp,
            "Orders",
            "Products",
            join_key="ProdID",
            how="inner",
            output_sheet="OrderProducts",
        )

        # Second merge: OrderProducts + Categories
        merge_datasets(
            fp,
            "OrderProducts",
            "Categories",
            join_key="CatID",
            how="inner",
            output_sheet="FullResult",
        )

        df = pd.read_excel(fp, sheet_name="FullResult")
        assert len(df) == 2
        assert "CatName" in df.columns
        assert "ProdName" in df.columns
        widget_row = df.loc[df["ProdName"] == "Widget"]
        assert widget_row["CatName"].iloc[0] == "Hardware"


# ---------------------------------------------------------------------------
# 14. Computed column on pivoted data
# ---------------------------------------------------------------------------


class TestComputedColumnOnPivotedData:
    def test_computed_column_on_pivoted_data(self, tmp_path: Path) -> None:
        """Pivot → add computed column → verify."""
        fp = str(tmp_path / "pivot_computed.xlsx")
        _create_sales_data(fp)

        create_pivot_table(
            fp,
            "Sales",
            index_cols=["Region"],
            value_cols=["Amount", "Qty"],
            aggfunc="sum",
            output_sheet="PivotCalc",
        )

        add_computed_column(
            fp,
            "PivotCalc",
            new_column_name="AvgPrice",
            expression="Amount / Qty",
            column_type="formula",
        )

        # Verify numeric values written via openpyxl (formula mode now writes computed values)
        from openpyxl import load_workbook as lw

        wb = lw(fp)
        ws = wb["PivotCalc"]
        # PivotCalc cols: Region(A=1), Amount(B=2), Qty(C=3), AvgPrice(D=4)
        avgprice_col = ws.max_column
        for row_idx in range(2, ws.max_row + 1):
            amount = ws.cell(row=row_idx, column=2).value  # Amount
            qty = ws.cell(row=row_idx, column=3).value      # Qty
            val = ws.cell(row=row_idx, column=avgprice_col).value
            assert isinstance(val, (int, float))
            assert abs(val - amount / qty) < 1e-9
        wb.close()


# ---------------------------------------------------------------------------
# 15. Deduplicate — keep last
# ---------------------------------------------------------------------------


class TestDeduplicateKeepLast:
    def test_deduplicate_keep_last(self, tmp_path: Path) -> None:
        """keep='last' → verify correct row kept."""
        fp = str(tmp_path / "dedup_last.xlsx")
        create_workbook(fp, sheet_names=["Data"])
        write_range(
            fp,
            "Data",
            "A1",
            [
                ["ID", "Value"],
                [1, "first"],
                [2, "only"],
                [1, "last"],
            ],
        )

        deduplicate_data(fp, "Data", columns=["ID"], keep="last")

        df = pd.read_excel(fp, sheet_name="Data")
        assert len(df) == 2
        id1_row = df.loc[df["ID"] == 1]
        assert id1_row["Value"].iloc[0] == "last"


# ---------------------------------------------------------------------------
# 16. Pivot with column_field
# ---------------------------------------------------------------------------


class TestPivotWithColumnField:
    def test_pivot_with_column_field(self, tmp_path: Path) -> None:
        """Pivot with column_field → verify cross-tabulated columns."""
        fp = str(tmp_path / "pivot_colfield.xlsx")
        _create_sales_data(fp)

        create_pivot_table(
            fp,
            "Sales",
            index_cols=["Region"],
            value_cols=["Amount"],
            aggfunc="sum",
            output_sheet="CrossTab",
            column_field="Product",
        )

        df = pd.read_excel(fp, sheet_name="CrossTab")
        # Should have Region + columns for each Product
        assert "Region" in df.columns
        col_names = [c for c in df.columns if c != "Region"]
        assert len(col_names) >= 2  # At least Gadget and Widget columns


# ---------------------------------------------------------------------------
# 17. Merge with outer join
# ---------------------------------------------------------------------------


class TestMergeOuterJoin:
    def test_merge_outer_join(self, tmp_path: Path) -> None:
        """Outer join → verify all rows from both sides present."""
        fp = str(tmp_path / "merge_outer.xlsx")
        _create_two_sheets(fp)

        result = merge_datasets(
            fp,
            "Employees",
            "Departments",
            join_key="DeptID",
            how="outer",
            output_sheet="OuterResult",
        )

        df = pd.read_excel(fp, sheet_name="OuterResult")
        # 4 employees + HR dept (DeptID=40) unmatched = 5 rows
        assert len(df) == 5
        hr_row = df.loc[df["DeptName"] == "HR"]
        assert hr_row["Name"].isna().iloc[0]
