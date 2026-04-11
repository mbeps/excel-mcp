"""Workflow tests for data analysis operations.

Tests chain multiple MCP tool calls simulating real analysis workflows and verify
results independently using openpyxl, pandas, and numpy — never via MCP tools.
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
import pytest

from mcp_server.tools.analysis import (
    aggregate_data,
    column_statistics,
    filter_data_advanced,
    find_duplicates,
    insert_subtotals,
    profile_data,
    sort_data,
    vlookup_helper,
)
from mcp_server.tools.cell_ops import write_range
from mcp_server.tools.workbook import create_workbook


def _create_analysis_data(fp: str, sheet: str = "Data") -> None:
    """Create a workbook with sample data for analysis tests."""
    create_workbook(fp, sheet_names=[sheet])
    write_range(
        fp,
        sheet,
        "A1",
        [
            ["Name", "Department", "Age", "Salary", "City"],
            ["Alice", "Engineering", 30, 95000, "New York"],
            ["Bob", "Marketing", 25, 55000, "Chicago"],
            ["Charlie", "Engineering", 35, 110000, "Boston"],
            ["Diana", "Marketing", 28, 62000, "Chicago"],
            ["Eve", "Engineering", 32, 88000, "New York"],
            ["Frank", "Sales", 40, 75000, "Boston"],
            ["Grace", "Sales", 27, 58000, "New York"],
            ["Hank", "Marketing", 33, 70000, "Chicago"],
        ],
    )


# ---------------------------------------------------------------------------
# 1. Filter data — basic equality
# ---------------------------------------------------------------------------


class TestFilterDataBasic:
    def test_filter_data_basic(self, tmp_path: Path) -> None:
        """Write data → filter with == → verify filtered rows."""
        fp = str(tmp_path / "filter_basic.xlsx")
        _create_analysis_data(fp)

        result = filter_data_advanced(
            fp,
            "Data",
            conditions=[{"column": "Department", "operator": "==", "value": "Engineering"}],
        )

        assert result["rows"] == 3
        for row in result["data"]:
            assert row[1] == "Engineering"  # Department is column index 1


# ---------------------------------------------------------------------------
# 2. Filter data — multiple AND conditions
# ---------------------------------------------------------------------------


class TestFilterDataMultipleConditions:
    def test_filter_data_multiple_conditions(self, tmp_path: Path) -> None:
        """AND conditions → verify both conditions met."""
        fp = str(tmp_path / "filter_and.xlsx")
        _create_analysis_data(fp)

        result = filter_data_advanced(
            fp,
            "Data",
            conditions=[
                {"column": "Department", "operator": "==", "value": "Engineering"},
                {"column": "Age", "operator": ">", "value": 30},
            ],
            logic="AND",
        )

        assert result["rows"] == 2  # Charlie (35) and Eve (32)
        for row in result["data"]:
            assert row[1] == "Engineering"
            assert row[2] > 30


# ---------------------------------------------------------------------------
# 3. Filter data — greater and less than
# ---------------------------------------------------------------------------


class TestFilterDataGreaterLess:
    def test_filter_data_greater_less(self, tmp_path: Path) -> None:
        """> and < operators → verify range filtering."""
        fp = str(tmp_path / "filter_range.xlsx")
        _create_analysis_data(fp)

        result = filter_data_advanced(
            fp,
            "Data",
            conditions=[
                {"column": "Salary", "operator": ">", "value": 60000},
                {"column": "Salary", "operator": "<", "value": 100000},
            ],
            logic="AND",
        )

        for row in result["data"]:
            salary = row[3]
            assert 60000 < salary < 100000


# ---------------------------------------------------------------------------
# 4. Sort data — single column ascending
# ---------------------------------------------------------------------------


class TestSortDataSingleColumn:
    def test_sort_data_single_column(self, tmp_path: Path) -> None:
        """Sort ascending → verify order with pandas."""
        fp = str(tmp_path / "sort_single.xlsx")
        _create_analysis_data(fp)

        sort_data(fp, "Data", column="Age", ascending=True)

        df = pd.read_excel(fp, sheet_name="Data")
        ages = df["Age"].tolist()
        assert ages == sorted(ages)


# ---------------------------------------------------------------------------
# 5. Sort data — multi-column
# ---------------------------------------------------------------------------


class TestSortDataMultiColumn:
    def test_sort_data_multi_column(self, tmp_path: Path) -> None:
        """Sort by 2 columns → verify order."""
        fp = str(tmp_path / "sort_multi.xlsx")
        _create_analysis_data(fp)

        sort_data(
            fp,
            "Data",
            sort_by=[
                {"column": "Department", "ascending": True},
                {"column": "Salary", "ascending": False},
            ],
        )

        df = pd.read_excel(fp, sheet_name="Data")
        # Within each department, salaries should be descending
        for dept in df["Department"].unique():
            dept_salaries = df.loc[df["Department"] == dept, "Salary"].tolist()
            assert dept_salaries == sorted(dept_salaries, reverse=True)


# ---------------------------------------------------------------------------
# 6. Sort data — descending
# ---------------------------------------------------------------------------


class TestSortDataDescending:
    def test_sort_data_descending(self, tmp_path: Path) -> None:
        """Descending sort → verify reverse order."""
        fp = str(tmp_path / "sort_desc.xlsx")
        _create_analysis_data(fp)

        sort_data(fp, "Data", column="Salary", ascending=False)

        df = pd.read_excel(fp, sheet_name="Data")
        salaries = df["Salary"].tolist()
        assert salaries == sorted(salaries, reverse=True)


# ---------------------------------------------------------------------------
# 7. Aggregate — sum and mean
# ---------------------------------------------------------------------------


class TestAggregateSumMean:
    def test_aggregate_sum_mean(self, tmp_path: Path) -> None:
        """Aggregate with sum and mean → verify values."""
        fp = str(tmp_path / "agg_sum.xlsx")
        _create_analysis_data(fp)

        result_sum = aggregate_data(
            fp,
            "Data",
            group_by="Department",
            value_column="Salary",
            operation="sum",
        )

        result_mean = aggregate_data(
            fp,
            "Data",
            group_by="Department",
            value_column="Salary",
            operation="mean",
        )

        # Verify sum
        df = pd.read_excel(fp, sheet_name="Data")
        expected_sum = df.groupby("Department")["Salary"].sum()
        for group in result_sum["groups"]:
            dept = group["Department"]
            assert group["Salary_sum"] == expected_sum[dept]

        # Verify mean
        expected_mean = df.groupby("Department")["Salary"].mean()
        for group in result_mean["groups"]:
            dept = group["Department"]
            assert abs(group["Salary_mean"] - expected_mean[dept]) < 0.01


# ---------------------------------------------------------------------------
# 8. Aggregate — count
# ---------------------------------------------------------------------------


class TestAggregateCount:
    def test_aggregate_count(self, tmp_path: Path) -> None:
        """Count aggregation → verify group sizes."""
        fp = str(tmp_path / "agg_count.xlsx")
        _create_analysis_data(fp)

        result = aggregate_data(
            fp,
            "Data",
            group_by="Department",
            value_column="Name",
            operation="count",
        )

        groups = {g["Department"]: g["Name_count"] for g in result["groups"]}
        assert groups["Engineering"] == 3
        assert groups["Marketing"] == 3
        assert groups["Sales"] == 2


# ---------------------------------------------------------------------------
# 9. Find duplicates
# ---------------------------------------------------------------------------


class TestFindDuplicates:
    def test_find_duplicates(self, tmp_path: Path) -> None:
        """Insert dups → find_duplicates → verify identified rows."""
        fp = str(tmp_path / "dupes.xlsx")
        create_workbook(fp, sheet_names=["Data"])
        write_range(
            fp,
            "Data",
            "A1",
            [
                ["Name", "City"],
                ["Alice", "NY"],
                ["Bob", "LA"],
                ["Alice", "NY"],
                ["Charlie", "SF"],
                ["Bob", "LA"],
            ],
        )

        result = find_duplicates(fp, "Data", columns=["Name", "City"])

        assert result["count"] == 4  # 2 pairs of duplicates = 4 rows flagged
        dup_names = [row[0] for row in result["duplicates"]]
        assert "Alice" in dup_names
        assert "Bob" in dup_names
        assert "Charlie" not in dup_names


# ---------------------------------------------------------------------------
# 10. VLookup helper
# ---------------------------------------------------------------------------


class TestVlookupHelper:
    def test_vlookup_helper(self, tmp_path: Path) -> None:
        """VLookup between two files → verify matched values."""
        lookup_fp = str(tmp_path / "lookup.xlsx")
        data_fp = str(tmp_path / "data.xlsx")

        create_workbook(lookup_fp, sheet_names=["Sheet1"])
        write_range(
            lookup_fp,
            "Sheet1",
            "A1",
            [
                ["EmpID"],
                [1],
                [2],
                [99],
            ],
        )

        create_workbook(data_fp, sheet_names=["Sheet1"])
        write_range(
            data_fp,
            "Sheet1",
            "A1",
            [
                ["EmpID", "Name", "Salary"],
                [1, "Alice", 90000],
                [2, "Bob", 75000],
                [3, "Charlie", 60000],
            ],
        )

        result = vlookup_helper(
            lookup_file=lookup_fp,
            data_file=data_fp,
            lookup_column="EmpID",
            data_key_column="EmpID",
            data_return_columns=["Name", "Salary"],
        )

        assert result["matched"] == 2
        assert result["unmatched"] == 1  # ID 99 has no match
        matched_results = [r for r in result["results"] if r["matched_value"] is not None]
        alice = [r for r in matched_results if r["lookup_value"] == 1][0]
        assert alice["return_values"]["Name"] == "Alice"
        assert alice["return_values"]["Salary"] == 90000


# ---------------------------------------------------------------------------
# 11. Profile data
# ---------------------------------------------------------------------------


class TestProfileData:
    def test_profile_data(self, tmp_path: Path) -> None:
        """Write mixed data → profile_data → verify stats."""
        fp = str(tmp_path / "profile.xlsx")
        _create_analysis_data(fp)

        result = profile_data(fp, sheet="Data")

        assert result["status"] == "success"
        assert result["row_count"] == 8
        assert result["column_count"] == 5

        # Find numeric column profile
        salary_profile = [c for c in result["columns"] if c["name"] == "Salary"][0]
        assert salary_profile["count"] == 8
        assert salary_profile["null_count"] == 0
        assert salary_profile["min"] is not None
        assert salary_profile["max"] is not None
        assert salary_profile["mean"] is not None
        assert salary_profile["p25"] is not None
        assert salary_profile["p75"] is not None
        assert salary_profile["iqr"] is not None

        # Find text column profile
        name_profile = [c for c in result["columns"] if c["name"] == "Name"][0]
        assert name_profile["unique_count"] == 8


# ---------------------------------------------------------------------------
# 12. Column statistics
# ---------------------------------------------------------------------------


class TestColumnStatistics:
    def test_column_statistics(self, tmp_path: Path) -> None:
        """Compute column stats → verify with numpy."""
        fp = str(tmp_path / "colstats.xlsx")
        _create_analysis_data(fp)

        result = column_statistics(fp, "Data", column="Salary")

        df = pd.read_excel(fp, sheet_name="Data")
        salaries = df["Salary"].values

        assert result.count == len(salaries)
        assert abs(result.mean - float(np.mean(salaries))) < 0.01
        assert abs(result.median - float(np.median(salaries))) < 0.01
        assert abs(result.min_val - float(np.min(salaries))) < 0.01
        assert abs(result.max_val - float(np.max(salaries))) < 0.01
        assert abs(result.sum_val - float(np.sum(salaries))) < 0.01


# ---------------------------------------------------------------------------
# 13. Insert subtotals
# ---------------------------------------------------------------------------


class TestInsertSubtotals:
    def test_insert_subtotals(self, tmp_path: Path) -> None:
        """Write grouped data → insert_subtotals → verify subtotal rows."""
        fp = str(tmp_path / "subtotals.xlsx")
        create_workbook(fp, sheet_names=["Sales"])
        write_range(
            fp,
            "Sales",
            "A1",
            [
                ["Region", "Amount"],
                ["East", 100],
                ["East", 200],
                ["West", 150],
                ["West", 250],
                ["West", 300],
            ],
        )

        result = insert_subtotals(
            fp,
            "Sales",
            group_col="Region",
            value_col="Amount",
            subtotal_func=9,  # SUM
        )

        assert result["status"] == "ok"
        assert result["groups"] == 2

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sales"]
        # Find subtotal rows by scanning column A for "Subtotal" or "Grand Total"
        subtotal_cells = []
        for row in ws.iter_rows(min_col=1, max_col=1):
            val = row[0].value
            if val in ("Subtotal", "Grand Total"):
                subtotal_cells.append(val)
        assert "Subtotal" in subtotal_cells
        assert "Grand Total" in subtotal_cells
        wb.close()


# ---------------------------------------------------------------------------
# 14. Filter → sort → aggregate chained workflow
# ---------------------------------------------------------------------------


class TestFilterThenSortThenAggregate:
    def test_filter_then_sort_then_aggregate(self, tmp_path: Path) -> None:
        """Chain: filter → sort → aggregate → verify end result."""
        fp = str(tmp_path / "chain.xlsx")
        _create_analysis_data(fp)

        # Step 1: Filter to salary > 60000
        filter_data_advanced(
            fp,
            "Data",
            conditions=[{"column": "Salary", "operator": ">", "value": 60000}],
            output_sheet="Filtered",
        )

        # Step 2: Sort the filtered data by department
        sort_data(fp, "Filtered", column="Department", ascending=True)

        # Step 3: Aggregate the sorted/filtered data
        result = aggregate_data(
            fp,
            "Filtered",
            group_by="Department",
            value_column="Salary",
            operation="mean",
        )

        # Verify the chain produced correct results
        groups = {g["Department"]: g["Salary_mean"] for g in result["groups"]}
        # Only employees with salary > 60000 should be in results
        assert "Engineering" in groups  # All 3 engineers > 60000
        assert "Sales" in groups  # Frank=75000
        # Diana(62000) and Hank(70000) from Marketing qualify
        assert "Marketing" in groups

        # Verify Engineering mean: (95000+110000+88000)/3
        assert abs(groups["Engineering"] - 97666.67) < 1


# ---------------------------------------------------------------------------
# 15. Analysis on empty/minimal data
# ---------------------------------------------------------------------------


class TestAnalysisOnEmptyData:
    def test_filter_empty_result(self, tmp_path: Path) -> None:
        """Filter that matches no rows → verify empty result."""
        fp = str(tmp_path / "empty_filter.xlsx")
        _create_analysis_data(fp)

        result = filter_data_advanced(
            fp,
            "Data",
            conditions=[{"column": "Salary", "operator": ">", "value": 999999}],
        )

        assert result["rows"] == 0
        assert result["data"] == []

    def test_aggregate_single_group(self, tmp_path: Path) -> None:
        """Single-value dataset → aggregate returns one group."""
        fp = str(tmp_path / "single.xlsx")
        create_workbook(fp, sheet_names=["Data"])
        write_range(
            fp,
            "Data",
            "A1",
            [
                ["Category", "Value"],
                ["A", 100],
            ],
        )

        result = aggregate_data(
            fp,
            "Data",
            group_by="Category",
            value_column="Value",
            operation="sum",
        )

        assert len(result["groups"]) == 1
        assert result["groups"][0]["Value_sum"] == 100

    def test_find_duplicates_no_dupes(self, tmp_path: Path) -> None:
        """All unique rows → find_duplicates returns count 0."""
        fp = str(tmp_path / "no_dupes.xlsx")
        create_workbook(fp, sheet_names=["Data"])
        write_range(
            fp,
            "Data",
            "A1",
            [
                ["Name", "Value"],
                ["Alice", 1],
                ["Bob", 2],
                ["Charlie", 3],
            ],
        )

        result = find_duplicates(fp, "Data", columns=["Name"])

        assert result["count"] == 0
        assert result["duplicates"] == []


# ---------------------------------------------------------------------------
# 16. Filter with OR logic
# ---------------------------------------------------------------------------


class TestFilterDataOrLogic:
    def test_filter_data_or_logic(self, tmp_path: Path) -> None:
        """OR conditions → verify either condition matched."""
        fp = str(tmp_path / "filter_or.xlsx")
        _create_analysis_data(fp)

        result = filter_data_advanced(
            fp,
            "Data",
            conditions=[
                {"column": "Department", "operator": "==", "value": "Engineering"},
                {"column": "Department", "operator": "==", "value": "Sales"},
            ],
            logic="OR",
        )

        assert result["rows"] == 5  # 3 Engineering + 2 Sales
        depts = {row[1] for row in result["data"]}
        assert depts == {"Engineering", "Sales"}


# ---------------------------------------------------------------------------
# 17. Filter with output_sheet
# ---------------------------------------------------------------------------


class TestFilterDataOutputSheet:
    def test_filter_data_output_sheet(self, tmp_path: Path) -> None:
        """Filter with output_sheet → verify written sheet matches."""
        fp = str(tmp_path / "filter_output.xlsx")
        _create_analysis_data(fp)

        filter_data_advanced(
            fp,
            "Data",
            conditions=[{"column": "City", "operator": "==", "value": "Chicago"}],
            output_sheet="ChicagoOnly",
        )

        df = pd.read_excel(fp, sheet_name="ChicagoOnly")
        assert len(df) == 3  # Bob, Diana, Hank
        assert all(df["City"] == "Chicago")


# ---------------------------------------------------------------------------
# 18. Aggregate with multiple group-by columns
# ---------------------------------------------------------------------------


class TestAggregateMultiGroupBy:
    def test_aggregate_multi_group_by(self, tmp_path: Path) -> None:
        """Group by 2 columns → verify correct grouping."""
        fp = str(tmp_path / "agg_multi.xlsx")
        _create_analysis_data(fp)

        result = aggregate_data(
            fp,
            "Data",
            group_by=["Department", "City"],
            value_column="Salary",
            operation="sum",
        )

        groups = result["groups"]
        # Engineering+New York: Alice(95000) + Eve(88000) = 183000
        eng_ny = [g for g in groups if g["Department"] == "Engineering" and g["City"] == "New York"]
        assert len(eng_ny) == 1
        assert eng_ny[0]["Salary_sum"] == 183000
