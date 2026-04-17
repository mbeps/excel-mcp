"""Workflow tests for financial calculations, goal seek, solver, and analytics.

Tests call tool functions directly and verify results independently using
openpyxl, numpy-financial, or plain arithmetic — never via MCP tools.
"""

from __future__ import annotations

from pathlib import Path

import numpy_financial as npf
import openpyxl
import pytest

from mcp_server.tools.cell_ops import write_range
from mcp_server.tools.financial import (
    break_even_analysis,
    budget_variance_analysis,
    calculate_depreciation,
    calculate_fv,
    calculate_nper,
    calculate_pv,
    calculate_rate,
    create_sensitivity_table,
    dcf_analysis,
    financial_ratio_analysis,
    goal_seek,
    loan_amortization,
)
from mcp_server.tools.solver import run_solver
from mcp_server.tools.workbook import create_workbook

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _create_empty(fp: str, sheet: str = "Sheet1") -> None:
    create_workbook(fp, sheet_names=[sheet])


def _create_budget_data(fp: str) -> None:
    create_workbook(fp, sheet_names=["Sheet1"])
    write_range(
        fp,
        "Sheet1",
        "A1",
        [
            ["Category", "Budget", "Actual"],
            ["Rent", 5000, 5000],
            ["Marketing", 3000, 3500],
            ["Salaries", 20000, 19500],
            ["Utilities", 1500, 1800],
            ["Travel", 2000, 1200],
        ],
    )


# ---------------------------------------------------------------------------
# 1. Time value — future value
# ---------------------------------------------------------------------------


class TestTimeValueFV:
    def test_time_value_fv(self) -> None:
        """calculate_fv → verify future value matches numpy-financial."""
        result = calculate_fv(rate=0.05, nper=10, pmt=-200, pv=-1000)
        expected = float(npf.fv(0.05, 10, -200, -1000, when=0))
        assert "fv" in result
        assert abs(result["fv"] - round(expected, 4)) < 0.01


# ---------------------------------------------------------------------------
# 2. Time value — present value
# ---------------------------------------------------------------------------


class TestTimeValuePV:
    def test_time_value_pv(self) -> None:
        """calculate_pv → verify present value matches numpy-financial."""
        result = calculate_pv(rate=0.08, nper=5, pmt=-500, fv=0.0)
        expected = float(npf.pv(0.08, 5, -500, 0, when=0))
        assert "pv" in result
        assert abs(result["pv"] - round(expected, 4)) < 0.01


# ---------------------------------------------------------------------------
# 3. Time value — number of periods
# ---------------------------------------------------------------------------


class TestTimeValueNper:
    def test_time_value_nper(self) -> None:
        """calculate_nper → verify number of periods."""
        result = calculate_nper(rate=0.005, pmt=-500, pv=10000, fv=0.0)
        expected = float(npf.nper(0.005, -500, 10000, 0, when=0))
        assert "nper" in result
        assert abs(result["nper"] - round(expected, 4)) < 0.1


# ---------------------------------------------------------------------------
# 4. Time value — rate
# ---------------------------------------------------------------------------


class TestTimeValueRate:
    def test_time_value_rate(self) -> None:
        """calculate_rate → verify no NameError and reasonable rate."""
        result = calculate_rate(nper=10, pmt=-500, pv=3000, fv=0.0)
        assert "rate" in result
        assert isinstance(result["rate"], float)
        assert 0 < result["rate"] < 1  # reasonable periodic rate


# ---------------------------------------------------------------------------
# 5. Depreciation — straight line
# ---------------------------------------------------------------------------


class TestDepreciationSLN:
    def test_time_value_depreciation_sln(self) -> None:
        """calculate_depreciation(method='sln') → verify annual depreciation."""
        result = calculate_depreciation(cost=10000, salvage=2000, life=5, method="sln")
        assert result["depreciation"] == round((10000 - 2000) / 5, 2)
        assert result["method"] == "sln"


# ---------------------------------------------------------------------------
# 6. Depreciation — sum-of-years-digits
# ---------------------------------------------------------------------------


class TestDepreciationSYD:
    def test_time_value_depreciation_syd(self) -> None:
        """calculate_depreciation(method='syd', period=1) → verify first-year depreciation."""
        result = calculate_depreciation(cost=10000, salvage=2000, life=5, method="syd", period=1)
        syd_sum = 5 * 6 / 2  # 15
        expected = (10000 - 2000) * 5 / syd_sum
        assert abs(result["depreciation"] - round(expected, 2)) < 0.01
        assert result["method"] == "syd"


# ---------------------------------------------------------------------------
# 7. Goal seek — simple
# ---------------------------------------------------------------------------


class TestGoalSeekSimple:
    def test_goal_seek_simple(self, tmp_path: Path) -> None:
        """goal_seek with simple expression → verify convergence."""
        fp = str(tmp_path / "gs.xlsx")
        _create_empty(fp)
        result = goal_seek(
            file_path=fp,
            sheet_name="Sheet1",
            variable_cell="A1",
            expression="x * 2 + 5",
            target_value=25.0,
            initial_value=1.0,
        )
        assert result["converged"] is True
        assert abs(result["found_value"] - 10.0) < 0.001
        assert abs(result["achieved_result"] - 25.0) < 0.001

        wb = openpyxl.load_workbook(fp)
        assert abs(wb["Sheet1"]["A1"].value - 10.0) < 0.001
        wb.close()


# ---------------------------------------------------------------------------
# 8. Goal seek — error message on invalid expression
# ---------------------------------------------------------------------------


class TestGoalSeekError:
    def test_goal_seek_error_message(self, tmp_path: Path) -> None:
        """Invalid expression variable → verify descriptive ValueError."""
        fp = str(tmp_path / "gs_err.xlsx")
        _create_empty(fp)
        with pytest.raises(ValueError, match="[Ii]nvalid|not allowed|forbidden"):
            goal_seek(
                file_path=fp,
                sheet_name="Sheet1",
                variable_cell="A1",
                expression="y * 2 + 5",
                target_value=25.0,
            )


# ---------------------------------------------------------------------------
# 9. Loan amortization
# ---------------------------------------------------------------------------


class TestLoanAmortization:
    def test_loan_amortization(self) -> None:
        """Compute schedule → verify payment, total interest, and balance trajectory."""
        result = loan_amortization(principal=100000, annual_rate=0.06, years=30)
        assert "monthly_payment" in result
        assert "schedule" in result
        assert result["total_periods"] == 360

        expected_pmt = float(-npf.pmt(0.06 / 12, 360, 100000))
        assert abs(result["monthly_payment"] - round(expected_pmt, 2)) < 0.02

        last_entry = result["schedule"][-1]
        assert last_entry["balance"] == 0.0


# ---------------------------------------------------------------------------
# 10. DCF analysis — keys exist
# ---------------------------------------------------------------------------


class TestDCFAnalysis:
    def test_dcf_analysis(self) -> None:
        """DCF → verify npv, terminal_value, irr keys exist."""
        result = dcf_analysis(
            cash_flows=[50000, 60000, 70000, 80000, 90000],
            discount_rate=0.10,
            terminal_growth_rate=0.03,
            initial_investment=200000,
        )
        assert "total_pv" in result
        assert "terminal_value" in result
        assert "irr" in result
        assert "enterprise_value" in result
        assert "net_value" in result


# ---------------------------------------------------------------------------
# 11. DCF analysis — IRR value reasonable
# ---------------------------------------------------------------------------


class TestDCFAnalysisIRR:
    def test_dcf_analysis_irr_value(self) -> None:
        """Verify IRR is reasonable (positive for positive net cash flows)."""
        result = dcf_analysis(
            cash_flows=[50000, 60000, 70000, 80000, 90000],
            discount_rate=0.10,
            terminal_growth_rate=0.03,
            initial_investment=200000,
        )
        assert result["irr"] is not None
        assert result["irr"] > 0  # positive cash flows should yield positive IRR


# ---------------------------------------------------------------------------
# 12. Financial ratio analysis
# ---------------------------------------------------------------------------


class TestFinancialRatioAnalysis:
    def test_financial_ratio_analysis(self) -> None:
        """Pass all inputs → verify 6+ ratios computed."""
        data = {
            "current_assets": 500000,
            "current_liabilities": 200000,
            "total_debt": 300000,
            "total_equity": 400000,
            "net_income": 80000,
            "total_assets": 900000,
            "revenue": 1000000,
            "gross_profit": 400000,
            "operating_income": 150000,
            "ebitda": 200000,
            "interest_expense": 30000,
        }
        result = financial_ratio_analysis(financial_data=data)
        ratios = result["ratios"]
        assert len(ratios) >= 6

        assert abs(ratios["current_ratio"]["value"] - 2.5) < 0.01
        assert abs(ratios["debt_to_equity"]["value"] - 0.75) < 0.01
        assert abs(ratios["gross_margin"]["value"] - 0.4) < 0.01


# ---------------------------------------------------------------------------
# 13. Sensitivity table
# ---------------------------------------------------------------------------


class TestSensitivityTable:
    def test_sensitivity_table(self, tmp_path: Path) -> None:
        """Create sensitivity table → verify with openpyxl."""
        fp = str(tmp_path / "sens.xlsx")
        _create_empty(fp)
        result = create_sensitivity_table(
            file_path=fp,
            sheet_name="Sheet1",
            output_cell="A1",
            expression="price * quantity",
            var1_name="price",
            var1_values=[10.0, 20.0, 30.0],
            var2_name="quantity",
            var2_values=[100.0, 200.0],
        )
        assert result["table"] is not None
        assert len(result["table"]) == 2  # two rows for var2
        assert len(result["table"][0]) == 3  # three cols for var1

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        # Header row: var1 values in B1, C1, D1
        assert ws["B1"].value == 10.0
        assert ws["C1"].value == 20.0
        assert ws["D1"].value == 30.0
        # First data row: quantity=100, price=10 → 1000
        assert ws["B2"].value == pytest.approx(1000.0, abs=0.1)
        # Second data row: quantity=200, price=30 → 6000
        assert ws["D3"].value == pytest.approx(6000.0, abs=0.1)
        wb.close()


# ---------------------------------------------------------------------------
# 14. Budget variance analysis
# ---------------------------------------------------------------------------


class TestBudgetVarianceAnalysis:
    def test_budget_variance_analysis(self, tmp_path: Path) -> None:
        """Budget vs actual → verify variances."""
        fp = str(tmp_path / "budget.xlsx")
        _create_budget_data(fp)
        result = budget_variance_analysis(file_path=fp)
        assert "items" in result
        assert "summary" in result
        assert len(result["items"]) == 5

        marketing = next(i for i in result["items"] if i["category"] == "Marketing")
        assert marketing["variance"] == 500.0
        assert marketing["status"] == "over_budget"

        travel = next(i for i in result["items"] if i["category"] == "Travel")
        assert travel["variance"] == -800.0
        assert travel["status"] == "under_budget"


# ---------------------------------------------------------------------------
# 15. Break-even analysis
# ---------------------------------------------------------------------------


class TestBreakEvenAnalysis:
    def test_break_even_analysis(self) -> None:
        """Compute break-even → verify point."""
        result = break_even_analysis(
            fixed_costs=50000,
            price_per_unit=25.0,
            variable_cost_per_unit=15.0,
        )
        assert result["break_even_units"] == 5000.0
        assert result["break_even_revenue"] == 125000.0
        assert abs(result["contribution_margin"] - 10.0) < 0.01
        assert abs(result["contribution_margin_ratio"] - 0.4) < 0.01


# ---------------------------------------------------------------------------
# 16. Solver — minimize
# ---------------------------------------------------------------------------


class TestSolverMinimize:
    def test_solver_minimize(self, tmp_path: Path) -> None:
        """run_solver minimize → verify solution close to known answer.

        Minimize (x - 3)^2 → optimal x = 3.
        """
        fp = str(tmp_path / "solver.xlsx")
        _create_empty(fp)
        write_range(fp, "Sheet1", "A1", [["x"], [0.0]])

        result = run_solver(
            file_path=fp,
            sheet_name="Sheet1",
            objective_expression="(A2 - 3) * (A2 - 3)",
            variable_cells={"A2": (0.0, 10.0)},
            maximize=False,
        )
        assert result["converged"] is True
        assert abs(result["found_values"]["A2"] - 3.0) < 0.01
        assert abs(result["objective_value"]) < 0.01

        wb = openpyxl.load_workbook(fp)
        assert abs(wb["Sheet1"]["A2"].value - 3.0) < 0.01
        wb.close()


# ---------------------------------------------------------------------------
# 17. Solver — with constraints
# ---------------------------------------------------------------------------


class TestSolverWithConstraints:
    def test_solver_with_constraints(self, tmp_path: Path) -> None:
        """Solver with constraints → verify feasible solution.

        Maximize A2 + B2, subject to A2 + B2 <= 10, A2 >= 0, B2 >= 0.
        Optimal: A2=10, B2=0 or similar summing to 10.
        """
        fp = str(tmp_path / "solver_c.xlsx")
        _create_empty(fp)
        write_range(fp, "Sheet1", "A1", [["x", "y"], [1.0, 1.0]])

        result = run_solver(
            file_path=fp,
            sheet_name="Sheet1",
            objective_expression="A2 + B2",
            variable_cells={"A2": (0.0, 100.0), "B2": (0.0, 100.0)},
            constraints=[{"expression": "A2 + B2 <= 10"}],
            maximize=True,
        )
        assert result["converged"] is True
        total = result["found_values"]["A2"] + result["found_values"]["B2"]
        assert abs(total - 10.0) < 0.1
        assert abs(result["objective_value"] - 10.0) < 0.1


# ---------------------------------------------------------------------------
# 18. End-to-end financial workflow
# ---------------------------------------------------------------------------


class TestFinancialWorkflowComplete:
    def test_financial_workflow_complete(self, tmp_path: Path) -> None:
        """DCF → ratios → sensitivity → comprehensive chain."""
        # Step 1: DCF analysis
        dcf = dcf_analysis(
            cash_flows=[100000, 120000, 140000],
            discount_rate=0.12,
            terminal_growth_rate=0.03,
            initial_investment=300000,
        )
        assert dcf["enterprise_value"] > 0

        # Step 2: Financial ratio analysis using DCF output
        ratios_result = financial_ratio_analysis(
            financial_data={
                "current_assets": dcf["enterprise_value"],
                "current_liabilities": 300000,
                "net_income": dcf["net_value"],
                "total_equity": dcf["enterprise_value"] * 0.6,
                "total_assets": dcf["enterprise_value"],
                "revenue": sum([100000, 120000, 140000]),
                "gross_profit": sum([100000, 120000, 140000]) * 0.4,
            }
        )
        assert len(ratios_result["ratios"]) >= 4

        # Step 3: Sensitivity table
        fp = str(tmp_path / "workflow.xlsx")
        _create_empty(fp)
        sens = create_sensitivity_table(
            file_path=fp,
            sheet_name="Sheet1",
            output_cell="A1",
            expression="rate * amount",
            var1_name="rate",
            var1_values=[0.08, 0.10, 0.12, 0.15],
            var2_name="amount",
            var2_values=[100000, 200000, 300000],
        )
        assert len(sens["table"]) == 3
        assert len(sens["table"][0]) == 4

        # Verify with openpyxl
        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        # rate=0.10, amount=200000 → cell C3
        assert ws["C3"].value == pytest.approx(0.10 * 200000, abs=1.0)
        wb.close()
