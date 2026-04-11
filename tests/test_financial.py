from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from mcp_server.tools.financial import (
    break_even_analysis,
    budget_variance_analysis,
    dcf_analysis,
    financial_ratio_analysis,
    goal_seek,
    loan_amortization,
)


def test_goal_seek(sample_xlsx):
    result = goal_seek(
        file_path=str(sample_xlsx),
        sheet_name="Sheet1",
        variable_cell="A1",
        expression="x * 2",
        target_value=10.0,
        initial_value=1.0,
    )
    assert result["converged"] is True
    assert abs(result["found_value"] - 5.0) < 0.001


def test_loan_amortization() -> None:
    result = loan_amortization(principal=100000, annual_rate=0.06, years=30)
    assert result["monthly_payment"] > 0
    assert result["total_periods"] == 360
    assert result["total_interest"] > 0
    assert len(result["schedule"]) == 360
    assert result["schedule"][0]["period"] == 1
    assert result["schedule"][0]["balance"] < 100000


def test_dcf_analysis() -> None:
    result = dcf_analysis(
        cash_flows=[100, 150, 200],
        discount_rate=0.10,
        terminal_growth_rate=0.02,
        initial_investment=500,
    )
    assert len(result["pv_cash_flows"]) == 3
    assert result["total_pv"] > 0
    assert result["terminal_value"] > 0
    assert result["pv_terminal_value"] > 0
    assert result["enterprise_value"] == pytest.approx(result["total_pv"] + result["pv_terminal_value"], abs=0.01)
    assert result["net_value"] == pytest.approx(result["enterprise_value"] - 500, abs=0.01)


def test_dcf_analysis_invalid_rates() -> None:
    with pytest.raises(ValueError, match="greater than"):
        dcf_analysis(cash_flows=[100], discount_rate=0.02, terminal_growth_rate=0.05)


def test_budget_variance_analysis(tmp_path: Path) -> None:
    fp = str(tmp_path / "budget.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Category", "Budget", "Actual"])
    ws.append(["Marketing", 10000, 12000])
    ws.append(["Engineering", 50000, 48000])
    ws.append(["Sales", 20000, 20050])
    wb.save(fp)

    result = budget_variance_analysis(fp)
    items = result["items"]
    assert len(items) == 3
    assert items[0]["status"] == "over_budget"
    assert items[1]["status"] == "under_budget"
    assert result["summary"]["total_budget"] == 80000
    assert result["summary"]["total_actual"] == 80050


def test_budget_variance_analysis_output(tmp_path: Path) -> None:
    fp = str(tmp_path / "budget.xlsx")
    out = str(tmp_path / "report.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Category", "Budget", "Actual"])
    ws.append(["R&D", 30000, 35000])
    wb.save(fp)

    result = budget_variance_analysis(fp, output_file=out)
    assert result["output_file"] == out
    assert Path(out).exists()


def test_financial_ratio_analysis() -> None:
    data = {
        "current_assets": 500000,
        "current_liabilities": 250000,
        "net_income": 100000,
        "total_equity": 400000,
        "revenue": 1000000,
        "gross_profit": 600000,
    }
    result = financial_ratio_analysis(data)
    assert result["ratios"]["current_ratio"]["value"] == pytest.approx(2.0)
    assert result["ratios"]["roe"]["value"] == pytest.approx(0.25)
    assert result["ratios"]["gross_margin"]["value"] == pytest.approx(0.6)
    assert result["ratios"]["net_margin"]["value"] == pytest.approx(0.1)


def test_financial_ratio_analysis_with_benchmarks() -> None:
    data = {"current_assets": 300000, "current_liabilities": 200000}
    benchmarks = {"current_ratio": 2.0}
    result = financial_ratio_analysis(data, industry_benchmarks=benchmarks)
    cr = result["ratios"]["current_ratio"]
    assert cr["value"] == pytest.approx(1.5)
    assert cr["comparison"] == "below_benchmark"


def test_break_even_analysis() -> None:
    result = break_even_analysis(fixed_costs=50000.0, price_per_unit=25.0, variable_cost_per_unit=15.0)
    assert result["break_even_units"] == pytest.approx(5000.0, abs=0.01)
    assert result["break_even_revenue"] == pytest.approx(125000.0, abs=0.01)
    assert result["contribution_margin"] == pytest.approx(10.0)
    assert result["contribution_margin_ratio"] == pytest.approx(0.4, abs=0.0001)


def test_break_even_analysis_invalid() -> None:
    with pytest.raises(ValueError, match="price_per_unit must be greater than variable_cost_per_unit"):
        break_even_analysis(fixed_costs=1000.0, price_per_unit=10.0, variable_cost_per_unit=10.0)


def test_loan_amortization_full() -> None:
    result = loan_amortization(principal=10000, annual_rate=0.12, years=1)
    assert result["total_periods"] == 12
    assert len(result["schedule"]) == 12
    assert result.get("truncated") is None


def test_loan_amortization_truncated() -> None:
    result = loan_amortization(principal=10000, annual_rate=0.12, years=1, max_periods=3)
    assert len(result["schedule"]) == 3
    assert result["truncated"] is True


def test_break_even_invalid_inputs() -> None:
    with pytest.raises(ValueError, match="fixed_costs must be non-negative"):
        break_even_analysis(fixed_costs=-1000.0, price_per_unit=25.0, variable_cost_per_unit=15.0)
    with pytest.raises(ValueError, match="price_per_unit must be positive"):
        break_even_analysis(fixed_costs=1000.0, price_per_unit=0.0, variable_cost_per_unit=15.0)
    with pytest.raises(ValueError, match="price_per_unit must be greater than variable_cost_per_unit"):
        break_even_analysis(fixed_costs=1000.0, price_per_unit=10.0, variable_cost_per_unit=10.0)


# ── comprehensive financial_ratio_analysis coverage ─────────────────────────


class TestFinancialRatioAnalysisComprehensive:
    """Thorough coverage of all 7 ratios, partial keys, division-by-zero, and benchmarks."""

    def test_basic_current_ratio_two_to_one(self) -> None:
        """current_assets=200, current_liabilities=100 → current_ratio=2.0."""
        result = financial_ratio_analysis({"current_assets": 200, "current_liabilities": 100})
        assert result["ratios"]["current_ratio"]["value"] == pytest.approx(2.0)

    def test_all_seven_ratios_present(self) -> None:
        """All 10 ratios are computed when all input keys are supplied."""
        data = {
            "current_assets": 400,
            "current_liabilities": 200,
            "inventory": 50,
            "total_debt": 300,
            "total_equity": 600,
            "net_income": 90,
            "total_assets": 900,
            "revenue": 1000,
            "gross_profit": 500,
            "operating_income": 200,
            "ebitda": 150,
            "interest_expense": 30,
        }
        result = financial_ratio_analysis(data)
        r = result["ratios"]
        assert set(r.keys()) == {
            "current_ratio",
            "quick_ratio",
            "debt_to_equity",
            "roe",
            "roa",
            "gross_margin",
            "operating_margin",
            "net_margin",
            "asset_turnover",
            "interest_coverage",
        }
        assert len(r) == 10

    def test_all_seven_ratio_values_correct(self) -> None:
        """Verify numerical accuracy of all 7 ratios simultaneously."""
        data = {
            "current_assets": 400,
            "current_liabilities": 200,
            "total_debt": 300,
            "total_equity": 600,
            "net_income": 90,
            "total_assets": 900,
            "revenue": 1000,
            "gross_profit": 500,
            "ebitda": 150,
            "interest_expense": 30,
        }
        result = financial_ratio_analysis(data)
        r = result["ratios"]
        assert r["current_ratio"]["value"] == pytest.approx(2.0)
        assert r["debt_to_equity"]["value"] == pytest.approx(0.5)
        assert r["roe"]["value"] == pytest.approx(0.15)
        assert r["roa"]["value"] == pytest.approx(0.1)
        assert r["gross_margin"]["value"] == pytest.approx(0.5)
        assert r["net_margin"]["value"] == pytest.approx(0.09)
        assert r["interest_coverage"]["value"] == pytest.approx(5.0)

    def test_partial_keys_only_matching_ratios_computed(self) -> None:
        """Partial keys: only ratios with both numerator and denominator present are computed."""
        data = {"net_income": 50000, "total_assets": 500000}  # only roa computable
        result = financial_ratio_analysis(data)
        r = result["ratios"]
        assert "roa" in r
        assert r["roa"]["value"] == pytest.approx(0.1)
        assert "current_ratio" not in r
        assert "debt_to_equity" not in r
        assert "roe" not in r
        assert "gross_margin" not in r
        assert "net_margin" not in r

    def test_division_by_zero_yields_error_entry(self) -> None:
        """denominator=0 produces {"value": None, "error": "Division by zero"}."""
        data = {"current_assets": 100, "current_liabilities": 0}
        result = financial_ratio_analysis(data)
        cr = result["ratios"]["current_ratio"]
        assert cr["value"] is None
        assert "error" in cr
        assert "zero" in cr["error"].lower()

    def test_division_by_zero_multiple_ratios(self) -> None:
        """Multiple zero denominators each get their own error entry."""
        data = {
            "current_assets": 100,
            "current_liabilities": 0,
            "total_debt": 200,
            "total_equity": 0,
        }
        result = financial_ratio_analysis(data)
        assert result["ratios"]["current_ratio"]["value"] is None
        assert result["ratios"]["debt_to_equity"]["value"] is None

    def test_completely_wrong_keys_returns_empty_ratios(self) -> None:
        """No recognized keys → {"ratios": {}} — regression guard."""
        data = {"foo": 100, "bar": 200, "baz": 300}
        result = financial_ratio_analysis(data)
        assert result == {"ratios": {}}

    def test_empty_financial_data_returns_empty_ratios(self) -> None:
        """Empty dict → {"ratios": {}}."""
        result = financial_ratio_analysis({})
        assert result == {"ratios": {}}

    def test_current_ratio_isolated(self) -> None:
        """current_assets / current_liabilities computed in isolation."""
        result = financial_ratio_analysis({"current_assets": 300, "current_liabilities": 150})
        assert result["ratios"]["current_ratio"]["value"] == pytest.approx(2.0)

    def test_debt_to_equity_isolated(self) -> None:
        """total_debt / total_equity computed in isolation."""
        result = financial_ratio_analysis({"total_debt": 500, "total_equity": 250})
        assert result["ratios"]["debt_to_equity"]["value"] == pytest.approx(2.0)

    def test_roe_isolated(self) -> None:
        """net_income / total_equity = roe."""
        result = financial_ratio_analysis({"net_income": 75, "total_equity": 500})
        assert result["ratios"]["roe"]["value"] == pytest.approx(0.15)

    def test_roa_isolated(self) -> None:
        """net_income / total_assets = roa."""
        result = financial_ratio_analysis({"net_income": 50, "total_assets": 1000})
        assert result["ratios"]["roa"]["value"] == pytest.approx(0.05)

    def test_gross_margin_isolated(self) -> None:
        """gross_profit / revenue = gross_margin."""
        result = financial_ratio_analysis({"gross_profit": 300, "revenue": 500})
        assert result["ratios"]["gross_margin"]["value"] == pytest.approx(0.6)

    def test_net_margin_isolated(self) -> None:
        """net_income / revenue = net_margin."""
        result = financial_ratio_analysis({"net_income": 100, "revenue": 1000})
        assert result["ratios"]["net_margin"]["value"] == pytest.approx(0.1)

    def test_interest_coverage_isolated(self) -> None:
        """ebitda / interest_expense = interest_coverage."""
        result = financial_ratio_analysis({"ebitda": 200, "interest_expense": 40})
        assert result["ratios"]["interest_coverage"]["value"] == pytest.approx(5.0)

    def test_benchmark_above(self) -> None:
        """value > benchmark (>5%) → comparison = 'above_benchmark'."""
        data = {"current_assets": 300, "current_liabilities": 100}
        result = financial_ratio_analysis(data, industry_benchmarks={"current_ratio": 2.0})
        cr = result["ratios"]["current_ratio"]
        assert cr["value"] == pytest.approx(3.0)
        assert cr["comparison"] == "above_benchmark"
        assert cr["benchmark"] == 2.0

    def test_benchmark_below(self) -> None:
        """value < benchmark (>5% difference) → comparison = 'below_benchmark'."""
        data = {"current_assets": 100, "current_liabilities": 200}
        result = financial_ratio_analysis(data, industry_benchmarks={"current_ratio": 2.0})
        cr = result["ratios"]["current_ratio"]
        assert cr["value"] == pytest.approx(0.5)
        assert cr["comparison"] == "below_benchmark"

    def test_benchmark_at(self) -> None:
        """value == benchmark → comparison = 'at_benchmark'."""
        data = {"current_assets": 200, "current_liabilities": 100}
        result = financial_ratio_analysis(data, industry_benchmarks={"current_ratio": 2.0})
        cr = result["ratios"]["current_ratio"]
        assert cr["comparison"] == "at_benchmark"

    def test_benchmark_only_for_ratios_in_benchmarks_dict(self) -> None:
        """Benchmark comparison only added for ratios present in industry_benchmarks."""
        data = {
            "current_assets": 200,
            "current_liabilities": 100,
            "net_income": 50,
            "total_assets": 500,
        }
        result = financial_ratio_analysis(data, industry_benchmarks={"current_ratio": 2.0})
        # current_ratio has benchmark info
        assert "comparison" in result["ratios"]["current_ratio"]
        assert "benchmark" in result["ratios"]["current_ratio"]
        # roa does NOT have benchmark info (not in benchmarks dict)
        assert "comparison" not in result["ratios"]["roa"]
        assert "benchmark" not in result["ratios"]["roa"]

    def test_output_structure_each_ratio_has_value_key(self) -> None:
        """Every computed ratio entry always has a 'value' key."""
        data = {
            "current_assets": 200,
            "current_liabilities": 100,
            "total_debt": 100,
            "total_equity": 200,
        }
        result = financial_ratio_analysis(data)
        for ratio_name, entry in result["ratios"].items():
            assert "value" in entry, f"Ratio '{ratio_name}' missing 'value' key"

    def test_ratio_values_rounded_to_four_decimal_places(self) -> None:
        """Ratio values are rounded to 4 decimal places."""
        data = {"net_income": 1, "revenue": 3}  # 1/3 = 0.3333...
        result = financial_ratio_analysis(data)
        val = result["ratios"]["net_margin"]["value"]
        # Should be 0.3333 (4 decimals), not 0.33333333...
        assert val == pytest.approx(round(1 / 3, 4))
