from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from mcp_server.tools.financial import (
    budget_variance_analysis,
    calculate_irr,
    calculate_npv,
    calculate_pmt,
    dcf_analysis,
    financial_ratio_analysis,
    goal_seek,
    loan_amortization,
    scenario_analysis,
    trend_analysis,
)


def test_calculate_npv() -> None:
    result = calculate_npv(0.1, [-1000, 300, 400, 500])
    assert result["npv"] == pytest.approx(-21.04, abs=1)
    assert result["discount_rate"] == 0.1


def test_calculate_irr() -> None:
    result = calculate_irr([-1000, 300, 400, 500])
    assert result["irr"] is not None
    assert result["irr"] == pytest.approx(0.089, abs=0.01)


def test_calculate_pmt() -> None:
    result = calculate_pmt(0.05 / 12, 360, 200000)
    assert result["payment"] == pytest.approx(-1073.64, abs=1)


def test_goal_seek() -> None:
    result = goal_seek("x**2", target_value=9.0, initial_guess=2.0)
    assert result["converged"] is True
    assert result["result"] == pytest.approx(3.0, abs=0.01)


def test_goal_seek_unsafe_expression() -> None:
    with pytest.raises(ValueError, match="Unsafe"):
        goal_seek("__import__('os').system('echo hi')", target_value=0)


def test_loan_amortization() -> None:
    result = loan_amortization(principal=100000, annual_rate=0.06, years=30)
    assert result["monthly_payment"] > 0
    assert result["total_periods"] == 360
    assert result["total_interest"] > 0
    assert len(result["schedule"]) == 24  # capped at 24 periods shown
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


def test_scenario_analysis() -> None:
    base = {"revenue": 1000, "costs": 600, "growth_rate": 0.05}
    scenarios = [
        {"name": "Bull", "adjustments": {"revenue": 1200, "growth_rate": 0.08}},
        {"name": "Bear", "adjustments": {"revenue": 800, "growth_rate": 0.02}},
    ]
    result = scenario_analysis(
        base_case=base,
        scenarios=scenarios,
        formula="(revenue - costs) * (1 + growth_rate)",
    )
    assert len(result["base_case"]["results"]) == 1
    assert result["base_case"]["results"][0] == pytest.approx(420.0, abs=1)
    assert len(result["scenarios"]) == 2
    assert result["scenarios"][0]["name"] == "Bull"


def test_scenario_analysis_unsafe_formula() -> None:
    with pytest.raises(ValueError, match="Unsafe"):
        scenario_analysis(
            base_case={"x": 1},
            scenarios=[],
            formula="__import__('os').system('echo hi')",
        )


def test_trend_analysis(tmp_path: Path) -> None:
    fp = str(tmp_path / "trends.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Period", "Revenue"])
    for i, val in enumerate([100, 120, 140, 160, 180, 200], 1):
        ws.append([f"Q{i}", val])
    wb.save(fp)

    result = trend_analysis(fp, periods_to_forecast=3)
    assert result["trend_direction"] == "increasing"
    assert result["slope"] > 0
    assert result["r_squared"] > 0.95
    assert len(result["forecast"]) == 3
    assert result["forecast"][0]["value"] > 200
    assert len(result["growth_rates"]) == 6
    assert len(result["moving_averages"]) == 6
