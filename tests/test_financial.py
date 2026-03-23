from __future__ import annotations

import pytest

from mcp_server.tools.financial import (
    calculate_irr,
    calculate_npv,
    calculate_pmt,
    goal_seek,
    loan_amortization,
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
