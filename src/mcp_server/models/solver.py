from __future__ import annotations

from typing import TypedDict


class SolverResult(TypedDict):
    """Result returned by run_solver."""

    found_values: dict[str, float]
    objective_value: float
    converged: bool
    iterations: int
    message: str
