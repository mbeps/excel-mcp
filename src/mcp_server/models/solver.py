from __future__ import annotations

"""Schema for constrained optimisation solver results.
"""

from typing import TypedDict


class SolverResult(TypedDict):
    """Result returned by run_solver.

    Keys:
        found_values (dict[str, float]): Mapping from variable names to found numeric values. Required.
        objective_value (float): Value of the objective function at the solution. Required.
        converged (bool): Whether the solver converged. Required.
        iterations (int): Number of iterations performed. Required.
        message (str): Informational or error message.
    """

    found_values: dict[str, float]
    objective_value: float
    converged: bool
    iterations: int
    message: str
