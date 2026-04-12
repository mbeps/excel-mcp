"""Schema for constrained optimisation solver results."""

from __future__ import annotations

from typing import TypedDict

from pydantic import BaseModel, Field


class VariableCellBounds(BaseModel):
    """Lower and upper bounds for a solver variable cell.

    Attributes:
        lower (float): Lower bound for the variable.
        upper (float): Upper bound for the variable.
    """

    lower: float = Field(..., description="Lower bound for the variable.")
    upper: float = Field(..., description="Upper bound for the variable.")


class SolverConstraint(BaseModel):
    """A constraint for the solver, expressed as an arithmetic inequality or equality.

    Attributes:
        expression (str): Arithmetic expression using cell refs, e.g. 'B2 + B3 <= 100'.
        type (str): Constraint type: 'ineq' (inequality, fun(x) >= 0) or 'eq' (equality, fun(x) == 0).
    """

    expression: str = Field(..., description="Arithmetic expression using cell refs, e.g. 'B2 + B3 <= 100'.")
    type: str = Field("ineq", description="Constraint type: 'ineq' (inequality) or 'eq' (equality).")


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
