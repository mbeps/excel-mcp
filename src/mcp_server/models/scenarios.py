from __future__ import annotations

from typing import TypedDict

from .common import ScenarioCellValue


class ScenarioInfo(TypedDict):
    """A what-if scenario stored in a workbook."""

    name: str
    description: str
    cell_values: dict[str, dict[str, ScenarioCellValue]]


class ScenarioChangeInfo(TypedDict):
    """A single cell change applied when a scenario is activated."""

    sheet: str
    cell: str
    new_value: ScenarioCellValue


class ScenarioApplyResult(TypedDict):
    """Result returned by apply_scenario."""

    scenario: str
    cells_updated: int
    changes: list[ScenarioChangeInfo]
