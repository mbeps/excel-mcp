"""Schemas describing scenario storage and application results."""

from __future__ import annotations

from typing import TypedDict

from .common import ScenarioCellValue


class ScenarioInfo(TypedDict):
    """A what-if scenario stored in a workbook.

    Keys:
        name (str): Scenario name.
        description (str): Human-readable description.
        cell_values (dict[str, dict[str, ScenarioCellValue]]): Mapping sheet -> cell -> value.
    """

    name: str
    description: str
    cell_values: dict[str, dict[str, ScenarioCellValue]]


class ScenarioChangeInfo(TypedDict):
    """A single cell change applied when a scenario is activated.

    Keys:
        sheet (str): Sheet name. Required.
        cell (str): Cell reference changed. Required.
        new_value (ScenarioCellValue): New value applied.
    """

    sheet: str
    cell: str
    new_value: ScenarioCellValue


class ScenarioApplyResult(TypedDict):
    """Result returned by apply_scenario.

    Keys:
        scenario (str): Name of the applied scenario. Required.
        cells_updated (int): Number of cells updated. Required.
        changes (list[ScenarioChangeInfo]): List of applied changes.
    """

    scenario: str
    cells_updated: int
    changes: list[ScenarioChangeInfo]
