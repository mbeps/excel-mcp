"""Scenario management: save, list, apply, and delete named parameter sets."""

from __future__ import annotations

import json
from logging import Logger

from mcp_server.utils.excel_helpers import get_sheet, load_workbook_safe, save_workbook_safe
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

_SCENARIOS_SHEET = "_mcp_scenarios"


def _load_scenarios(wb) -> dict:
    """Load scenarios from the hidden scenarios sheet."""
    if _SCENARIOS_SHEET not in wb.sheetnames:
        return {}
    ws = wb[_SCENARIOS_SHEET]
    raw = ws["A1"].value
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except (json.JSONDecodeError, TypeError):
        return {}


def _save_scenarios(wb, scenarios: dict) -> None:
    """Save scenarios dict to hidden sheet."""
    if _SCENARIOS_SHEET in wb.sheetnames:
        ws = wb[_SCENARIOS_SHEET]
    else:
        ws = wb.create_sheet(_SCENARIOS_SHEET)
        ws.sheet_state = "hidden"
    ws["A1"] = json.dumps(scenarios)


def add_scenario(
    file_path: str,
    name: str,
    cell_values: dict[str, dict[str, float | str | int | None]],
    description: str = "",
) -> str:
    """Save a named scenario (parameter set) to the workbook.

    cell_values: {sheet_name: {cell_ref: value}} mapping
    E.g. {"Assumptions": {"B2": 0.05, "B3": 1000000}}
    """
    if not name:
        raise ValueError("Scenario name must not be empty.")
    if not cell_values:
        raise ValueError("cell_values must not be empty.")

    wb = load_workbook_safe(file_path)
    try:
        for sheet_name, cells in cell_values.items():
            if not isinstance(cells, dict):
                raise ValueError(
                    f"cell_values must be {{sheet_name: {{cell_ref: value}}}}. "
                    f"Key '{sheet_name}' has a non-dict value (got {type(cells).__name__}). "
                    "Did you forget to nest by sheet name?"
                )
            if sheet_name not in wb.sheetnames:
                raise ValueError(
                    f"Sheet '{sheet_name}' not found in workbook. "
                    f"Available sheets: {wb.sheetnames}. "
                    "cell_values must be structured as {sheet_name: {cell_ref: value}}."
                )
        scenarios = _load_scenarios(wb)
        scenarios[name] = {
            "description": description,
            "cell_values": cell_values,
        }
        _save_scenarios(wb, scenarios)
        save_workbook_safe(wb, file_path)
        logger.info("Added scenario '%s' to %s", name, file_path)
        return f"Scenario '{name}' saved with {sum(len(v) for v in cell_values.values())} cell(s)."
    finally:
        wb.close()


def list_scenarios(file_path: str) -> list[dict]:
    """List all saved scenarios with their cell/value mappings."""
    wb = load_workbook_safe(file_path, read_only=True)
    try:
        scenarios = _load_scenarios(wb)
        result = [
            {
                "name": name,
                "description": data.get("description", ""),
                "cell_values": data.get("cell_values", {}),
            }
            for name, data in scenarios.items()
        ]
        logger.info("Listed %d scenarios from %s", len(result), file_path)
        return result
    finally:
        wb.close()


def apply_scenario(file_path: str, name: str) -> dict:
    """Apply a scenario by writing its cell values to the workbook."""
    wb = load_workbook_safe(file_path)
    try:
        scenarios = _load_scenarios(wb)
        if name not in scenarios:
            raise ValueError(f"Scenario '{name}' not found. Available: {list(scenarios.keys())}")

        cell_values: dict[str, dict[str, object]] = scenarios[name].get("cell_values", {})
        changes: list[dict] = []

        for sheet_name, cells in cell_values.items():
            ws = get_sheet(wb, sheet_name)
            for cell_ref, value in cells.items():
                ws[cell_ref] = value
                changes.append({"sheet": sheet_name, "cell": cell_ref, "new_value": value})

        save_workbook_safe(wb, file_path)
        logger.info("Applied scenario '%s': %d cells updated in %s", name, len(changes), file_path)
        return {
            "scenario": name,
            "cells_updated": len(changes),
            "changes": changes,
        }
    finally:
        wb.close()
