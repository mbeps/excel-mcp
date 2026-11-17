"""Named range operations: list, create, delete, update."""

from __future__ import annotations

from logging import Logger

from openpyxl.workbook.defined_name import DefinedName

from mcp_server.models.named_ranges import NamedRangeInfo
from mcp_server.utils.excel_helpers import (
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def _quote_destination_sheet(destination: str) -> str:
    """Auto-quote the sheet name in a destination like 'My Sheet!$A$1' if needed."""
    if "!" not in destination:
        return destination
    sheet_part, ref_part = destination.split("!", 1)
    if sheet_part.startswith("'") and sheet_part.endswith("'"):
        return destination
    if not sheet_part.replace("_", "").isalnum() or (sheet_part and sheet_part[0].isdigit()):
        escaped = sheet_part.replace("'", "''")
        return f"'{escaped}'!{ref_part}"
    return destination


def list_named_ranges(file_path: str) -> list[NamedRangeInfo]:
    """List all named ranges in a workbook with name, destination, and scope."""
    wb = load_workbook_safe(file_path, read_only=True)
    try:
        results: list[NamedRangeInfo] = []
        # Global named ranges
        for name, defn in wb.defined_names.items():
            dest = getattr(defn, "attr_text", str(defn.value))
            results.append(
                {
                    "name": name,
                    "destination": str(dest),
                    "scope": "workbook",
                }
            )

        # Sheet-scoped named ranges
        for ws in wb.worksheets:
            for name, defn in ws.defined_names.items():
                # Some sheet-scoped ranges might already be in wb.defined_names in older openpyxl?
                # But typically they are in ws.defined_names in newer versions.
                dest = getattr(defn, "attr_text", str(defn.value))
                results.append(
                    {
                        "name": name,
                        "destination": str(dest),
                        "scope": ws.title,
                    }
                )
        return results
    finally:
        wb.close()


def create_named_range(file_path: str, name: str, destination: str, scope: str = "workbook") -> str:
    """Create a named range. destination e.g. 'Sheet1!$A$1:$A$10'."""
    destination = _quote_destination_sheet(destination)
    wb = load_workbook_safe(file_path)
    try:
        local_sheet_id = None
        if scope != "workbook":
            if scope not in wb.sheetnames:
                raise ValueError(f"Sheet '{scope}' not found. Available: {wb.sheetnames}")
            local_sheet_id = wb.sheetnames.index(scope)

        defn = DefinedName(name, attr_text=destination)
        if local_sheet_id is not None:
            defn.localSheetId = local_sheet_id
            wb[scope].defined_names.add(defn)
        else:
            wb.defined_names.add(defn)

        save_workbook_safe(wb, file_path)
        logger.info("Created named range '%s' -> %s in %s", name, destination, file_path)
        return f"Named range '{name}' created with destination '{destination}'."
    finally:
        wb.close()


def delete_named_range(file_path: str, name: str) -> str:
    """Delete a named range by name (searches workbook and sheet scopes)."""
    wb = load_workbook_safe(file_path)
    try:
        if name in wb.defined_names:
            del wb.defined_names[name]
        else:
            found = False
            for ws in wb.worksheets:
                if name in ws.defined_names:
                    del ws.defined_names[name]
                    found = True
                    break
            if not found:
                raise ValueError(f"Named range '{name}' not found.")
        save_workbook_safe(wb, file_path)
        logger.info("Deleted named range '%s' from %s", name, file_path)
        return f"Named range '{name}' deleted."
    finally:
        wb.close()


def update_named_range(file_path: str, name: str, new_destination: str) -> str:
    """Update the destination of an existing named range, preserving scope."""
    new_destination = _quote_destination_sheet(new_destination)
    wb = load_workbook_safe(file_path)
    try:
        container = None
        if name in wb.defined_names:
            container = wb.defined_names
        else:
            for ws in wb.worksheets:
                if name in ws.defined_names:
                    container = ws.defined_names
                    break
        if container is None:
            raise ValueError(f"Named range '{name}' not found.")

        old_defn = container[name]
        local_sheet_id = old_defn.localSheetId

        del container[name]
        new_defn = DefinedName(name, attr_text=new_destination)
        if local_sheet_id is not None:
            new_defn.localSheetId = local_sheet_id
        container.add(new_defn)

        save_workbook_safe(wb, file_path)
        logger.info("Updated named range '%s' -> %s in %s", name, new_destination, file_path)
        return f"Named range '{name}' updated to '{new_destination}'."
    finally:
        wb.close()
