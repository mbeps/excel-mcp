"""Worksheet-level operations: freeze panes, filters, visibility, grouping, zoom, print setup."""

from __future__ import annotations

from logging import Logger

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
    validate_file_path,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def freeze_panes(file_path: str, sheet_name: str, cell_ref: str | None = None) -> str:
    """Freeze panes at cell_ref (e.g. 'B2' freezes row 1 and col A). Pass None or omit to unfreeze."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        if cell_ref is None or cell_ref == "A1":
            ws.freeze_panes = None
            save_workbook_safe(wb, file_path)
            logger.info("Unfroze panes in '%s' of %s", sheet_name, file_path)
            return f"Panes unfrozen in sheet '{sheet_name}'."
        ws.freeze_panes = cell_ref
        save_workbook_safe(wb, file_path)
        logger.info("Froze panes at %s in '%s' of %s", cell_ref, sheet_name, file_path)
        return f"Panes frozen at '{cell_ref}' in sheet '{sheet_name}'."
    finally:
        wb.close()


def set_auto_filter(file_path: str, sheet_name: str, cell_range: str | None = None, remove: bool = False) -> str:
    """Enable auto-filter on a range (e.g. 'A1:E1'). Pass remove=True to clear the filter."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        if remove:
            ws.auto_filter.ref = None
            save_workbook_safe(wb, file_path)
            logger.info("Removed auto-filter in '%s' of %s", sheet_name, file_path)
            return f"Auto-filter removed from sheet '{sheet_name}'."
        ws.auto_filter.ref = cell_range
        save_workbook_safe(wb, file_path)
        logger.info("Set auto-filter %s in '%s' of %s", cell_range, sheet_name, file_path)
        return f"Auto-filter set on '{cell_range}' in sheet '{sheet_name}'."
    finally:
        wb.close()


def copy_range_across_sheets(
    file_path: str,
    source_sheet: str,
    source_range: str,
    target_sheet: str,
    target_start_cell: str = "A1",
    copy_values: bool = True,
    copy_styles: bool = False,
) -> str:
    """Copy a range from one sheet to another within the same workbook."""
    import copy

    from openpyxl.utils import column_index_from_string
    from openpyxl.utils.cell import coordinate_from_string

    wb = load_workbook_safe(file_path)
    try:
        src_ws = get_sheet(wb, source_sheet)
        tgt_ws = get_sheet(wb, target_sheet)
        tgt_col_letter, tgt_row = coordinate_from_string(target_start_cell)
        tgt_col = column_index_from_string(tgt_col_letter)
        cells = list(src_ws[source_range])
        row_count = 0
        for row_cells in cells:
            col_offset = 0
            for cell in row_cells:
                dest_col = tgt_col + col_offset
                dest = tgt_ws.cell(row=tgt_row + row_count, column=dest_col)
                if copy_values:
                    dest.value = cell.value
                if copy_styles and cell.has_style:
                    dest.font = copy.copy(cell.font)
                    dest.fill = copy.copy(cell.fill)
                    dest.border = copy.copy(cell.border)
                    dest.alignment = copy.copy(cell.alignment)
                    dest.number_format = cell.number_format
                col_offset += 1
            row_count += 1
        save_workbook_safe(wb, file_path)
        logger.info(
            "Copied %s from '%s' to '%s'!%s",
            source_range,
            source_sheet,
            target_sheet,
            target_start_cell,
        )
        return (
            f"Copied range '{source_range}' from '{source_sheet}'"
            f" to '{target_sheet}' starting at '{target_start_cell}'"
            f" ({row_count} rows)."
        )
    finally:
        wb.close()


def copy_sheet_across_workbooks(
    source_file: str,
    source_sheet: str,
    dest_file: str,
    dest_sheet_name: str | None = None,
) -> str:
    """Copy a sheet from one workbook to another by replicating cell data and column widths."""
    import copy
    from pathlib import Path

    import openpyxl

    validate_file_path(source_file, must_exist=True)
    dest_path = Path(dest_file).resolve()
    from mcp_server.utils.excel_helpers import ALLOWED_EXTENSIONS

    if dest_path.suffix.lower() not in ALLOWED_EXTENSIONS:
        raise ValueError(f"Destination must have an allowed extension: {ALLOWED_EXTENSIONS}")
    src_wb = load_workbook_safe(source_file)
    try:
        src_ws = get_sheet(src_wb, source_sheet)
        new_name = dest_sheet_name or source_sheet
        if dest_path.exists():
            dest_wb = load_workbook_safe(dest_file)
        else:
            dest_wb = openpyxl.Workbook()
            dest_wb.remove(dest_wb.active)
        try:
            if new_name in dest_wb.sheetnames:
                raise ValueError(f"Sheet '{new_name}' already exists in '{dest_file}'. Use dest_sheet_name to rename.")
            tgt_ws = dest_wb.create_sheet(new_name)
            for row in src_ws.iter_rows():
                for cell in row:
                    dest_cell = tgt_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.has_style:
                        dest_cell.font = copy.copy(cell.font)
                        dest_cell.fill = copy.copy(cell.fill)
                        dest_cell.border = copy.copy(cell.border)
                        dest_cell.alignment = copy.copy(cell.alignment)
                        dest_cell.number_format = cell.number_format
            for col_letter, dim in src_ws.column_dimensions.items():
                tgt_ws.column_dimensions[col_letter].width = dim.width
            dest_path.parent.mkdir(parents=True, exist_ok=True)
            dest_wb.save(str(dest_path))
            logger.info("Copied sheet '%s' from %s to %s as '%s'", source_sheet, source_file, dest_file, new_name)
            return f"Sheet '{source_sheet}' copied to '{dest_file}' as '{new_name}'."
        finally:
            dest_wb.close()
    finally:
        src_wb.close()


def merge_workbooks(
    source_files: list[str],
    output_file: str,
    conflict_strategy: str = "rename",
) -> dict[str, object]:
    """Merge all sheets from multiple workbooks into a single output workbook.

    conflict_strategy: 'rename' (append _2, _3, ...) or 'overwrite'.
    """
    import copy
    from pathlib import Path

    import openpyxl

    if conflict_strategy not in ("rename", "overwrite"):
        raise ValueError("conflict_strategy must be 'rename' or 'overwrite'.")
    out_path = Path(output_file).resolve()
    from mcp_server.utils.excel_helpers import ALLOWED_EXTENSIONS

    if out_path.suffix.lower() not in ALLOWED_EXTENSIONS:
        raise ValueError(f"output_file must have an allowed extension: {ALLOWED_EXTENSIONS}")
    if out_path.exists():
        out_wb = openpyxl.load_workbook(output_file)
        sheet_names_used: list[str] = list(out_wb.sheetnames)
    else:
        out_wb = openpyxl.Workbook()
        out_wb.remove(out_wb.active)
        sheet_names_used = []
    merged_files = 0
    for src_file in source_files:
        validate_file_path(src_file, must_exist=True)
        src_wb = load_workbook_safe(src_file)
        try:
            for sheet_name in src_wb.sheetnames:
                src_ws = src_wb[sheet_name]
                target_name = sheet_name
                if target_name in sheet_names_used:
                    if conflict_strategy == "overwrite":
                        out_wb.remove(out_wb[target_name])
                        sheet_names_used.remove(target_name)
                    else:
                        counter = 2
                        while f"{sheet_name}_{counter}" in sheet_names_used:
                            counter += 1
                        target_name = f"{sheet_name}_{counter}"
                tgt_ws = out_wb.create_sheet(target_name)
                for row in src_ws.iter_rows():
                    for cell in row:
                        dest_cell = tgt_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                        if cell.has_style:
                            dest_cell.font = copy.copy(cell.font)
                            dest_cell.fill = copy.copy(cell.fill)
                            dest_cell.border = copy.copy(cell.border)
                            dest_cell.alignment = copy.copy(cell.alignment)
                            dest_cell.number_format = cell.number_format
                for col_letter, dim in src_ws.column_dimensions.items():
                    tgt_ws.column_dimensions[col_letter].width = dim.width
                sheet_names_used.append(target_name)
            merged_files += 1
        finally:
            src_wb.close()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_wb.save(str(out_path))
    out_wb.close()
    logger.info("Merged %d files into %s (%d sheets)", merged_files, output_file, len(sheet_names_used))
    return {
        "merged_files": merged_files,
        "total_sheets": len(sheet_names_used),
        "output_file": output_file,
        "sheets": sheet_names_used,
    }
