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


def insert_rows(file_path: str, sheet: str, row: int, count: int = 1) -> dict[str, str]:
    """Insert one or more rows at the given row index, shifting existing rows down."""
    validate_file_path(file_path, must_exist=True)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet)
        ws.insert_rows(row, count)
        save_workbook_safe(wb, file_path)
        logger.info("Inserted %d row(s) at row %d in '%s' of %s", count, row, sheet, file_path)
        return {"status": "success", "message": f"Inserted {count} row(s) at row {row}"}
    finally:
        wb.close()


def delete_rows(file_path: str, sheet: str, row: int, count: int = 1) -> dict[str, str]:
    """Delete one or more rows starting at the given row index."""
    validate_file_path(file_path, must_exist=True)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet)
        ws.delete_rows(row, count)
        save_workbook_safe(wb, file_path)
        logger.info("Deleted %d row(s) starting at row %d in '%s' of %s", count, row, sheet, file_path)
        return {"status": "success", "message": f"Deleted {count} row(s) starting at row {row}"}
    finally:
        wb.close()


def insert_cols(file_path: str, sheet: str, col: int, count: int = 1) -> dict[str, str]:
    """Insert one or more columns at the given column index, shifting existing columns right."""
    validate_file_path(file_path, must_exist=True)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet)
        ws.insert_cols(col, count)
        save_workbook_safe(wb, file_path)
        logger.info("Inserted %d column(s) at column %d in '%s' of %s", count, col, sheet, file_path)
        return {"status": "success", "message": f"Inserted {count} column(s) at column {col}"}
    finally:
        wb.close()


def delete_cols(file_path: str, sheet: str, col: int, count: int = 1) -> dict[str, str]:
    """Delete one or more columns starting at the given column index."""
    validate_file_path(file_path, must_exist=True)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet)
        ws.delete_cols(col, count)
        save_workbook_safe(wb, file_path)
        logger.info("Deleted %d column(s) starting at column %d in '%s' of %s", count, col, sheet, file_path)
        return {"status": "success", "message": f"Deleted {count} column(s) starting at column {col}"}
    finally:
        wb.close()


def set_print_area(file_path: str, sheet: str, print_area: str) -> dict[str, str]:
    """Set the print area for a worksheet (e.g. 'A1:H20')."""
    validate_file_path(file_path, must_exist=True)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet)
        ws.print_area = print_area
        save_workbook_safe(wb, file_path)
        logger.info("Set print area to %s in '%s' of %s", print_area, sheet, file_path)
        return {"status": "success", "message": f"Print area set to {print_area}"}
    finally:
        wb.close()


def set_page_setup(
    file_path: str,
    sheet: str,
    orientation: str = "portrait",
    paper_size: int = 1,
    fit_to_width: int | None = None,
    fit_to_height: int | None = None,
) -> dict[str, object]:
    """Configure page setup: orientation, paper size, and fit-to-page scaling."""
    validate_file_path(file_path, must_exist=True)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet)
        ws.page_setup.orientation = orientation
        ws.page_setup.paperSize = paper_size
        if fit_to_width is not None:
            ws.page_setup.fitToWidth = fit_to_width
        if fit_to_height is not None:
            ws.page_setup.fitToHeight = fit_to_height
        if fit_to_width is not None or fit_to_height is not None:
            ws.sheet_properties.pageSetUpPr.fitToPage = True
        save_workbook_safe(wb, file_path)
        logger.info("Configured page setup in '%s' of %s", sheet, file_path)
        return {
            "status": "success",
            "message": "Page setup configured",
            "orientation": orientation,
            "paper_size": paper_size,
            "fit_to_width": fit_to_width,
            "fit_to_height": fit_to_height,
        }
    finally:
        wb.close()


def group_rows(
    file_path: str, sheet: str, start_row: int, end_row: int, outline_level: int = 1, hidden: bool = False
) -> dict[str, str]:
    """Group rows by setting their outline level and optional hidden state.

    When hidden=True, the summary row at end_row+1 gets collapsed=True so the
    collapse button renders correctly in Excel.
    """
    validate_file_path(file_path, must_exist=True)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet)
        for row in range(start_row, end_row + 1):
            ws.row_dimensions[row].outline_level = outline_level
            ws.row_dimensions[row].hidden = hidden
        if hidden:
            # Mark the summary row so Excel renders the collapse button correctly
            ws.row_dimensions[end_row + 1].collapsed = True
        save_workbook_safe(wb, file_path)
        logger.info("Grouped rows %d-%d at level %d in '%s' of %s", start_row, end_row, outline_level, sheet, file_path)
        return {"status": "success", "message": f"Grouped rows {start_row}-{end_row} at level {outline_level}"}
    finally:
        wb.close()


def group_cols(
    file_path: str, sheet: str, start_col: int, end_col: int, outline_level: int = 1, hidden: bool = False
) -> dict[str, str]:
    """Group columns by setting their outline level and optional hidden state.

    When hidden=True, the summary column at end_col+1 gets collapsed=True so the
    collapse button renders correctly in Excel.
    """
    from openpyxl.utils import get_column_letter

    validate_file_path(file_path, must_exist=True)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet)
        for col in range(start_col, end_col + 1):
            letter = get_column_letter(col)
            ws.column_dimensions[letter].outline_level = outline_level
            ws.column_dimensions[letter].hidden = hidden
        if hidden:
            # Mark the summary column so Excel renders the collapse button correctly
            summary_letter = get_column_letter(end_col + 1)
            ws.column_dimensions[summary_letter].collapsed = True
        save_workbook_safe(wb, file_path)
        logger.info(
            "Grouped columns %d-%d at level %d in '%s' of %s",
            start_col,
            end_col,
            outline_level,
            sheet,
            file_path,
        )
        return {"status": "success", "message": f"Grouped columns {start_col}-{end_col} at level {outline_level}"}
    finally:
        wb.close()


def ungroup_rows(file_path: str, sheet: str, start_row: int, end_row: int) -> dict[str, str]:
    """Ungroup rows by resetting their outline level to 0 and unhiding them."""
    validate_file_path(file_path, must_exist=True)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet)
        for row in range(start_row, end_row + 1):
            ws.row_dimensions[row].outline_level = 0
            ws.row_dimensions[row].hidden = False
        save_workbook_safe(wb, file_path)
        logger.info("Ungrouped rows %d-%d in '%s' of %s", start_row, end_row, sheet, file_path)
        return {"status": "success", "message": f"Ungrouped rows {start_row}-{end_row}"}
    finally:
        wb.close()


def ungroup_cols(file_path: str, sheet: str, start_col: int, end_col: int) -> dict[str, str]:
    """Ungroup columns by resetting their outline level to 0 and unhiding them."""
    from openpyxl.utils import get_column_letter

    validate_file_path(file_path, must_exist=True)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet)
        for col in range(start_col, end_col + 1):
            letter = get_column_letter(col)
            ws.column_dimensions[letter].outline_level = 0
            ws.column_dimensions[letter].hidden = False
        save_workbook_safe(wb, file_path)
        logger.info("Ungrouped columns %d-%d in '%s' of %s", start_col, end_col, sheet, file_path)
        return {"status": "success", "message": f"Ungrouped columns {start_col}-{end_col}"}
    finally:
        wb.close()


def set_print_titles(
    file_path: str,
    sheet_name: str,
    title_rows: str | None = None,
    title_cols: str | None = None,
) -> dict[str, str | None]:
    """Set repeat rows/cols for print titles (e.g. title_rows='1:2', title_cols='A:B')."""
    import re

    if title_rows is None and title_cols is None:
        raise ValueError("At least one of title_rows or title_cols must be provided.")
    if title_rows is not None and not re.match(r"^\d+:\d+$", title_rows):
        raise ValueError(f"title_rows must be a row-range string like '1:3', got: {title_rows!r}")
    if title_cols is not None and not re.match(r"^[A-Za-z]+:[A-Za-z]+$", title_cols):
        raise ValueError(f"title_cols must be a column-range string like 'A:B', got: {title_cols!r}")
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        if title_rows is not None:
            ws.print_title_rows = title_rows
        if title_cols is not None:
            ws.print_title_cols = title_cols
        save_workbook_safe(wb, file_path)
        logger.info("Set print titles rows=%s cols=%s in '%s' of %s", title_rows, title_cols, sheet_name, file_path)
        return {"status": "ok", "sheet": sheet_name, "title_rows": title_rows, "title_cols": title_cols}
    finally:
        wb.close()


def set_row_height(file_path: str, sheet_name: str, rows: list[int], height: float) -> dict[str, object]:
    """Set the row height for one or more rows."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        for r in rows:
            if r < 1:
                raise ValueError(f"Row number must be >= 1, got {r}")
            ws.row_dimensions[r].height = height
        save_workbook_safe(wb, file_path)
        logger.info("Set height=%s for %d row(s) in '%s' of %s", height, len(rows), sheet_name, file_path)
        return {"status": "ok", "sheet": sheet_name, "updated": len(rows)}
    finally:
        wb.close()


def set_col_width(file_path: str, sheet_name: str, cols: list[str], width: float) -> dict[str, object]:
    """Set the column width for one or more columns (specified as column letters)."""
    wb = load_workbook_safe(file_path)
    try:
        from openpyxl.utils import column_index_from_string

        ws = get_sheet(wb, sheet_name)
        for c in cols:
            column_index_from_string(c)  # raises ValueError on invalid column letter
            ws.column_dimensions[c].width = width
        save_workbook_safe(wb, file_path)
        logger.info("Set width=%s for %d col(s) in '%s' of %s", width, len(cols), sheet_name, file_path)
        return {"status": "ok", "sheet": sheet_name, "updated": len(cols)}
    finally:
        wb.close()


def stack_sheets(
    file_path: str,
    sheet_names: list[str],
    dest_sheet: str,
    include_header: bool = True,
    output_path: str | None = None,
) -> dict[str, object]:
    """Concatenate multiple sheets row-wise into a single destination sheet."""
    import pandas as pd

    validate_file_path(file_path, must_exist=True)
    frames: list[pd.DataFrame] = []
    for name in sheet_names:
        df = pd.read_excel(file_path, sheet_name=name, engine="calamine")
        frames.append(df)
    combined = pd.concat(frames, ignore_index=True)

    target = output_path if output_path is not None else file_path
    if output_path is not None:
        from pathlib import Path as _Path

        validate_file_path(output_path, must_exist=False)
        if _Path(output_path).resolve().exists():
            out_wb = load_workbook_safe(output_path)
        else:
            import openpyxl as _openpyxl

            out_wb = _openpyxl.Workbook()
            out_wb.remove(out_wb.active)
    else:
        out_wb = load_workbook_safe(file_path)

    try:
        if dest_sheet in out_wb.sheetnames:
            del out_wb[dest_sheet]
        dest_ws = out_wb.create_sheet(dest_sheet)
        if include_header:
            dest_ws.append(list(combined.columns))
        for row_tuple in combined.itertuples(index=False, name=None):
            dest_ws.append(list(row_tuple))
        save_workbook_safe(out_wb, target)
        total_rows = len(combined)
        logger.info(
            "Stacked %d sheet(s) into '%s' (%d rows) in %s",
            len(sheet_names),
            dest_sheet,
            total_rows,
            target,
        )
        return {
            "status": "ok",
            "dest_sheet": dest_sheet,
            "total_rows": total_rows,
            "source_sheets": sheet_names,
        }
    finally:
        out_wb.close()


def set_gridlines(file_path: str, sheet_name: str, show: bool = True) -> dict[str, object]:
    """Show or hide gridlines for a worksheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.sheet_view.showGridLines = show
        save_workbook_safe(wb, file_path)
        logger.info("Set showGridLines=%s in '%s' of %s", show, sheet_name, file_path)
        return {"status": "ok", "sheet": sheet_name, "showGridLines": show}
    finally:
        wb.close()


def add_page_break(
    file_path: str,
    sheet_name: str,
    row: int | None = None,
    col: int | None = None,
) -> dict[str, str]:
    """Insert a manual page break before a row or column.

    row: insert a horizontal page break BEFORE this row number (1-based).
    col: insert a vertical page break BEFORE this column index (1-based).
    At least one of row or col must be provided. Both can be provided simultaneously.
    """
    from openpyxl.worksheet.pagebreak import Break

    if row is None and col is None:
        raise ValueError("At least one of 'row' or 'col' must be provided.")

    validate_file_path(file_path, must_exist=True)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        added = []
        if row is not None:
            if row < 2:
                raise ValueError("row must be >= 2 (a break before row 1 has no effect).")
            ws.row_breaks.append(Break(id=row))
            added.append(f"row {row}")
        if col is not None:
            if col < 2:
                raise ValueError("col must be >= 2 (a break before column 1 has no effect).")
            ws.col_breaks.append(Break(id=col))
            added.append(f"col {col}")
        save_workbook_safe(wb, file_path)
        msg = f"Page break(s) added: {', '.join(added)} in sheet '{sheet_name}'."
        logger.info(msg)
        return {"status": "success", "message": msg}
    finally:
        wb.close()


def remove_page_break(
    file_path: str,
    sheet_name: str,
    row: int | None = None,
    col: int | None = None,
) -> dict[str, str]:
    """Remove a manual page break at a specific row or column.

    row: remove the horizontal page break at this row number.
    col: remove the vertical page break at this column index.
    If row/col is None, all breaks of that type are cleared.
    """
    if row is None and col is None:
        raise ValueError("At least one of 'row' or 'col' must be provided.")

    validate_file_path(file_path, must_exist=True)
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        removed = []
        if row is not None:
            before = len(ws.row_breaks.brk)
            ws.row_breaks.brk = [b for b in ws.row_breaks.brk if b.id != row]
            after = len(ws.row_breaks.brk)
            removed.append(f"row {row} ({before - after} removed)")
        if col is not None:
            before = len(ws.col_breaks.brk)
            ws.col_breaks.brk = [b for b in ws.col_breaks.brk if b.id != col]
            after = len(ws.col_breaks.brk)
            removed.append(f"col {col} ({before - after} removed)")
        save_workbook_safe(wb, file_path)
        msg = f"Page break(s) removed: {', '.join(removed)} in sheet '{sheet_name}'."
        logger.info(msg)
        return {"status": "success", "message": msg}
    finally:
        wb.close()
