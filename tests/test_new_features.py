from __future__ import annotations

from pathlib import Path

import openpyxl

from mcp_server.tools.analysis import (
    calculate_correlation,
    calculate_percentiles,
    create_histogram,
    extract_unique_values,
    sample_data,
)
from mcp_server.tools.cell_ops import read_cell_detailed
from mcp_server.tools.charts import create_chart, list_charts, update_chart_properties
from mcp_server.tools.conditional_formatting import (
    add_duplicate_rule,
    add_formula_rule,
    add_top_bottom_rule,
    list_conditional_formats,
)
from mcp_server.tools.data_validation import (
    add_date_validation,
    add_dropdown_validation,
    add_text_length_validation,
    list_validations,
    remove_validation,
)
from mcp_server.tools.doc_properties import (
    get_document_properties,
    protect_workbook,
    set_calculation_mode,
    set_document_properties,
)
from mcp_server.tools.formatting import (
    format_cells,
    get_cell_formatting,
    list_merged_ranges,
    merge_cells,
)
from mcp_server.tools.images import insert_image, list_images
from mcp_server.tools.row_col import delete_columns_by_letter, insert_columns_by_letter

# ── Enhanced formatting ──────────────────────────────────────────────


def test_format_cells_extended(sample_xlsx: str) -> None:
    format_cells(
        sample_xlsx,
        "Sheet1",
        "A1:A1",
        font_name="Arial",
        underline="single",
        strikethrough=True,
        text_rotation=45,
        indent=2,
        shrink_to_fit=True,
    )
    wb = openpyxl.load_workbook(sample_xlsx)
    cell = wb["Sheet1"]["A1"]
    assert cell.font.name == "Arial"
    assert cell.font.underline == "single"
    assert cell.font.strike is True
    assert cell.alignment.textRotation == 45
    assert cell.alignment.indent == 2
    assert cell.alignment.shrinkToFit is True
    wb.close()


def test_get_cell_formatting(sample_xlsx: str) -> None:
    format_cells(sample_xlsx, "Sheet1", "A1:A1", bold=True, bg_color="FFFF00")
    result = get_cell_formatting(sample_xlsx, "Sheet1", "A1")
    assert "font" in result
    assert "fill" in result
    assert "border" in result
    assert "alignment" in result
    assert "number_format" in result
    assert result["font"]["bold"] is True


def test_list_merged_ranges(empty_xlsx: str) -> None:
    merge_cells(empty_xlsx, "Sheet1", "A1:C1")
    merge_cells(empty_xlsx, "Sheet1", "D2:E3")
    ranges = list_merged_ranges(empty_xlsx, "Sheet1")
    assert len(ranges) == 2
    assert "A1:C1" in ranges
    assert "D2:E3" in ranges


# ── Enhanced cell_ops ────────────────────────────────────────────────


def test_read_cell_detailed(empty_xlsx: str) -> None:
    wb = openpyxl.load_workbook(empty_xlsx)
    ws = wb["Sheet1"]
    ws["A1"] = 10
    ws["A2"] = 20
    ws["A3"] = "=SUM(A1:A2)"
    wb.save(empty_xlsx)
    wb.close()

    result = read_cell_detailed(empty_xlsx, "Sheet1", "A3")
    assert result["data_type"] == "formula"
    assert result["formula"] == "=SUM(A1:A2)"
    assert result["cell_ref"] == "A3"


# ── Enhanced row_col ─────────────────────────────────────────────────


def test_insert_columns_by_letter(sample_xlsx: str) -> None:
    insert_columns_by_letter(sample_xlsx, "Sheet1", "C", count=1)
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    # Original col C ("City") should now be at D; C should be empty
    assert ws["C1"].value is None
    assert ws["D1"].value == "City"
    wb.close()


def test_delete_columns_by_letter(sample_xlsx: str) -> None:
    delete_columns_by_letter(sample_xlsx, "Sheet1", "B", count=1)
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    # "Age" (col B) deleted; "City" should now be col B
    assert ws["A1"].value == "Name"
    assert ws["B1"].value == "City"
    wb.close()


# ── Enhanced charts ──────────────────────────────────────────────────


def test_create_chart_radar(sample_xlsx: str) -> None:
    create_chart(
        sample_xlsx,
        "Sheet1",
        data_range="A1:D6",
        chart_type="radar",
        target_cell="F1",
        title="Radar",
    )
    charts = list_charts(sample_xlsx, "Sheet1")
    assert len(charts) == 1
    assert charts[0]["type"] == "RadarChart"


def test_update_chart_properties(sample_xlsx: str) -> None:
    create_chart(
        sample_xlsx,
        "Sheet1",
        data_range="A1:D6",
        chart_type="column",
        target_cell="F1",
        title="Original",
    )
    update_chart_properties(
        sample_xlsx,
        "Sheet1",
        chart_index=0,
        title="Updated Title",
    )
    charts = list_charts(sample_xlsx, "Sheet1")
    assert "Updated Title" in str(charts[0]["title"])


# ── Enhanced conditional formatting ──────────────────────────────────


def test_add_formula_rule(sample_xlsx: str) -> None:
    add_formula_rule(
        sample_xlsx,
        "Sheet1",
        cell_range="D2:D6",
        formula="$D2>70000",
    )
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert len(list(ws.conditional_formatting)) > 0
    wb.close()


def test_add_top_bottom_rule(sample_xlsx: str) -> None:
    add_top_bottom_rule(
        sample_xlsx,
        "Sheet1",
        cell_range="D2:D6",
        rank=10,
    )
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert len(list(ws.conditional_formatting)) > 0
    wb.close()


def test_add_duplicate_rule(sample_xlsx: str) -> None:
    add_duplicate_rule(
        sample_xlsx,
        "Sheet1",
        cell_range="C2:C6",
    )
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Sheet1"]
    assert len(list(ws.conditional_formatting)) > 0
    wb.close()


def test_list_conditional_formats(sample_xlsx: str) -> None:
    add_formula_rule(sample_xlsx, "Sheet1", "D2:D6", formula="$D2>70000")
    add_duplicate_rule(sample_xlsx, "Sheet1", "C2:C6")
    result = list_conditional_formats(sample_xlsx, "Sheet1")
    assert isinstance(result, list)
    assert len(result) >= 2


# ── Enhanced data validation ─────────────────────────────────────────


def test_add_date_validation(empty_xlsx: str) -> None:
    add_date_validation(
        empty_xlsx,
        "Sheet1",
        cell_range="A1:A10",
        operator="greaterThan",
        date1="2024-01-01",
    )
    validations = list_validations(empty_xlsx, "Sheet1")
    assert len(validations) == 1
    assert validations[0]["type"] == "date"


def test_add_text_length_validation(empty_xlsx: str) -> None:
    add_text_length_validation(
        empty_xlsx,
        "Sheet1",
        cell_range="B1:B10",
        operator="lessThanOrEqual",
        length1=50,
    )
    validations = list_validations(empty_xlsx, "Sheet1")
    assert len(validations) == 1
    assert validations[0]["type"] == "textLength"


def test_remove_validation(empty_xlsx: str) -> None:
    add_dropdown_validation(
        empty_xlsx,
        "Sheet1",
        cell_range="A1:A5",
        options=["X", "Y", "Z"],
    )
    assert len(list_validations(empty_xlsx, "Sheet1")) == 1
    remove_validation(empty_xlsx, "Sheet1", cell_range="A1:A5")
    assert len(list_validations(empty_xlsx, "Sheet1")) == 0


# ── Enhanced analysis ────────────────────────────────────────────────


def test_calculate_correlation(sample_xlsx: str) -> None:
    result = calculate_correlation(sample_xlsx, "Sheet1", columns=["Age", "Salary"])
    assert "matrix" in result
    assert len(result["matrix"]) == 2
    assert len(result["matrix"][0]) == 2


def test_calculate_percentiles(sample_xlsx: str) -> None:
    result = calculate_percentiles(sample_xlsx, "Sheet1", column="Salary")
    assert "percentile_values" in result
    assert len(result["percentile_values"]) > 0


def test_sample_data(sample_xlsx: str) -> None:
    result = sample_data(sample_xlsx, "Sheet1", n=3, random_state=42)
    assert result["sample_size"] == 3
    assert result["total_rows"] == 5


def test_create_histogram(sample_xlsx: str) -> None:
    result = create_histogram(sample_xlsx, "Sheet1", column="Salary", bins=3)
    assert "bin_edges" in result
    assert "frequencies" in result
    assert len(result["frequencies"]) == 3


def test_extract_unique_values(sample_xlsx: str) -> None:
    result = extract_unique_values(sample_xlsx, "Sheet1", column="City")
    assert result["count"] == 3
    assert set(result["unique_values"]) == {"New York", "Chicago", "Boston"}


# ── Images ───────────────────────────────────────────────────────────


def _create_test_png(path: Path) -> str:
    from PIL import Image as PILImage

    img = PILImage.new("RGB", (10, 10), color="red")
    png_path = str(path / "test.png")
    img.save(png_path)
    return png_path


def test_insert_image(empty_xlsx: str, tmp_path: Path) -> None:
    png_path = _create_test_png(tmp_path)
    result = insert_image(empty_xlsx, "Sheet1", png_path, "A1")
    assert "inserted" in result.lower()
    images = list_images(empty_xlsx, "Sheet1")
    assert len(images) == 1


def test_list_images(empty_xlsx: str, tmp_path: Path) -> None:
    png_path = _create_test_png(tmp_path)
    insert_image(empty_xlsx, "Sheet1", png_path, "B3")
    images = list_images(empty_xlsx, "Sheet1")
    assert len(images) == 1
    assert "width" in images[0]
    assert "height" in images[0]


# ── Document properties ─────────────────────────────────────────────


def test_get_document_properties(sample_xlsx: str) -> None:
    props = get_document_properties(sample_xlsx)
    assert "title" in props
    assert "creator" in props


def test_set_document_properties(empty_xlsx: str) -> None:
    set_document_properties(empty_xlsx, title="Test Workbook", creator="TestBot")
    props = get_document_properties(empty_xlsx)
    assert props["title"] == "Test Workbook"
    assert props["creator"] == "TestBot"


def test_protect_workbook(empty_xlsx: str) -> None:
    result = protect_workbook(empty_xlsx, lock_structure=True, lock_windows=True)
    assert "protected" in result.lower()
    wb = openpyxl.load_workbook(empty_xlsx)
    assert wb.security.lockStructure is True
    assert wb.security.lockWindows is True
    wb.close()


# ── Calculation mode ─────────────────────────────────────────────────


def test_set_calculation_mode(empty_xlsx: str) -> None:
    set_calculation_mode(empty_xlsx, mode="manual")
    wb = openpyxl.load_workbook(empty_xlsx)
    assert wb.calculation is not None
    assert wb.calculation.calcMode == "manual"
    wb.close()
