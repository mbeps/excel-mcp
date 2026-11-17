"""Phase 5 remediation & 100% coverage tests for Formatting, Governance, Metadata, Prompts, and Models."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl
import pytest
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image as PILImage

from mcp_server.routes import formatting as r_formatting
from mcp_server.routes import governance as r_governance
from mcp_server.routes import metadata as r_metadata
from mcp_server.tools import comments, conditional_formatting, data_validation, formatting, hyperlinks, images, named_ranges, protection, scenarios


class _FakeMCP:
    def __init__(self) -> None:
        self.tools: list[str] = []

    def tool(self, *args: Any, **kwargs: Any) -> Any:
        def _dec(fn: Any) -> Any:
            self.tools.append(fn.__name__)
            return fn

        return _dec


def _make_wb(tmp_path: Path, name: str = "book.xlsx") -> str:
    fp = str(tmp_path / name)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Value"])
    ws.append(["Alpha", 10])
    ws.append(["Beta", 20])
    wb.create_sheet("Sheet2")
    wb.save(fp)
    wb.close()
    return fp


def _make_png(tmp_path: Path, name: str = "logo.png", size: tuple[int, int] = (200, 100)) -> str:
    fp = str(tmp_path / name)
    img = PILImage.new("RGB", size, color=(255, 0, 0))
    img.save(fp)
    return fp


# ── Formatting & Conditional Formatting ──────────────────────────────────────


def test_formatting_hex_normalization_preserves_unspecified_styles_and_empty_numfmt(tmp_path: Path) -> None:
    fp = _make_wb(tmp_path)
    formatting.format_cells(
        fp,
        "Sheet1",
        "A1",
        bold=True,
        font_color="#f00",
        bg_color="#00ff00",
        horizontal_alignment="center",
        border_style="thin",
        border_color="#112233",
        number_format="0.00",
    )
    # Calling format_cells with only number_format="" should NOT clobber font/alignment/border/fill
    formatting.format_cells(fp, "Sheet1", "A1", number_format="")

    wb = openpyxl.load_workbook(fp)
    cell = wb["Sheet1"]["A1"]
    assert cell.font.bold is True
    assert "FF0000" in cell.font.color.rgb
    assert "00FF00" in cell.fill.start_color.rgb
    assert cell.alignment.horizontal == "center"
    assert cell.border.left.style == "thin"
    assert cell.number_format == ""
    wb.close()

    # preserve_existing=True with alignment, border, number_format
    formatting.format_cells(
        fp,
        "Sheet1",
        "A1",
        vertical_alignment="top",
        top_border_style="medium",
        number_format="0.0",
        preserve_existing=True,
    )
    wb = openpyxl.load_workbook(fp)
    cell = wb["Sheet1"]["A1"]
    assert cell.alignment.horizontal == "center"
    assert cell.alignment.vertical == "top"
    assert cell.border.top.style == "medium"
    assert cell.number_format == "0.0"
    wb.close()

    with pytest.raises(ValueError, match="Invalid hex color"):
        formatting.format_cells(fp, "Sheet1", "A1", font_color="not-a-color")


def test_auto_fit_columns_merged_multiline_and_clamping(tmp_path: Path) -> None:
    fp = _make_wb(tmp_path)
    wb = openpyxl.load_workbook(fp)
    ws = wb["Sheet1"]
    ws.merge_cells("A4:C4")
    ws["A4"] = "A very long merged banner title across three columns"
    ws["A5"] = "Line1\nLine2"
    ws["B5"] = "X" * 300
    wb.save(fp)
    wb.close()

    formatting.auto_fit_columns(fp, "Sheet1")
    wb = openpyxl.load_workbook(fp)
    ws = wb["Sheet1"]
    assert ws.column_dimensions["A"].width >= 8
    assert ws.column_dimensions["B"].width == 255
    assert ws.column_dimensions["C"].width == 8
    wb.close()


def test_formatting_invalid_ranges_and_routes(tmp_path: Path) -> None:
    fp = _make_wb(tmp_path)
    with pytest.raises(ValueError, match="Invalid source_cell"):
        formatting.copy_cell_format(fp, "Sheet1", "BAD!", "A1")
    with pytest.raises(ValueError, match="Invalid target_range"):
        formatting.copy_cell_format(fp, "Sheet1", "A1", "BAD!")
    with pytest.raises(ValueError, match="Invalid range"):
        formatting.clear_cell_format(fp, "Sheet1", "BAD!")
    with pytest.raises(ValueError, match="Invalid range"):
        formatting.apply_named_style(fp, "Sheet1", "BAD!", "Good")

    r_formatting.format_cells(fp, "Sheet1", "A1", bold=True)
    r_formatting.auto_fit_columns(fp, "Sheet1")
    r_formatting.copy_cell_format(fp, "Sheet1", "A1", "B1")
    r_formatting.apply_named_style(fp, "Sheet1", "A2", "Good")
    r_formatting.clear_cell_format(fp, "Sheet1", "A2")

    mcp = _FakeMCP()
    r_formatting.register(mcp)
    assert len(mcp.tools) == 5


def test_conditional_formatting_icons_fills_list_and_remove(tmp_path: Path) -> None:
    fp = _make_wb(tmp_path)
    conditional_formatting.apply_conditional_formatting(fp, "Sheet1", "B2:B3", "icon_set", icon_style="4Arrows")
    conditional_formatting.apply_conditional_formatting(fp, "Sheet1", "B2:B3", "icon_set", icon_style="5Arrows")
    conditional_formatting.add_highlight_rule(
        fp, "Sheet1", "B2:B3", "greaterThan", "5", font_color="#00ff00", bg_color="#112233"
    )
    conditional_formatting.add_formula_rule(fp, "Sheet1", "B2:B3", "=B2>0", stop_if_true=True)

    rules = conditional_formatting.list_conditional_formatting(fp, "Sheet1")
    assert len(rules) == 4
    assert all(r["range"] == "B2:B3" for r in rules)

    # Remove all matching normalized sqrefs (e.g. lowercase + $)
    conditional_formatting.remove_conditional_formatting(fp, "Sheet1", "$b$2:$b$3")
    assert conditional_formatting.list_conditional_formatting(fp, "Sheet1") == []


# ── Governance: Protection, Data Validation, Routes ──────────────────────────


def test_protection_single_cell_hidden_and_unprotect_clears_hashes(tmp_path: Path) -> None:
    fp = _make_wb(tmp_path)
    wb = openpyxl.load_workbook(fp)
    from openpyxl.styles import Protection as CellProtection

    wb["Sheet1"]["A1"].protection = CellProtection(locked=False, hidden=True)
    wb["Sheet1"]["B1"].protection = CellProtection(locked=True, hidden=True)
    wb.save(fp)
    wb.close()

    protection.protect_cells(fp, "Sheet1", locked_range="A1", unlocked_ranges=["B1"])
    wb = openpyxl.load_workbook(fp)
    assert wb["Sheet1"]["A1"].protection.locked is True
    assert wb["Sheet1"]["A1"].protection.hidden is True
    assert wb["Sheet1"]["B1"].protection.locked is False
    assert wb["Sheet1"]["B1"].protection.hidden is True
    wb.close()

    protection.protect_sheet(fp, "Sheet1", password="pw")
    protection.unprotect_sheet(fp, "Sheet1")
    wb = openpyxl.load_workbook(fp)
    assert wb["Sheet1"].protection.sheet is False
    assert wb["Sheet1"].protection.hashValue is None
    assert wb["Sheet1"].protection.saltValue is None
    wb.close()


def test_data_validation_edge_cases(tmp_path: Path) -> None:
    fp = _make_wb(tmp_path)
    with pytest.raises(ValueError, match="non-empty"):
        data_validation.add_dropdown_validation(fp, "Sheet1", "A2", options=[])
    with pytest.raises(ValueError, match="255-character"):
        data_validation.add_dropdown_validation(fp, "Sheet1", "A2", options=["X" * 260])

    data_validation.add_numeric_validation(fp, "Sheet1", "B2:B3", "between", 0.0, 0.5)
    data_validation.add_date_validation(fp, "Sheet1", "C2", "greaterThan", "DATE(2025,1,1)")
    with pytest.raises(ValueError, match="empty"):
        data_validation.add_date_validation(fp, "Sheet1", "C2", "greaterThan", "")
    with pytest.raises(ValueError, match="YYYY-MM-DD"):
        data_validation.add_date_validation(fp, "Sheet1", "C2", "greaterThan", "not-a-date")

    wb = openpyxl.load_workbook(fp)
    num_dv = next(dv for dv in wb["Sheet1"].data_validations.dataValidation if dv.type == "decimal")
    assert num_dv.formula1 == "0.0"
    assert num_dv.formula2 == "0.5"
    wb.close()

    res = data_validation.remove_validation(fp, "Sheet1", "b2:b3")
    assert "Removed 1" in res


def test_governance_route_all_branches_and_errors(tmp_path: Path) -> None:
    fp = _make_wb(tmp_path)
    with pytest.raises(ValueError, match="sheet_name is required"):
        r_governance.protection("protect_sheet", fp)
    with pytest.raises(ValueError, match="sheet_name is required"):
        r_governance.protection("unprotect_sheet", fp)
    with pytest.raises(ValueError, match="locked_range is required"):
        r_governance.protection("protect_cells", fp, sheet_name="Sheet1")
    with pytest.raises(ValueError, match="sheet_name is required"):
        r_governance.protection("protect_cells", fp, locked_range="A1")
    with pytest.raises(ValueError, match="Unknown action"):
        r_governance.protection("bad", fp)  # type: ignore[arg-type]

    with pytest.raises(ValueError, match="options or source_range"):
        r_governance.data_validation("dropdown", fp, "Sheet1", "A1")
    with pytest.raises(ValueError, match="operator is required"):
        r_governance.data_validation("numeric", fp, "Sheet1", "A1")
    with pytest.raises(ValueError, match="value1 is required"):
        r_governance.data_validation("numeric", fp, "Sheet1", "A1", operator="greaterThan")
    with pytest.raises(ValueError, match="formula is required"):
        r_governance.data_validation("formula", fp, "Sheet1", "A1")
    with pytest.raises(ValueError, match="Unknown action"):
        r_governance.data_validation("bad", fp, "Sheet1", "A1")  # type: ignore[arg-type]

    r_governance.data_validation("dropdown", fp, "Sheet1", "A2", options=["Y", "N"])
    r_governance.data_validation("numeric", fp, "Sheet1", "B2", operator="greaterThan", value1=0)
    r_governance.data_validation("date", fp, "Sheet1", "C2", date1="2025-01-01")
    r_governance.data_validation("formula", fp, "Sheet1", "D2", formula="=D2>0")
    r_governance.data_validation("remove", fp, "Sheet1", "D2")

    with pytest.raises(ValueError, match="Unknown action"):
        r_governance.doc_properties("bad", fp)  # type: ignore[arg-type]

    # conditional_format route: list allows sheet_name=None, others require sheet_name
    assert r_governance.conditional_format("list", fp) == []
    with pytest.raises(ValueError, match="sheet_name is required"):
        r_governance.conditional_format("apply", fp)
    for act in ("apply", "highlight", "formula_rule", "top_bottom", "above_below_average"):
        with pytest.raises(ValueError, match="cell_range is required"):
            r_governance.conditional_format(act, fp, "Sheet1")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="format_type is required"):
        r_governance.conditional_format("apply", fp, "Sheet1", "A1")
    with pytest.raises(ValueError, match="operator is required"):
        r_governance.conditional_format("highlight", fp, "Sheet1", "A1")
    with pytest.raises(ValueError, match="formula is required"):
        r_governance.conditional_format("highlight", fp, "Sheet1", "A1", operator="greaterThan")
    with pytest.raises(ValueError, match="formula is required"):
        r_governance.conditional_format("formula_rule", fp, "Sheet1", "A1")
    with pytest.raises(ValueError, match="Unknown action"):
        r_governance.conditional_format("bad", fp, "Sheet1")  # type: ignore[arg-type]

    r_governance.conditional_format("apply", fp, "Sheet1", "B2:B3", format_type="color_scale")
    r_governance.conditional_format("highlight", fp, "Sheet1", "B2:B3", operator="greaterThan", formula="5")
    r_governance.conditional_format("formula_rule", fp, "Sheet1", "B2:B3", formula="=B2>5", stop_if_true=True)
    r_governance.conditional_format("top_bottom", fp, "Sheet1", "B2:B3")
    r_governance.conditional_format("above_below_average", fp, "Sheet1", "B2:B3")
    r_governance.conditional_format("remove", fp, "Sheet1", "B2:B3")

    mcp = _FakeMCP()
    r_governance.register(mcp)
    assert len(mcp.tools) == 4


# ── Metadata: Named Ranges, Comments, Hyperlinks, Scenarios, Images, Route ───


def test_named_ranges_sheet_scoped_and_auto_quote(tmp_path: Path) -> None:
    fp = _make_wb(tmp_path)
    wb = openpyxl.load_workbook(fp)
    wb.create_sheet("My Sheet")
    wb.save(fp)
    wb.close()

    res = named_ranges.create_named_range(fp, "LocalVal", "My Sheet!$A$1:$A$2", scope="My Sheet")
    assert "'My Sheet'!$A$1:$A$2" in res

    listed = named_ranges.list_named_ranges(fp)
    assert any(nr["name"] == "LocalVal" and nr["scope"] == "My Sheet" for nr in listed)

    named_ranges.update_named_range(fp, "LocalVal", "Sheet1!$B$1:$B$2")
    listed2 = named_ranges.list_named_ranges(fp)
    assert any(nr["name"] == "LocalVal" and nr["destination"] == "Sheet1!$B$1:$B$2" for nr in listed2)

    named_ranges.delete_named_range(fp, "LocalVal")
    assert not any(nr["name"] == "LocalVal" for nr in named_ranges.list_named_ranges(fp))

    with pytest.raises(ValueError, match="not found"):
        named_ranges.delete_named_range(fp, "MissingRange")
    with pytest.raises(ValueError, match="not found"):
        named_ranges.update_named_range(fp, "MissingRange", "Sheet1!$A$1")

    # Already-quoted destination stays as-is
    res_quoted = named_ranges.create_named_range(fp, "QuotedVal", "'My Sheet'!$B$1")
    assert "'My Sheet'!$B$1" in res_quoted
    # Plain destination without spaces stays unquoted
    res_plain = named_ranges.create_named_range(fp, "PlainVal", "Sheet1!$A$1")
    assert "Sheet1!$A$1" in res_plain
    # Bare ref without '!' stays as-is
    res_bare = named_ranges.create_named_range(fp, "BareVal", "$A$1")
    assert "$A$1" in res_bare


def test_single_cell_validation_hyperlinks_scenarios_images(tmp_path: Path) -> None:
    fp = _make_wb(tmp_path)
    for bad in ("A1:B2", "NOTACELL"):
        with pytest.raises(ValueError):
            comments.add_comment(fp, "Sheet1", bad, "hi")
        with pytest.raises(ValueError):
            hyperlinks.add_hyperlink(fp, "Sheet1", bad, "https://example.com")
        with pytest.raises(ValueError):
            images.insert_image(fp, "Sheet1", _make_png(tmp_path), bad)

    # Internal sheet hyperlink
    hyperlinks.add_hyperlink(fp, "Sheet1", "A2", "#Sheet2!B3", display_text="Go", tooltip="Jump")
    hl = hyperlinks.read_hyperlink(fp, "Sheet1", "A2")
    assert hl is not None
    assert hl["location"] == "Sheet2!B3"
    assert hl["target"] is None

    # Scenarios: block _mcp_scenarios self-targeting
    with pytest.raises(ValueError, match="_mcp_scenarios"):
        scenarios.add_scenario(fp, "Bad", {"_mcp_scenarios": {"A1": 1}})

    # Images: single-axis aspect ratio, invalid dims, corrupt image
    png = _make_png(tmp_path, size=(200, 100))
    images.insert_image(fp, "Sheet1", png, "D2", width=100)
    images.insert_image(fp, "Sheet1", png, "D5", height=50)
    images.insert_image(fp, "Sheet1", png, "D8", width=40, height=40)

    with pytest.raises(ValueError, match="positive"):
        images.insert_image(fp, "Sheet1", png, "D2", width=0)
    with pytest.raises(ValueError, match="positive"):
        images.insert_image(fp, "Sheet1", png, "D2", height=-5)

    bad_png = tmp_path / "bad.png"
    bad_png.write_text("not a real png")
    with pytest.raises(ValueError, match="Failed to load image"):
        images.insert_image(fp, "Sheet1", str(bad_png), "D2")


def test_metadata_route_all_branches_and_errors(tmp_path: Path) -> None:
    fp = _make_wb(tmp_path)

    for act in ("add", "read", "delete"):
        with pytest.raises(ValueError, match="cell_ref is required"):
            r_metadata.comment(act, fp, "Sheet1")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="text is required"):
        r_metadata.comment("add", fp, "Sheet1", cell_ref="A1")
    with pytest.raises(ValueError, match="Unknown action"):
        r_metadata.comment("bad", fp, "Sheet1")  # type: ignore[arg-type]
    r_metadata.comment("add", fp, "Sheet1", cell_ref="A1", text="note")
    assert r_metadata.comment("read", fp, "Sheet1", cell_ref="A1") is not None
    assert len(r_metadata.comment("list", fp, "Sheet1")) == 1  # type: ignore[arg-type]
    r_metadata.comment("delete", fp, "Sheet1", cell_ref="A1")

    for act in ("add", "read", "delete"):
        with pytest.raises(ValueError, match="cell_ref is required"):
            r_metadata.hyperlink(act, fp, "Sheet1")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="url is required"):
        r_metadata.hyperlink("add", fp, "Sheet1", cell_ref="A1")
    with pytest.raises(ValueError, match="Unknown action"):
        r_metadata.hyperlink("bad", fp, "Sheet1")  # type: ignore[arg-type]
    r_metadata.hyperlink("add", fp, "Sheet1", cell_ref="A1", url="https://example.com")
    assert r_metadata.hyperlink("read", fp, "Sheet1", cell_ref="A1") is not None
    assert len(r_metadata.hyperlink("list", fp, "Sheet1")) == 1  # type: ignore[arg-type]
    r_metadata.hyperlink("delete", fp, "Sheet1", cell_ref="A1")

    with pytest.raises(ValueError, match="name is required"):
        r_metadata.scenario("add", fp)
    with pytest.raises(ValueError, match="cell_values is required"):
        r_metadata.scenario("add", fp, name="S1")
    with pytest.raises(ValueError, match="name is required"):
        r_metadata.scenario("apply", fp)
    with pytest.raises(ValueError, match="Unknown action"):
        r_metadata.scenario("bad", fp)  # type: ignore[arg-type]
    r_metadata.scenario("add", fp, name="S1", cell_values={"Sheet1": {"B2": 99}})
    assert len(r_metadata.scenario("list", fp)) == 1  # type: ignore[arg-type]
    r_metadata.scenario("apply", fp, name="S1")

    with pytest.raises(ValueError, match="name is required"):
        r_metadata.named_range("create", fp)
    with pytest.raises(ValueError, match="destination is required"):
        r_metadata.named_range("create", fp, name="N1")
    with pytest.raises(ValueError, match="name is required"):
        r_metadata.named_range("delete", fp)
    with pytest.raises(ValueError, match="name is required"):
        r_metadata.named_range("update", fp)
    with pytest.raises(ValueError, match="new_destination is required"):
        r_metadata.named_range("update", fp, name="N1")
    with pytest.raises(ValueError, match="Unknown action"):
        r_metadata.named_range("bad", fp)  # type: ignore[arg-type]
    r_metadata.named_range("create", fp, name="N1", destination="Sheet1!$A$1:$B$2")
    r_metadata.named_range("update", fp, name="N1", new_destination="Sheet1!$A$1:$B$3")
    r_metadata.named_range("delete", fp, name="N1")

    with pytest.raises(ValueError, match="data_range is required"):
        r_metadata.table("create", fp, "Sheet1")
    with pytest.raises(ValueError, match="table_name is required"):
        r_metadata.table("create", fp, "Sheet1", data_range="A1:B3")
    with pytest.raises(ValueError, match="table_name is required"):
        r_metadata.table("resize", fp, "Sheet1")
    with pytest.raises(ValueError, match="new_range is required"):
        r_metadata.table("resize", fp, "Sheet1", table_name="T1")
    with pytest.raises(ValueError, match="table_name is required"):
        r_metadata.table("totals", fp, "Sheet1")
    with pytest.raises(ValueError, match="show_totals is required"):
        r_metadata.table("totals", fp, "Sheet1", table_name="T1")
    with pytest.raises(ValueError, match="table_name is required"):
        r_metadata.table("data", fp, "Sheet1")
    with pytest.raises(ValueError, match="table_name is required"):
        r_metadata.table("convert_to_range", fp, "Sheet1")
    with pytest.raises(ValueError, match="Unknown action"):
        r_metadata.table("bad", fp, "Sheet1")  # type: ignore[arg-type]
    r_metadata.table("create", fp, "Sheet1", table_name="T1", data_range="A1:B3")
    r_metadata.table("resize", fp, "Sheet1", table_name="T1", new_range="A1:B2")
    r_metadata.table("totals", fp, "Sheet1", table_name="T1", show_totals=True, column_totals={"Value": "sum"})
    r_metadata.table("data", fp, "Sheet1", table_name="T1")
    r_metadata.table("convert_to_range", fp, "Sheet1", table_name="T1")

    mcp = _FakeMCP()
    r_metadata.register(mcp)
    assert len(mcp.tools) == 5
