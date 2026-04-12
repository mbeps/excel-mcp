"""Schemas for hyperlink listing and read operations."""

from __future__ import annotations

from typing import TypedDict

from .common import CellScalar


class HyperlinkInfo(TypedDict):
    """A hyperlink attached to a cell (from list_hyperlinks).

    Keys:
        cell_ref (str): Cell reference. Required.
        target (str | None): External URL target, if present.
        location (str | None): Internal workbook location (sheet/name), if present.
        tooltip (str | None): Tooltip text shown on hover.
    """

    cell_ref: str
    target: str | None
    location: str | None
    tooltip: str | None


class HyperlinkReadResult(TypedDict):
    """Hyperlink details returned by read_hyperlink.

    Keys:
        target, location, tooltip (str | None): As above.
        display_text (CellScalar): Cell text shown for the hyperlink.
    """

    target: str | None
    location: str | None
    tooltip: str | None
    display_text: CellScalar
