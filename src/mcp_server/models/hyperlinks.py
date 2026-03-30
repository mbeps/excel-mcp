from __future__ import annotations

from typing import TypedDict

from .common import CellScalar


class HyperlinkInfo(TypedDict):
    """A hyperlink attached to a cell (from list_hyperlinks)."""

    cell_ref: str
    target: str | None
    location: str | None
    tooltip: str | None


class HyperlinkReadResult(TypedDict):
    """Hyperlink details returned by read_hyperlink."""

    target: str | None
    location: str | None
    tooltip: str | None
    display_text: CellScalar
