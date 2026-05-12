"""Chart configuration model and metadata types used by chart tools."""

from __future__ import annotations

from typing import TypedDict


class ChartInfo(TypedDict):
    """Info about a single chart on a worksheet returned by list operations.

    Keys:
        title (str): Chart title.
        type (str): Chart type string.
        position (str | None): Anchor or position description, if available.
    """

    title: str
    type: str
    position: str | None
