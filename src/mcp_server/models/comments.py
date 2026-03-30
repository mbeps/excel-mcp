from __future__ import annotations

from typing import TypedDict


class CommentInfo(TypedDict):
    """A comment attached to a cell."""

    cell_ref: str
    text: str
    author: str
