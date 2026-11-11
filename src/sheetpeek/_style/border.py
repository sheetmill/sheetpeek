"""Borders: not drawn yet, so only reported."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from sheetpeek._style import Context, Found

ATTRIBUTE = "border"
_SIDES = ("left", "right", "top", "bottom", "start", "end")


def read(part: Any, context: Context) -> dict[str, Any]:
    """The model carries no borders yet."""
    return {}


def unread(part: Any, context: Context) -> Found:
    """A side with a style; a diagonal only with its flags (what Excel draws)."""
    every: Found = ([], [])
    if any(getattr(getattr(part, side, None), "style", None) for side in _SIDES):
        every[0].append(("border", "", ()))
    diagonal = getattr(part, "diagonal", None)
    if getattr(diagonal, "style", None) and (part.diagonalUp or part.diagonalDown):
        every[0].append(("border.diagonal", "", ()))
    return every
