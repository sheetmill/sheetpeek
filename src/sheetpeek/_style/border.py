"""Borders: the four sides in their 13 styles and colors, and the diagonals, all read."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from sheetpeek._fill_border_model import STYLES, Borders, Side

if TYPE_CHECKING:
    from sheetpeek._style import Context, Found

ATTRIBUTE = "border"
_STYLES = frozenset(STYLES)


def read(part: Any, context: Context) -> dict[str, Any]:
    """The cell's sides, or ``None`` when it has none Excel draws."""
    return {"borders": to_borders(part, context)}


def to_borders(part: Any, context: Context) -> Borders | None:
    """An ``xlio.Border`` as the model's :class:`Borders`.

    xlio reads ``start`` and ``end`` as ``left`` and ``right`` where those are absent (unverified).
    A diagonal counts only with a style and a direction, as Excel draws it.
    """
    if part is None:
        return None
    left = side(part.left, context)
    right = side(part.right, context)
    top = side(part.top, context)
    bottom = side(part.bottom, context)
    diagonal = side(part.diagonal, context)
    up = bool(part.diagonal_up)
    down = bool(part.diagonal_down)
    if diagonal is None or not (up or down):
        diagonal, up, down = None, False, False
    if not (left or right or top or bottom or diagonal):
        return None
    return Borders(left, right, top, bottom, diagonal, up, down)


def side(source: Any, context: Context) -> Side | None:
    """One side: ``None`` without a style Excel draws.

    An automatic color is black (captured, ``auto_colors``).
    """
    style = source.style if source is not None else None
    if style not in _STYLES:
        return None
    return Side(style, context.color(source.color))


def unread(part: Any, context: Context) -> Found:
    """Every side is read."""
    return [], []
