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
    """An openpyxl ``Border`` as the model's :class:`Borders`.

    ``start`` and ``end`` are read as ``left`` and ``right`` where those are absent (captured).
    A diagonal counts only with a style and a direction, as Excel draws it.
    """
    if part is None:
        return None
    left = side(getattr(part, "left", None), context) or side(getattr(part, "start", None), context)
    right = side(getattr(part, "right", None), context) or side(getattr(part, "end", None), context)
    top = side(getattr(part, "top", None), context)
    bottom = side(getattr(part, "bottom", None), context)
    diagonal = side(getattr(part, "diagonal", None), context)
    up = bool(getattr(part, "diagonalUp", False))
    down = bool(getattr(part, "diagonalDown", False))
    if diagonal is None or not (up or down):
        diagonal, up, down = None, False, False
    if not (left or right or top or bottom or diagonal):
        return None
    return Borders(left, right, top, bottom, diagonal, up, down)


def side(source: Any, context: Context) -> Side | None:
    """One side: ``None`` without a style Excel draws. An automatic color is black (captured)."""
    style = getattr(source, "style", None)
    if style not in _STYLES:
        return None
    return Side(style, context.color(getattr(source, "color", None)))


def unread(part: Any, context: Context) -> Found:
    """Every side is read."""
    return [], []
