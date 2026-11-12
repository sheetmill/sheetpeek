"""Alignment: everything but rotation is read; rotated and stacked text are reported."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from sheetpeek._model import HorizontalAlignment, ReadingOrder, VerticalAlignment

if TYPE_CHECKING:
    from sheetpeek._style import Context, Found

ATTRIBUTE = "alignment"

HORIZONTAL: dict[str | None, HorizontalAlignment] = {
    None: "general",
    "general": "general",
    "left": "left",
    "center": "center",
    "right": "right",
    "fill": "fill",
    "justify": "justify",
    "distributed": "distributed",
    "centerContinuous": "center_across",
}
VERTICAL: dict[str | None, VerticalAlignment] = {
    None: "bottom",
    "bottom": "bottom",
    "top": "top",
    "center": "center",
    "justify": "justify",
    "distributed": "distributed",
}


def read(part: Any, context: Context) -> dict[str, Any]:
    """The alignment fields the model carries."""
    return {
        "h_align": HORIZONTAL.get(part.horizontal, "general"),
        "v_align": VERTICAL.get(part.vertical, "bottom"),
        "wrap": bool(part.wrap_text),
        "indent": int(part.indent or 0),
        "shrink": bool(part.shrinkToFit),
        "reading_order": _reading_order(part.readingOrder),
        "justify_last_line": bool(part.justifyLastLine),
    }


def _reading_order(value: Any) -> ReadingOrder:
    """1 or 2 as stated; anything else, which no Excel interface writes, as 0 (context)."""
    if value == 1:
        return 1
    if value == 2:
        return 2
    return 0


def unread(part: Any, context: Context) -> Found:
    """Rotation and stacking, for cells with a value."""
    found: Found = ([], [])
    rotation = int(part.textRotation or 0)
    if rotation == 255:
        found[1].append(("text.stacked", "", ()))
    elif rotation:
        found[1].append(("text.rotation", "", ()))
    return found
