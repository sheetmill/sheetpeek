"""Alignment: horizontal and vertical alignment, wrap and indent are read; the rest reported."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from sheetpeek._model import HorizontalAlignment, VerticalAlignment

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
    """The four alignment fields the model carries."""
    return {
        "h_align": HORIZONTAL.get(part.horizontal, "general"),
        "v_align": VERTICAL.get(part.vertical, "bottom"),
        "wrap": bool(part.wrap_text),
        "indent": int(part.indent or 0),
    }


def unread(part: Any, context: Context) -> Found:
    """Rotation, stacking, shrink to fit and reading order, for cells with a value."""
    found: Found = ([], [])
    rotation = int(part.textRotation or 0)
    if rotation == 255:
        found[1].append(("text.stacked", "", ()))
    elif rotation:
        found[1].append(("text.rotation", "", ()))
    if part.shrinkToFit:
        found[1].append(("text.shrink", "", ()))
    if part.readingOrder in (1, 2):
        found[1].append(("text.reading_order", "", ()))
    return found
