"""Alignment: horizontal and vertical, wrap, indent, shrink, reading order and rotation."""

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
        "rotation": _rotation(part.textRotation),
        "stacked": part.textRotation == 255,
    }


def _rotation(value: Any) -> int:
    """Degrees counterclockwise: 0 to 90 as stored, 91 to 180 for -1 to -90 (§4.17).

    255 (stacked) is 0. openpyxl refuses a workbook with any other value (measured), so the
    last branch only guards the move to the shared xlsx reader.
    """
    angle = int(value or 0)
    if 0 <= angle <= 90:
        return angle
    if 90 < angle <= 180:
        return 90 - angle
    return 0


def _reading_order(value: Any) -> ReadingOrder:
    """1 or 2 as stated; anything else, which no Excel interface writes, as 0 (context)."""
    if value == 1:
        return 1
    if value == 2:
        return 2
    return 0


def unread(part: Any, context: Context) -> Found:
    """Nothing: every alignment field is read."""
    return [], []
