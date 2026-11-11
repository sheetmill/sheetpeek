"""Fills: a solid color is read; pattern and gradient fills are reported."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from sheetpeek._style import Context, Found

ATTRIBUTE = "fill"


def read(part: Any, context: Context) -> dict[str, Any]:
    """A pattern fill's foreground color, which is the whole of a solid fill."""
    color = None
    if getattr(part, "tagname", None) == "patternFill" and part.fill_type not in (None, "none"):
        color = context.color(part.fgColor)
    return {"fill": color}


def unread(part: Any, context: Context) -> Found:
    """Gradients, and patterns other than solid (drawn as their pattern color)."""
    found: Found = ([], [])
    tag = getattr(part, "tagname", None)
    if tag == "gradientFill":
        found[0].append(("fill.gradient", "", ()))
    elif tag == "patternFill" and (pattern := getattr(part, "fill_type", None)) not in (
        None,
        "none",
        "solid",
    ):
        found[0].append(("fill.pattern", "", (("pattern", str(pattern)),)))
    return found
