"""Fills: pattern fills (``solid`` and the 17 patterns) and gradient fills, all read.

``CellStyle.fill`` is one representative color, for outputs that show one; ``fill_spec`` is the
fill itself (``ARCHITECTURE.md`` §4.12).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from sheetpeek._fill_border_model import (
    BLACK,
    TILES,
    WHITE,
    Fill,
    Gradient,
    Pattern,
    finite,
    unit,
)

if TYPE_CHECKING:
    from sheetpeek._style import Context, Found

ATTRIBUTE = "fill"
_SYSTEM_FOREGROUND, _SYSTEM_BACKGROUND = 64, 65
_MISSING = "00000000"  # what openpyxl reads for a missing color element


def read(part: Any, context: Context) -> dict[str, Any]:
    """The fill, and its representative color."""
    spec = to_fill(part, context)
    return {"fill": spec.color if spec is not None else None, "fill_spec": spec}


def to_fill(part: Any, context: Context) -> Fill | None:
    """An openpyxl fill as the model's, or ``None`` where Excel draws no fill."""
    tag = getattr(part, "tagname", None)
    if tag == "patternFill":
        kind = getattr(part, "fill_type", None)
        if kind != "solid" and kind not in TILES:  # none, or no patternType: nothing drawn
            return None
        fg = color(getattr(part, "fgColor", None), context)
        bg = color(getattr(part, "bgColor", None), context, background=True)
        if kind == "solid":
            # A solid fill in the system foreground color is no fill (captured for indexed 64;
            # automatic taken alike, unverified).
            return Pattern("solid", fg) if fg is not None else None
        return Pattern(kind, fg, bg)
    if tag == "gradientFill":
        stops = sorted(
            (unit(getattr(stop, "position", 0)), color(getattr(stop, "color", None), context))
            for stop in getattr(part, "stop", ()) or ()
        )
        if not stops:  # nothing to draw; what Excel does is not captured
            return None
        return Gradient(
            "path" if getattr(part, "type", None) == "path" else "linear",
            tuple((position, found or BLACK) for position, found in stops),
            degree=finite(getattr(part, "degree", 0)) % 360,
            left=unit(getattr(part, "left", 0)),
            right=unit(getattr(part, "right", 0)),
            top=unit(getattr(part, "top", 0)),
            bottom=unit(getattr(part, "bottom", 0)),
        )
    return None


def color(source: Any, context: Context, *, background: bool = False) -> str | None:
    """A fill color as ``RRGGBB``, or ``None`` for automatic (captures on Excel 16.114).

    Indexed 64 is automatic, 65 white, and an index past the palette black. openpyxl reads a
    missing color element as ``00000000``: for a foreground that is black either way, and a
    missing background is automatic (white).
    """
    if source is None:
        return None
    kind = getattr(source, "type", None)
    if kind == "indexed":
        index = getattr(source, "indexed", None)
        if index == _SYSTEM_FOREGROUND:
            return None
        if index == _SYSTEM_BACKGROUND:
            return WHITE
        return context.color(source) or BLACK
    if kind == "rgb" and background and getattr(source, "rgb", None) == _MISSING:
        return None
    return context.color(source)


def unread(part: Any, context: Context) -> Found:
    """Every fill is read."""
    return [], []
