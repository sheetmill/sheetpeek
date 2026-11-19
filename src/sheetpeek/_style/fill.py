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
_ZERO = "00000000"


def read(part: Any, context: Context) -> dict[str, Any]:
    """The fill, and its representative color."""
    spec = to_fill(part, context)
    return {"fill": spec.color if spec is not None else None, "fill_spec": spec}


def to_fill(part: Any, context: Context) -> Fill | None:
    """An ``xlio.PatternFill`` or ``GradientFill`` as the model's, or ``None``: no fill drawn."""
    if part is None:
        return None
    if hasattr(part, "pattern"):
        kind = part.pattern
        if kind != "solid" and kind not in TILES:  # none, or no patternType: nothing drawn
            return None
        fg = color(part.foreground, context)
        bg = color(part.background, context, background=True)
        if kind == "solid":
            # A solid fill in the system foreground color is no fill (indexed 64; automatic
            # taken alike; both unverified).
            return Pattern("solid", fg) if fg is not None else None
        return Pattern(kind, fg, bg)
    # By position only: stops at one position keep the file's order (a stable sort).
    stops = sorted(
        ((unit(stop.position), color(stop.color, context)) for stop in part.stops),
        key=lambda stop: stop[0],
    )
    if not stops:  # nothing to draw; what Excel does is not captured
        return None
    return Gradient(
        "path" if part.type == "path" else "linear",
        tuple((position, found or BLACK) for position, found in stops),
        degree=finite(part.degree) % 360,
        left=unit(part.left),
        right=unit(part.right),
        top=unit(part.top),
        bottom=unit(part.bottom),
    )


def color(source: Any, context: Context, *, background: bool = False) -> str | None:
    """A fill color as ``RRGGBB``, or ``None`` for automatic (unverified, §4.12).

    Indexed 64 is automatic, 65 white, and an index past the palette black. A missing color
    element, or one without attributes, is read as ``00000000`` (unverified; the reading of
    0.2.0, kept): for a foreground that is black, and a background in it is automatic (white).
    """
    if source is None or (
        source.indexed is None and source.theme is None and not source.auto and source.rgb is None
    ):
        return None if background else BLACK
    if source.indexed is not None:
        if source.indexed == _SYSTEM_FOREGROUND:
            return None
        if source.indexed == _SYSTEM_BACKGROUND:
            return WHITE
        return context.color(source) or BLACK
    if background and source.theme is None and not source.auto and source.rgb == _ZERO:
        return None
    return context.color(source)


def unread(part: Any, context: Context) -> Found:
    """Every fill is read."""
    return [], []
