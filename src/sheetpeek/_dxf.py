"""The differential-format layer: a cell's own style stacked over the partial formats beneath it.

A table style (§4.20), and later a conditional format, draws through :class:`DifferentialFormat`
records: partial formats, of which only what each sets applies. :func:`stack` takes them in the
order that stacks them, a later format's property winning over an earlier one's, and lays the
cell's own :class:`CellStyle` over the result property by property: the cell's own wins where it
differs from the default (a fill; a font color other than the Normal style's; bold, italic,
underline or strikethrough set), and the cell's own border sides replace
the formats' sides one side at a time. The font's name and size are always the cell's, since a
differential format in a table style sets neither.

Captured by the screen suite ``tables`` (``ARCHITECTURE.md`` §4.20), underline and
strikethrough taken alike with bold and italic included.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import replace

from sheetpeek._fill_border_model import Borders, Fill, Side
from sheetpeek._model import CellStyle, Font
from sheetpeek._table_model import DifferentialFormat


def stack(
    style: CellStyle,
    formats: Sequence[DifferentialFormat],
    *,
    own_color: bool,
    sides: Borders | None = None,
) -> CellStyle:
    """The style a cell is drawn in: its own ``style`` over ``formats`` and ``sides``.

    Args:
        style: The cell's own style, as the file states it.
        formats: The differential formats that apply to the cell, bottom first: a later one's
            property wins over an earlier one's.
        own_color: Whether the cell's font color is its own: any color other than the Normal
            style's font color as the file states it. ``style.font.color`` cannot tell, since
            it reads theme 1 and automatic alike as ``None``.
        sides: The sides the formats draw on this cell's edges, already placed (a format's
            ``horizontal`` and ``vertical`` lines inside its region, its outer sides on the
            region's edges), or ``None`` for none.

    Returns:
        ``style`` itself when nothing beneath it draws anything, else a new style.
    """
    fill: Fill | None = None
    color: str | None = None
    bold = italic = strike = None
    underline: str | None = None
    for found in formats:
        if found.fill is not None:
            fill = found.fill
        if found.color is not None:
            color = found.color
        if found.bold is not None:
            bold = found.bold
        if found.italic is not None:
            italic = found.italic
        if found.underline is not None:
            underline = found.underline
        if found.strike is not None:
            strike = found.strike
    borders = _sides(style.borders, sides)
    if (
        fill is None
        and color is None
        and bold is None
        and italic is None
        and underline is None
        and strike is None
        and borders is style.borders
    ):
        return style
    font = _font(style.font, own_color, color, bold, italic, underline, strike)
    if style.fill_spec is not None or fill is None:
        return replace(style, font=font, borders=borders)
    return replace(style, font=font, fill=fill.color, fill_spec=fill, borders=borders)


def _font(
    own: Font,
    own_color: bool,
    color: str | None,
    bold: bool | None,
    italic: bool | None,
    underline: str | None,
    strike: bool | None,
) -> Font:
    """The cell's own font over the formats' parts: each own part set wins."""
    return replace(
        own,
        color=own.color if own_color or color is None else color,
        bold=own.bold or bool(bold),
        italic=own.italic or bool(italic),
        underline=own.underline or (underline if underline != "none" else None),
        strike=own.strike or bool(strike),
    )


def _sides(own: Borders | None, beneath: Borders | None) -> Borders | None:
    """The cell's own sides, each replacing the side beneath it (``own`` when nothing changes)."""
    if beneath is None:
        return own
    if own is None:
        return Borders(beneath.left, beneath.right, beneath.top, beneath.bottom)
    chosen = (
        _side(own.left, beneath.left),
        _side(own.right, beneath.right),
        _side(own.top, beneath.top),
        _side(own.bottom, beneath.bottom),
    )
    if chosen == (own.left, own.right, own.top, own.bottom):
        return own
    return replace(own, left=chosen[0], right=chosen[1], top=chosen[2], bottom=chosen[3])


def _side(own: Side | None, beneath: Side | None) -> Side | None:
    return own if own is not None else beneath
