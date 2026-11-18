"""Filter buttons: a square at a header cell's bottom right, over its text (§4.19).

The size is min(row height at the zoom, 17) of Excel's screen points, at 4/3 CSS pixels a screen
point as the cells are, so the button keeps Excel's proportion to the row: a 15 pt row's fills
it, a taller row's stops at 17 screen points and sits 1 above the row's bottom gridline. It does
not grow with the page's zoom beyond that. The face is a gradient from 254 to 245 gray and the
outline black at 40 %, which darkens a colored header as Excel's does; the glyph is drawn in a
34-unit box (the capture's pixels at 17 screen points) and scaled.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from sheetpeek._filter_model import FilterButtonState

if TYPE_CHECKING:
    from sheetpeek._rendered import RenderedColumn, RenderedRow, RenderedSheet
    from sheetpeek.backends.html import HTMLBackend
    from sheetpeek.backends.html._table import Page

_LARGEST = 17.0
"""The largest button, in screen points."""

CSS = """
.sheetpeek .sp-fb {
  position: absolute;
  box-sizing: border-box;
  border: 1px solid rgba(0, 0, 0, 0.4);
  border-radius: 2px;
  background: linear-gradient(#fefefe, #f5f5f5);
}
.sheetpeek .sp-fb svg { display: block; width: 100%; height: 100%; fill: #7b7b7b; }
"""

_TRIANGLE = "M9 13h16l-8 10z"
_SMALL = "M4 18h9l-4.5 6z"
_FUNNEL = "M16 8h14l-5.5 7v9l-3-2v-7z"
_UP = "M22 26v-13h-3l4.5-6 4.5 6h-3v13z"
_DOWN = "M22 7v13h-3l4.5 6 4.5-6h-3v-13z"
_UP_LEFT = "M7 26v-13h-3l4.5-6 4.5 6h-3v13z"
_DOWN_LEFT = "M7 7v13h-3l4.5 6 4.5-6h-3v-13z"

_GLYPHS: dict[FilterButtonState, str] = {
    "plain": _TRIANGLE,
    "filtered": _SMALL + _FUNNEL,
    "ascending": _SMALL + _UP,
    "descending": _SMALL + _DOWN,
    "filtered_ascending": _UP_LEFT + _FUNNEL,
    "filtered_descending": _DOWN_LEFT + _FUNNEL,
}

_LABELS: dict[FilterButtonState, str] = {
    "plain": "filter button",
    "filtered": "filter button: filtered",
    "ascending": "filter button: sorted ascending",
    "descending": "filter button: sorted descending",
    "filtered_ascending": "filter button: filtered, sorted ascending",
    "filtered_descending": "filter button: filtered, sorted descending",
}


def button(page: Page, row: RenderedRow, column: RenderedColumn, state: FilterButtonState) -> str:
    """The button of the header cell at ``row`` and ``column``, positioned on the page."""
    top_pt = row.y_pt or 0.0
    bottom_pt = top_pt + (row.height_pt or 0.0)
    right_pt = (column.x_pt or 0.0) + (column.width_pt or 0.0)
    size = round(min((row.height_pt or 0.0) * page.zoom / 100, _LARGEST) * 4 / 3)
    top, bottom = page.y(top_pt), page.y(bottom_pt)
    top = max(top, bottom - 1 - size)
    left = page.x(right_pt) - size
    return (
        f'<span class="sp-fb" style="left:{left}px;top:{top}px;width:{size}px;height:{size}px">'
        f'<svg viewBox="0 0 34 34" role="img" aria-label="{_LABELS[state]}">'
        f'<path d="{_GLYPHS[state]}"/></svg></span>'
    )


def css(sheet: RenderedSheet, backend: HTMLBackend) -> str:
    """The buttons' stylesheet, when the sheet draws one; else nothing."""
    if backend.buttons == "none":
        return ""
    if any(cell.button for row in sheet.rows for cell in row.cells):
        return CSS
    return ""
