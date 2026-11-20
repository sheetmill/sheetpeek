"""The outline gutter: bars, dots, minus and plus buttons and level buttons, beside the headings.

The gutter is cells of the table: a column before the row headings and a row above the column
headings, each cell holding absolutely positioned segments, so it scrolls with the table. Its
sizes are screen-fixed, as Excel keeps them at any zoom: 4/3 CSS pixels a screen point, the
cells' scale at zoom 100, rounded to whole pixels, so the gutter keeps its proportion
to the rows; its lines (bars, borders, strokes) keep their width in CSS pixels, as gridlines and
borders do (``ARCHITECTURE.md`` §4.11; captured, ``outline``). The column gutter is the row
gutter's geometry transposed.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, TextIO

from sheetpeek._outline_model import OutlineMark, RenderedOutline
from sheetpeek._rendered import RenderedRow, RenderedSheet

if TYPE_CHECKING:
    from sheetpeek.backends.html import HTMLBackend
    from sheetpeek.backends.html._table import Page

# In screen points (captured, outline): lanes 12 apart from 2 in, 3 after the last; bars 2 wide,
# 3 into the lane, from 2 past the edge, with a 6 long tick; buttons 12 square filling the lane;
# dots 2 square. In CSS px, 4/3 each, rounded; the bars' width is a line's and stays 2 px.
_MARGIN, _LANE, _END = 3, 16, 4
_BAR, _INSET, _START, _TICK = 2, 4, 3, 8
_BUTTON, _DOT = 16, 3

# Colors (captured, outline): the gutter, its bars, ticks and dots, a button's border, paper and
# minus or plus stroke, in the light appearance and the dark one.
_CSS = """\
.sheetpeek {
  --sheetpeek-outline-paper: #e6e6e6;
  --sheetpeek-outline-ink: #999999;
  --sheetpeek-outline-button: #8a8a8a;
  --sheetpeek-outline-button-paper: #ffffff;
  --sheetpeek-outline-stroke: #333333;
}
.sheetpeek[data-color-scheme="dark"] {
  --sheetpeek-outline-paper: #555555;
  --sheetpeek-outline-ink: #999999;
  --sheetpeek-outline-button: #333333;
}
.sheetpeek td.sp-ol {
  position: relative; overflow: hidden; padding: 0; border: 0;
  background: var(--sheetpeek-outline-paper);
}
.sheetpeek .sp-ol i {
  position: absolute; display: block; background: var(--sheetpeek-outline-ink);
}
.sheetpeek .sp-olbutton, .sheetpeek .sp-ollevel {
  position: absolute; display: block; box-sizing: border-box; width: 16px; height: 16px;
  border: 1px solid var(--sheetpeek-outline-button); border-radius: 2px;
  background: var(--sheetpeek-outline-button-paper);
}
.sheetpeek .sp-olbutton i { background: var(--sheetpeek-outline-stroke); }
.sheetpeek .sp-ollevel {
  font: 11px/14px system-ui, sans-serif; text-align: center; font-weight: normal;
  color: var(--sheetpeek-outline-stroke);
}
"""


def drawn(sheet: RenderedSheet, backend: HTMLBackend) -> RenderedOutline | None:
    """The outline the page draws, or ``None``."""
    return sheet.outline if backend.buttons != "none" else None


def css(sheet: RenderedSheet, backend: HTMLBackend) -> str:
    """The gutter's stylesheet, only on a page that draws one."""
    return _CSS if drawn(sheet, backend) is not None else ""


def _size(lanes: int) -> int:
    return _MARGIN + _LANE * lanes + _END if lanes else 0


def offsets(sheet: RenderedSheet, backend: HTMLBackend) -> tuple[int, int]:
    """The row gutter's width and the column gutter's height, in px."""
    outline = drawn(sheet, backend)
    if outline is None:
        return 0, 0
    return _size(outline.row_lanes), _size(outline.column_lanes)


def _centered(length: int, size: int) -> int:
    """Where a box of ``size`` starts, centered on a row ``length`` long below its gridline."""
    return 1 + -(-(length - 1 - size) // 2)


def _segments(marks: list[OutlineMark], length: int, summary_after: bool) -> list[tuple[int, ...]]:
    """(across, along, across size, along size, kind) of what one row's gutter cell draws.

    ``across`` runs over the lanes, toward the headings; ``along`` along the row's height (or
    the column's width). Kind 0 is a bar, a tick or a dot; 1 a minus button; 2 a plus button.
    """
    found: list[tuple[int, ...]] = []
    button = _centered(length, _BUTTON)
    for mark in marks:
        x0 = _MARGIN + _LANE * (mark.lane - 1)
        x = x0 + _INSET
        if mark.kind == "bar":
            found.append((x, 0, _BAR, length, 0))
        elif mark.kind == "start" and summary_after:
            found += [(x, _START, _BAR, length - _START, 0), (x, _START, _TICK, _BAR, 0)]
        elif mark.kind == "start":
            end = length - _START  # summaries above or left: Excel's tick ends at the edge
            found += [(x, 0, _BAR, end, 0), (x, end, _TICK, _BAR, 0)]
        elif mark.kind == "open":
            if summary_after:
                found.append((x, 0, _BAR, button, 0))
            else:
                found.append((x, button + _BUTTON, _BAR, length - button - _BUTTON, 0))
            found.append((x0, button, _BUTTON, _BUTTON, 1))
        elif mark.kind == "closed":
            found.append((x0, button, _BUTTON, _BUTTON, 2))
        else:
            found.append((x, _centered(length, _DOT), _DOT, _DOT, 0))
    return [s for s in found if s[2] > 0 and s[3] > 0]


def _cell(marks: list[OutlineMark], length: int, summary_after: bool, rows: bool) -> str:
    """A gutter cell's markup; ``rows`` false transposes it for the column gutter."""
    parts = []
    for across, along, wide, long, kind in _segments(marks, length, summary_after):
        left, top, width, height = (
            (across, along, wide, long) if rows else (along, across, long, wide)
        )
        if kind == 0:
            style = f"left:{left}px;top:{top}px;width:{width}px;height:{height}px"
            parts.append(f'<i style="{style}"></i>')
        else:
            # strokes 8 long (6 screen points) and 1 thick, centered in the 14 px inside the
            # border
            strokes = '<i style="left:3px;top:6.5px;width:8px;height:1px"></i>'
            if kind == 2:
                strokes += '<i style="left:6.5px;top:3px;width:1px;height:8px"></i>'
            parts.append(f'<b class="sp-olbutton" style="left:{left}px;top:{top}px">{strokes}</b>')
    return f'<td class="sp-ol">{"".join(parts)}</td>'


def _levels(count: int, along: int, rows: bool) -> str:
    """Level buttons 1 to ``count``: across a band ``along`` px deep (rows), or down one."""
    parts = []
    for k in range(count):
        lane = _MARGIN + _LANE * k
        # centered on the headings' band with no gridline term, rounded down the band: Excel's
        # buttons sit 5 points into bands of 21 and 22 (captured, outline)
        middle = -(-(along - _BUTTON) // 2)
        left, top = (lane, middle) if rows else (middle, lane)
        parts.append(f'<b class="sp-ollevel" style="left:{left}px;top:{top}px">{k + 1}</b>')
    return "".join(parts)


def write_column(page: Page, out: TextIO) -> None:
    """The ``<col>`` of the row gutter."""
    if page.gx:
        out.write(f'<col style="width:{page.gx}px">')


def write_header(page: Page, out: TextIO) -> None:
    """The column gutter: a row above the column headings."""
    outline = drawn(page.sheet, page.backend)
    if outline is None or not page.gy:
        return
    out.write(f'<tr style="height:{page.gy}px">')
    if page.gx:
        out.write('<td class="sp-ol"></td>')
    if page.backend.row_headings:
        width = page.x(0) - page.gx
        levels = _levels(outline.column_lanes, width, rows=False)
        out.write(f'<td class="sp-ol">{levels}</td>')
    by_index = _by_index(outline.columns)
    for c in page.sheet.columns:
        width = page.x((c.x_pt or 0) + (c.width_pt or 0)) - page.x(c.x_pt or 0)
        out.write(_cell(by_index.get(c.index, []), width, outline.summary_right, rows=False))
    out.write("</tr>")


def lead(page: Page) -> str:
    """The row gutter's cell in the column headings' row, with the row level buttons."""
    outline = drawn(page.sheet, page.backend)
    if outline is None or not page.gx:
        return ""
    levels = _levels(outline.row_lanes, page.y(0) - page.gy, rows=True)
    return f'<td class="sp-ol">{levels}</td>'


def write_row(page: Page, row: RenderedRow, out: TextIO) -> None:
    """A row's gutter cell."""
    outline = drawn(page.sheet, page.backend)
    if outline is None or not page.gx:
        return
    y0, y1 = row.y_pt or 0, (row.y_pt or 0) + (row.height_pt or 0)
    length = page.y(y1) - page.y(y0)
    marks = page.outline_rows.get(row.index, [])
    out.write(_cell(marks, length, outline.summary_below, rows=True))


def row_marks(sheet: RenderedSheet, backend: HTMLBackend) -> dict[int, list[OutlineMark]]:
    """The row gutter's marks by row."""
    outline = drawn(sheet, backend)
    return _by_index(outline.rows) if outline is not None else {}


def _by_index(marks: tuple[OutlineMark, ...]) -> dict[int, list[OutlineMark]]:
    found: dict[int, list[OutlineMark]] = {}
    for mark in marks:
        found.setdefault(mark.index, []).append(mark)
    return found
