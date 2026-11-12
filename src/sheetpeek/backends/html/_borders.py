"""Borders: one layer of lines in sheet coordinates, over the fills and gridlines, under the text.

The layout resolves one side per edge (``RenderedSheet.extra["borders"]``, ``_pass_borders.py``);
this layer draws each run of equal sides along a line as one box. Widths, dashes and the place of
a line on its gridline are Excel's in screen points, which are CSS px at any page scale
(``ARCHITECTURE.md`` §4.12, captures on Excel 16.114):

- thin 1 px on the gridline; medium 2 px from 1 before it; thick 3 px from 1 before it; double
  two 1 px lines around a 1 px gap, from 1 before it;
- dashed, dotted, hair and the dash-dot styles as tiles in device pixels, two to a CSS px, their
  gaps white: no gridline shows under any border;
- slantDashDot as mediumDashDot, its gaps square (reported);
- where two lines meet, each reaches over the other's band, so corners are square.

Diagonals are an SVG line from corner to corner of the cell or merge. Each distinct line is one
generated class (``sp-b-`` and letters), not a public interface.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, TextIO

from sheetpeek._fill_border_model import BorderStyle, Diagonal, Side
from sheetpeek._pass_borders import rendered_borders
from sheetpeek.backends.html._patterns import Classes, svg_uri, tile_svg

if TYPE_CHECKING:
    from sheetpeek.backends.html._table import Page

_D, _G = "#", "."


def _dashes(*lengths: int) -> str:
    """One period of a dash pattern in device px: on, off, on, off..."""
    return "".join((_D if k % 2 == 0 else _G) * n for k, n in enumerate(lengths))


# Per style: the band before and after the gridline in CSS px, and the tile across the band in
# device rows (two to a CSS px) along one period of the line; None draws a solid band.
_BAND: dict[str, tuple[int, int]] = {
    "thin": (0, 1),
    "hair": (0, 1),
    "dotted": (0, 1),
    "dashed": (0, 1),
    "dashDot": (0, 1),
    "dashDotDot": (0, 1),
    "medium": (1, 1),
    "mediumDashed": (1, 1),
    "mediumDashDot": (1, 1),
    "mediumDashDotDot": (1, 1),
    "slantDashDot": (1, 1),
    "thick": (1, 2),
    "double": (1, 2),
}
_DASH_DOT = _dashes(18, 6, 6, 6)
_DASH_DOT_DOT = _dashes(18, 6, 6, 6, 6, 6)
_TILES: dict[str, tuple[str, ...] | None] = {
    "thin": None,
    "medium": None,
    "thick": None,
    "double": (_D, _D, _G, _G, _D, _D),
    "hair": ("#.", ".#"),
    "dotted": ("##..", "..##"),
    "dashed": ("###.", "##.#"),
    "dashDot": (_DASH_DOT,) * 2,
    "dashDotDot": (_DASH_DOT_DOT,) * 2,
    "mediumDashed": (_dashes(18, 6),) * 4,
    "mediumDashDot": (_DASH_DOT,) * 4,
    "mediumDashDotDot": (_DASH_DOT_DOT,) * 4,
    "slantDashDot": (_DASH_DOT,) * 4,
}
# Diagonals: stroke width and dashes in CSS px; hair is drawn as thin (captured).
_DIAGONAL: dict[str, tuple[float, str]] = {
    "thin": (1, ""),
    "hair": (1, ""),
    "dotted": (1, "1 1"),
    "dashed": (1, "2 2"),
    "dashDot": (1, "9 3 3 3"),
    "dashDotDot": (1, "9 3 3 3 3 3"),
    "medium": (2, ""),
    "mediumDashed": (2, "9 3"),
    "mediumDashDot": (2, "9 3 3 3"),
    "mediumDashDotDot": (2, "9 3 3 3 3 3"),
    "slantDashDot": (2, "9 3 3 3"),
    "thick": (3, ""),
    "double": (1, ""),
}


def band(style: BorderStyle) -> tuple[int, int]:
    """CSS px of a line before and after its gridline's first pixel."""
    return _BAND[style]


class Borders:
    """The border layer of a page."""

    def __init__(self, page: Page) -> None:
        self.classes = Classes("sp-b-")
        self.lines: list[str] = []
        sheet = page.sheet
        resolved = rendered_borders(sheet)
        rows, columns = sheet.rows, sheet.columns
        if resolved is None or not rows or not columns:
            return
        xs = [page.x(c.x_pt or 0) for c in columns]
        xs.append(page.x((columns[-1].x_pt or 0) + (columns[-1].width_pt or 0)))
        ys = [page.y(r.y_pt or 0) for r in rows]
        ys.append(page.y((rows[-1].y_pt or 0) + (rows[-1].height_pt or 0)))
        horizontal, vertical = resolved.horizontal, resolved.vertical

        def reach(sides: list[Side | None], end: bool) -> int:
            """How far a line reaches past a node to cover the bands of the lines across it."""
            found = [band(s.style)[1 if end else 0] for s in sides if s is not None]
            return max(found, default=1 if end else 0)

        for k, line in enumerate(horizontal):
            for first, last, side in _runs(line):
                across_first = [vertical[i][first] for i in (k - 1, k) if 0 <= i < len(rows)]
                across_last = [vertical[i][last + 1] for i in (k - 1, k) if 0 <= i < len(rows)]
                start = xs[first] - reach(across_first, False)
                end = xs[last + 1] + reach(across_last, True)
                before, after = band(side.style)
                self.line(side, start, ys[k] - before, end - start, before + after, False)
        for q in range(len(columns) + 1):
            column = [vertical[i][q] for i in range(len(rows))]
            for first, last, side in _runs(column):
                across_first = [horizontal[first][p] for p in (q - 1, q) if 0 <= p < len(columns)]
                across_last = [horizontal[last + 1][p] for p in (q - 1, q) if 0 <= p < len(columns)]
                start = ys[first] - reach(across_first, False)
                end = ys[last + 1] + reach(across_last, True)
                before, after = band(side.style)
                self.line(side, xs[q] - before, start, before + after, end - start, True)
        for diagonal in resolved.diagonals:
            self.diagonal(diagonal, xs, ys)

    def line(self, side: Side, left: int, top: int, width: int, height: int, upright: bool) -> None:
        if width <= 0 or height <= 0:
            return
        color = side.drawn_color
        tile = _TILES[side.style]
        if tile is None:
            body = f"background:#{color}"
        else:
            if upright:  # the tile turned: its rows become columns
                tile = tuple("".join(row[y] for row in tile) for y in range(len(tile[0])))
            body = (
                f"background:#FFFFFF {svg_uri(tile_svg(tile, color, 'FFFFFF'))};"
                f"background-size:{len(tile[0]) / 2:g}px {len(tile) / 2:g}px"
            )
        name = self.classes.name(body)
        self.lines.append(
            f'<i class="{name}" style="left:{left}px;top:{top}px;width:{width}px;'
            f'height:{height}px"></i>'
        )

    def diagonal(self, diagonal: Diagonal, xs: list[int], ys: list[int]) -> None:
        left, top = xs[diagonal.left], ys[diagonal.top]
        width, height = xs[diagonal.right + 1] - left + 1, ys[diagonal.bottom + 1] - top + 1
        if width <= 1 or height <= 1:
            return
        side = diagonal.side
        stroke, dashes = _DIAGONAL[side.style]
        dash = f" stroke-dasharray='{dashes}'" if dashes else ""
        paint = f"stroke='#{side.drawn_color}' stroke-width='{stroke:g}'{dash}"
        lines: list[str] = []
        for drawn, (x1, y1, x2, y2) in (
            (diagonal.up, (0, height, width, 0)),
            (diagonal.down, (0, 0, width, height)),
        ):
            if not drawn:
                continue
            if side.style == "double":  # two lines, 1 px either side of the diagonal
                lines.extend(
                    f"<line x1='{x1}' y1='{y1}' x2='{x2}' y2='{y2}' {paint} "
                    f"transform='translate({offset} {-offset if x1 == y1 else offset})'/>"
                    for offset in (-1, 1)
                )
            else:
                lines.append(f"<line x1='{x1}' y1='{y1}' x2='{x2}' y2='{y2}' {paint}/>")
        svg = (
            f"<svg xmlns='http://www.w3.org/2000/svg' width='{width}' height='{height}'>"
            f"{''.join(lines)}</svg>"
        )
        name = self.classes.name(f"background-image:{svg_uri(svg)}")
        self.lines.append(
            f'<i class="{name}" style="left:{left}px;top:{top}px;width:{width}px;'
            f'height:{height}px"></i>'
        )

    def css(self) -> str:
        """The layer's stylesheet, empty on a page without borders."""
        return self.classes.css()

    def write(self, out: TextIO) -> None:
        if self.lines:
            out.write(f'<div class="sp-borders">{"".join(self.lines)}</div>')


def _runs(line: list[Side | None] | tuple[Side | None, ...]) -> list[tuple[int, int, Side]]:
    """(first, last, side) of each run of equal sides along a line."""
    runs: list[tuple[int, int, Side]] = []
    for k, side in enumerate(line):
        if side is None:
            continue
        if runs and runs[-1][2] == side and runs[-1][1] == k - 1:
            runs[-1] = (runs[-1][0], k, side)
        else:
            runs.append((k, k, side))
    return runs
