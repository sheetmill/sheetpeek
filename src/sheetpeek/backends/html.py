"""HTML: a static page drawn from ``excel`` mode's geometry (``ARCHITECTURE.md`` §7.3).

A fixed-layout table carries the cells, their fills and their gridlines. Each text is an
absolutely positioned span inside its origin cell, placed in sheet coordinates against the
wrapper, with its clip window around it; the cells themselves are not positioned, so the
browser paints every background first and every text over it, as Excel paints spilled text
over a neighbor's fill. Every position is a whole CSS pixel rounded from its exact point, and
every size the difference of two rounded positions, so nothing drifts on large sheets. Calibri
runs are spaced from Carlito's widths to Excel's whole-point glyphs. The output has no script.

The custom properties ``--sheetpeek-*`` and the class names ``sheetpeek`` and ``sp-*`` are a
public interface for callers' stylesheets.
"""

from __future__ import annotations

import html
from typing import Literal, TextIO

from sheetpeek._carlito import ADVANCES, FIRST, UNITS_PER_EM
from sheetpeek._model import CellRange, Font
from sheetpeek._rendered import Mode, RenderedCell, RenderedSheet, TextRun

ColorScheme = Literal["light", "dark"]

# Calibri runs use a named face that is only Calibri or Carlito, installed locally, so that the
# letter-spacing computed for Calibri lands on Calibri's glyphs; without either, a local Arial
# or Helvetica (metric-compatible with each other and with Liberation Sans) scaled to Calibri's
# widths and given its vertical metrics, as Next.js builds fallback fonts.
_CALIBRI = "'sheetpeek-calibri', 'sheetpeek-calibri-fallback', sans-serif"
# size-adjust: Carlito's widths over Arial's on spreadsheet text (digits, numbers, labels),
# measured once; the overrides give the fallback Calibri's ascent 0.952 em and descent 0.269 em.
_FALLBACK_SIZE = {"regular": 91.6, "bold": 90.2, "italic": 91.4, "bold italic": 90.0}
_STYLE_NAMES = {  # style -> (CSS weight, CSS style, suffixes of the local font names)
    "regular": ("400", "normal", ("", "-Regular")),
    "bold": ("700", "normal", (" Bold", "-Bold")),
    "italic": ("400", "italic", (" Italic", "-Italic", " Oblique", "-Oblique")),
    "bold italic": ("700", "italic", (" Bold Italic", "-BoldItalic", " Bold Oblique")),
}
_LOCAL = ("Calibri", "Carlito")
_LOCAL_FALLBACK = ("Arial", "Helvetica", "Liberation Sans")


def _local(families: tuple[str, ...], suffixes: tuple[str, ...]) -> str:
    return ", ".join(f'local("{f}{s}")' for f in families for s in suffixes)


def _font_faces() -> str:
    rules = []
    for style, (weight, slant, suffixes) in _STYLE_NAMES.items():
        size = _FALLBACK_SIZE[style]
        rules.append(
            f'@font-face {{ font-family: "sheetpeek-calibri"; src: {_local(_LOCAL, suffixes)}; '
            f"font-weight: {weight}; font-style: {slant}; }}\n"
            '@font-face { font-family: "sheetpeek-calibri-fallback"; '
            f"src: {_local(_LOCAL_FALLBACK, suffixes)}; font-weight: {weight}; "
            f"font-style: {slant}; size-adjust: {size}%; "
            f"ascent-override: {95.2 / size * 100:.2f}%; "
            f"descent-override: {26.9 / size * 100:.2f}%; line-gap-override: 0%; }}\n"
        )
    return "".join(rules)


# Calibri's and Carlito's vertical metrics in em: ascent and descent, and where the baseline
# sits in a 1 em line box (half the leading above the ascent).
_ASCENT, _DESCENT = 0.952, 0.269
_BASELINE = 0.8418
# Placeholders until the screen capture: heading sizes in points.
_COLUMN_HEADING_HEIGHT = 16.0
_ROW_HEADING_DIGIT = 6.0

CSS = (
    _font_faces()
    + """\
.sheetpeek {
  --sheetpeek-paper: #ffffff;
  --sheetpeek-ink: #000000;
  --sheetpeek-grid: #d4d4d4;
  --sheetpeek-heading-paper: #f3f3f3;
  --sheetpeek-heading-ink: #444444;
  --sheetpeek-heading-grid: #c8c8c8;
  position: relative; display: inline-block; vertical-align: top;
  background: var(--sheetpeek-paper); color-scheme: light;
}
.sheetpeek[data-color-scheme="dark"] {
  --sheetpeek-heading-paper: #2b2b2b;
  --sheetpeek-heading-ink: #cccccc;
  --sheetpeek-heading-grid: #3d3d3d;
}
.sheetpeek.sp-nogrid { --sheetpeek-grid: transparent; }
.sheetpeek table {
  table-layout: fixed; border-collapse: separate; border-spacing: 0;
  border-right: 1px solid var(--sheetpeek-grid); border-bottom: 1px solid var(--sheetpeek-grid);
  color: var(--sheetpeek-ink);
}
.sheetpeek td {
  padding: 0; box-sizing: border-box;
  border-left: 1px solid var(--sheetpeek-grid); border-top: 1px solid var(--sheetpeek-grid);
}
.sheetpeek th {
  padding: 0; box-sizing: border-box; overflow: hidden; white-space: nowrap;
  font-size: 9pt; font-weight: normal; text-align: center; vertical-align: middle;
  background: var(--sheetpeek-heading-paper); color: var(--sheetpeek-heading-ink);
  border-left: 1px solid var(--sheetpeek-heading-grid);
  border-top: 1px solid var(--sheetpeek-heading-grid);
}
.sheetpeek .sp-crossed { border-left-color: transparent; }
.sheetpeek .sp-clip { position: absolute; overflow: hidden; }
.sheetpeek .sp-run {
  position: absolute; white-space: pre; line-height: 1;
  font-kerning: none; font-variant-ligatures: none;
}
.sheetpeek .sp-center { transform: translateX(-50%); }
"""
)
"""The stylesheet every page and fragment carries."""


class HTMLBackend:
    """A static HTML page, or a fragment, in ``excel`` mode.

    Excel draws white cells in its light and dark appearance alike, so the paper is white and
    the automatic ink black; ``color_scheme`` changes only the headings, as in Excel. Colors
    are CSS custom properties on ``.sheetpeek`` (``--sheetpeek-paper``, ``-ink``, ``-grid``,
    ``-heading-paper``, ``-heading-ink``, ``-heading-grid``), which a caller's stylesheet can
    set.

    Args:
        fragment: Write only the sheet (its stylesheet and a ``<div class="sheetpeek">``), for
            embedding, instead of a whole page.
        column_headings: Add Excel's column letters above the sheet.
        row_headings: Add Excel's row numbers before each row.
        color_scheme: ``"light"`` or ``"dark"``, for the headings; never read from the system.
    """

    name = "html"
    mode: Mode = "excel"

    def __init__(
        self,
        *,
        fragment: bool = False,
        column_headings: bool = False,
        row_headings: bool = False,
        color_scheme: ColorScheme = "light",
    ) -> None:
        if color_scheme not in ("light", "dark"):
            raise ValueError(f"color_scheme must be 'light' or 'dark', not {color_scheme!r}")
        self.fragment = fragment
        self.column_headings = column_headings
        self.row_headings = row_headings
        self.color_scheme = color_scheme

    def render(self, sheet: RenderedSheet, out: TextIO, /) -> None:
        """Write the page or the fragment to ``out``."""
        if not self.fragment:
            title = html.escape(sheet.sheet.name)
            background = "#1e1e1e" if self.color_scheme == "dark" else "#ffffff"
            out.write(
                '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
                '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
                f"<title>{title}</title>\n"
                f"<style>\n{CSS}body {{ margin: 8px; background: {background}; }}\n</style>\n"
                "</head>\n<body>\n"
            )
        else:
            out.write(f"<style>\n{CSS}</style>\n")
        _Page(sheet, self).write(out)
        if not self.fragment:
            out.write("</body>\n</html>\n")


def _px(points: float) -> int:
    """A sheet position in whole CSS pixels."""
    return round(points * 4 / 3)


class _Page:
    """One sheet's markup."""

    def __init__(self, sheet: RenderedSheet, backend: HTMLBackend) -> None:
        if sheet.mode != "excel":
            raise ValueError("the HTML backend draws a sheet laid out in excel mode")
        self.sheet = sheet
        self.backend = backend
        digits = max((len(str(r.index)) for r in sheet.rows), default=1)
        self.ox = _ROW_HEADING_DIGIT * max(3, digits) + 6 if backend.row_headings else 0.0
        self.oy = _COLUMN_HEADING_HEIGHT if backend.column_headings else 0.0
        self.normal = sheet.sheet.normal_font
        self.width = self.x(sheet.width_pt) + 1  # px, the wrapper's width: the table's

    def x(self, points: float) -> int:
        return _px(self.ox + points)

    def y(self, points: float) -> int:
        return _px(self.oy + points)

    def write(self, out: TextIO) -> None:
        sheet = self.sheet
        classes = "sheetpeek" if sheet.sheet.show_gridlines else "sheetpeek sp-nogrid"
        out.write(
            f'<div class="{classes}" data-color-scheme="{self.backend.color_scheme}">'
            f'<table style="width:{self.width}px;font:{_font(self.normal)}">\n'
        )
        out.write("<colgroup>")
        if self.backend.row_headings:
            out.write(f'<col style="width:{self.x(0)}px">')
        for c in sheet.columns:
            width = self.x((c.x_pt or 0) + (c.width_pt or 0)) - self.x(c.x_pt or 0)
            out.write(f'<col style="width:{width}px">')
        out.write("</colgroup>\n")
        if self.backend.column_headings:
            out.write(f'<thead><tr style="height:{self.y(0)}px">')
            if self.backend.row_headings:
                out.write('<th class="sp-corner"></th>')
            for c in sheet.columns:
                out.write(f'<th class="sp-colhead" scope="col">{c.letter}</th>')
            out.write("</tr></thead>\n")
        out.write("<tbody>\n")
        self.write_rows(out)
        out.write("</tbody></table></div>\n")

    def write_rows(self, out: TextIO) -> None:
        sheet = self.sheet
        grid = [row.cells for row in sheet.rows]
        # Each merge is one cell at its first visible position, spanning its visible area.
        spans: dict[CellRange, tuple[int, int, int, int]] = {}
        for i, cells in enumerate(grid):
            for p, cell in enumerate(cells):
                if cell.merge is not None:
                    first = spans.get(cell.merge, (i, p, i, p))
                    spans[cell.merge] = (first[0], first[1], i, p)
        holder: dict[tuple[int, int], tuple[int, int]] = {}  # sheet position -> grid position
        for i, cells in enumerate(grid):
            for p, cell in enumerate(cells):
                if cell.merge is not None:
                    top, left = spans[cell.merge][:2]
                    holder[(cell.row, cell.column)] = (top, left)
                    anchor = (cell.merge.min_row, cell.merge.min_col)
                    holder.setdefault(anchor, (top, left))
                else:
                    holder[(cell.row, cell.column)] = (i, p)
        runs: dict[tuple[int, int], list[TextRun]] = {}
        for i, row in enumerate(sheet.rows):
            for run in row.runs:
                runs.setdefault(holder.get(run.source, (i, 0)), []).append(run)
        for i, row in enumerate(sheet.rows):
            y0, y1 = row.y_pt or 0, (row.y_pt or 0) + (row.height_pt or 0)
            out.write(f'<tr style="height:{self.y(y1) - self.y(y0)}px">')
            if self.backend.row_headings:
                out.write(f'<th class="sp-rowhead" scope="row">{row.index}</th>')
            for p, cell in enumerate(row.cells):
                colspan = rowspan = 1
                if cell.merge is not None:
                    top, left, bottom, right = spans[cell.merge]
                    if (i, p) != (top, left):
                        continue  # covered by its merge
                    colspan, rowspan = right - left + 1, bottom - top + 1
                out.write(self.cell(i, p, cell, colspan, rowspan, runs.get((i, p), ())))
            out.write("</tr>\n")

    def cell(
        self,
        i: int,
        p: int,
        cell: RenderedCell,
        colspan: int,
        rowspan: int,
        runs: list[TextRun] | tuple[()],
    ) -> str:
        rows = self.sheet.rows
        style = []
        left = rows[i].cells[p - 1].fill if p else None
        above = rows[i - 1].cells[p].fill if i else None
        if cell.fill:
            # A fill covers the cell's own gridlines and its far ones (captured).
            style += [f"background:#{cell.fill}", f"border-color:#{cell.fill}"]
        else:
            if left:
                style.append(f"border-left-color:#{left}")
            if above:
                style.append(f"border-top-color:#{above}")
        attrs = ""
        if p and rows[i].crossed[p - 1]:
            attrs += ' class="sp-crossed"'
        if style:
            attrs += f' style="{";".join(style)}"'
        if colspan > 1:
            attrs += f' colspan="{colspan}"'
        if rowspan > 1:
            attrs += f' rowspan="{rowspan}"'
        return f"<td{attrs}>{''.join(self.run(run) for run in runs)}</td>"

    def run(self, run: TextRun) -> str:
        left, top, right, bottom = run.clip
        size = run.font.size
        spacing = _spacing(run)
        # A Calibri run spaced to Excel's width that lies inside its clip window needs no clip.
        inside = (
            run.font.name.lower() == "calibri"
            and not run.text.strip(_PRINTABLE)  # spaced from the Carlito table
            and left <= run.x
            and run.x + run.width <= right
            and top <= run.y - _ASCENT * size
            and run.y + _DESCENT * size <= bottom
        )
        # Each run is anchored where Excel anchors it: left text by its start, right text by its
        # end, centered text by its center, so another font changes only the far end.
        x0, x1, y0, y1 = (
            (0, self.width, 0, 0)
            if inside
            else (self.x(left), self.x(right), self.y(top), self.y(bottom))
        )
        if run.align == "left":
            anchor, classes = f"left:{self.x(run.x) - x0}px", "sp-run"
        elif run.align == "right":
            anchor, classes = f"right:{x1 - self.x(run.x + run.width)}px", "sp-run"
        else:
            center = (self.x(run.x) + self.x(run.x + run.width)) / 2 - x0
            anchor, classes = f"left:{center:g}px", "sp-run sp-center"
        baseline = self.y(run.y) - y0 - _BASELINE * size * 4 / 3
        style = [anchor, f"top:{_number(baseline)}px"]
        if _face(run.font) != _face(self.normal):
            style.append(f"font:{_font(run.font)}")
        decoration = _decoration(run.font)
        if decoration:
            style.append(f"text-decoration:{decoration}")
        if run.color:
            style.append(f"color:#{run.color}")
        if spacing:
            style.append(f"letter-spacing:{_number(spacing)}pt")
        span = f'<span class="{classes}" style="{";".join(style)}">{html.escape(run.text)}</span>'
        if inside:
            return span
        return (
            f'<span class="sp-clip" style="left:{x0}px;top:{y0}px;width:{x1 - x0}px;'
            f'height:{y1 - y0}px">{span}</span>'
        )


_PRINTABLE = "".join(chr(code) for code in range(32, 127))


def _number(value: float) -> str:
    """A CSS number to three decimals, without trailing zeros."""
    text = f"{value:.3f}".rstrip("0").rstrip(".")
    return "0" if text in ("-0", "") else text


def _face(font: Font) -> tuple[str, float, bool, bool]:
    return font.name.lower(), font.size, font.bold, font.italic


def _style(font: Font) -> str:
    if font.bold:
        return "bold italic" if font.italic else "bold"
    return "italic" if font.italic else "regular"


def _font(font: Font) -> str:
    family = _CALIBRI if font.name.lower() == "calibri" else f"'{_quoted(font.name)}', sans-serif"
    weight = "bold " if font.bold else ""
    italic = "italic " if font.italic else ""
    return f"{italic}{weight}{font.size:g}pt {family}"


def _quoted(name: str) -> str:
    return html.escape(name.replace("\\", "\\\\").replace("'", "\\'"))


def _decoration(font: Font) -> str:
    parts = []
    if font.underline in ("double", "doubleAccounting"):
        parts.append("underline double")
    elif font.underline:
        parts.append("underline")
    if font.strike:
        parts.append("line-through")
    return " ".join(parts)


def _spacing(run: TextRun) -> float:
    """Points of letter-spacing that make Calibri text as wide as Excel draws it.

    Excel for Mac sets glyphs on whole points; a browser draws Calibri's own advances, which
    Carlito shares. Only Calibri runs of printable ASCII are corrected.
    """
    font = run.font
    if font.name.lower() != "calibri" or not run.text:
        return 0.0
    table = ADVANCES[_style(font)]
    total = 0
    for char in run.text:
        k = ord(char) - FIRST
        if not 0 <= k < len(table):
            return 0.0
        total += table[k]
    own = total * font.size / UNITS_PER_EM
    return (run.width - own) / len(run.text)
