"""``excel`` mode: a sheet laid out in points, as Excel for Mac lays it out.

Columns take xlformat's widths in points and rows the heights of ``ARCHITECTURE.md`` §4.7. Each
cell is formatted at its column's width in dialog units (a merge at its merged width), then
placed by its alignment and spilled into free neighbors by the rules of §4.1, in points. The
result is a list of :class:`~sheetpeek.TextRun` per row, with each run's clip window, and the
column boundaries text crosses. Two spills may take the same free cell; both are kept, as Excel
draws both.
"""

from __future__ import annotations

import bisect
import functools
import math
import re
from dataclasses import replace
from typing import Any

import xlformat

from sheetpeek._layout import (
    _UNDRAWN,
    Draft,
    _column_hidden,
    _edge_columns,
    _Fonts,
    _Formatter,
    _Grid,
    _Origin,
    _Shown,
    check_area,
)
from sheetpeek._model import Column, Font, Row, Sheet
from sheetpeek._pass_borders import slant_shift, slants
from sheetpeek._rendered import RenderedCell, RenderedColumn, RenderedRow, TextRun
from sheetpeek._rows import (
    TALLEST_ROW,
    fits_cells,
    font_row_height,
    print_baseline_offsets,
    row_height,
    wrapped_height,
)
from sheetpeek._runs import (
    Text,
    drawn_spans,
    fitted_height,
    line_font,
    line_metrics,
    runs_shrink_step,
    wrap,
)
from sheetpeek._runs_model import TextSpan
from sheetpeek._text_model import TextDirection, TextRule
from sheetpeek._text_styles import (
    SHRINK_STEPS,
    accounting,
    right_to_left,
    row_extras,
    rules,
    scaled,
    shrink_step,
    vert_scale,
    vert_shift,
)

_SPREAD = ("justify", "distributed")
_WORD = re.compile(r"\S+")
ROTATED_INSETS = (3.0, 4.0)
"""Left and right insets of angled text's box, fitted to the rotation capture (§4.17)."""
SLANT_INSETS = (2.0, 1.5)
"""How far inside the slanted sides angled text's ascent and descent lie, across the line,
fitted to the rotation capture (§4.17)."""
ROTATED_EXTRA = 7.0
"""Points a row fitted to angled text adds to its footprint, fitted the same way (§4.17)."""
STACKED_EXTRA = 2.0
"""Points a row fitted to stacked text adds to its lines (captured, §4.17)."""
TOP_DROP = 3.0
"""Points below the cell's top that top-aligned angled text's highest corner sits (§4.17)."""
STACKED_AXIS = (1.43, 0.758)
"""Stacked left or right text's axis from the near edge: points, and points per point of size."""


def lay_out(draft: Draft) -> tuple[tuple[RenderedColumn, ...], tuple[RenderedRow, ...]]:
    """Lay out the draft's columns and rows in points: rendered columns and rows."""
    sheet, columns, rows, notes, fonts = (
        draft.sheet,
        draft.kept_columns,
        draft.kept_rows,
        draft.notes,
        draft.fonts,
    )
    normal = fonts.get(sheet.normal_font)

    def metrics(column: Column) -> xlformat.ColumnMetrics:
        return xlformat.file_column_metrics(
            column.width,
            base_col_width=sheet.base_column_width,
            default_col_width=sheet.default_column_width,
            normal_font=normal,
        )

    base = metrics(Column(1))
    if sheet.ref is None or not columns or not rows:
        rendered = tuple(
            RenderedColumn(c.index, None, _column_hidden(c), False, 0.0, _points(c, metrics))
            for c in columns
        )
        return _placed(rendered), ()
    # Added columns are at the default width; a text spills at most twice a digit per character.
    before, after = _edge_columns(sheet, max(1, int(base.points // (2 * base.digit_width))))
    every = before + columns + after
    widths = [_points(c, metrics) for c in every]
    added = set(range(len(before))) | set(range(len(before) + len(columns), len(every)))
    grid = _Grid(
        sheet,
        rows,
        every,
        [0] * len(every),
        added,
        0,
        False,
        _Formatter(sheet, notes, fonts),
        draft.display_styles,
    )

    blocks: dict[tuple[int, int], float | None] = {}

    def units(cols: list[int]) -> float | None:
        key = (cols[0], cols[-1])
        if key not in blocks:
            total = sum(widths[p] for p in cols)
            blocks[key] = (
                xlformat.column_metrics_from_points(total, normal_font=normal).width
                if total > 0
                else None
            )
        return blocks[key]

    grid.units = units
    read = [p for p in range(len(every)) if p not in added]
    fitted = [_fitted(grid, i, row, sheet, read, widths, base, fonts) for i, row in enumerate(rows)]
    heights = [f[0] for f in fitted]
    bottoms = [f[1] for f in fitted]
    layout = _Points(grid, widths, heights, base, fonts, notes, sheet.normal_font, bottoms)
    rendered_rows, kept = layout.render()
    x0 = layout.xs[kept[0]]
    rendered_columns = tuple(
        RenderedColumn(
            every[p].index,
            None,
            _column_hidden(every[p]),
            p in added,
            layout.xs[p] - x0,
            widths[p],
        )
        for p in kept
    )
    return rendered_columns, rendered_rows


def _fitted(
    grid: _Grid,
    i: int,
    row: Row,
    sheet: Sheet,
    read: list[int],
    widths: list[float],
    base: xlformat.ColumnMetrics,
    fonts: _Fonts,
) -> tuple[float, tuple[float, float, float, float]]:
    """A row's height, fitted to its cells as Excel fits it on opening (§4.7), and its bottom.

    The bottom is ``(B_row, I_row)``: bottom-aligned text in the row sits on one baseline,
    ``B_row`` above the row's bottom, the largest bottom offset over the Normal font and the
    fonts of the row's cells, empty styled cells and accounting underlines included (§3.6); and
    ``I_row``, the largest line inset ``B - 2`` over the same fonts, which lifts wrapped text
    in a face with a smaller one by the difference (captured, ``text_style_more`` group
    ``faces``: three wrapped lines of five faces beside Calibri 17 in a 90 pt row). A fixed row
    has a bottom too. The cells counted are those whose fonts fit the row (captured for cells
    with values; empty styled cells and the other kinds unverified).

    The bottom is given twice: with the fonts of rich-text runs, and with each cell's own font
    alone. A rich-text cell's last line sits on its own line's bottom, not on a larger run's
    of another line (captured, ``runs`` C8, C10 and C12), so it is placed by the second, which
    also leaves out other cells' runs (unverified).
    """
    below, inset, ascent = fonts.bottom(sheet.normal_font)
    plain_below, plain_inset = below, inset
    if not fits_cells(row, sheet):
        for p in read:
            cell = grid.cell(i, p)
            style = cell.style
            if (style.rotation or style.stacked) and cell.kind != "empty" and _upright(style):
                continue
            b, n, _ = fonts.bottom(cell.font)
            plain_below, plain_inset = max(plain_below, b), max(plain_inset, n)
            own = [cell.font]
            if cell.runs:
                content = _UNDRAWN.sub("", str(cell.value))
                own += [span.font for span in drawn_spans(cell, content) or ()]
            for font in own:
                b, n, _ = fonts.bottom(font)
                below, inset = max(below, b), max(inset, n)
        return row_height(row, sheet), (below, inset, plain_below, plain_inset)
    height = row_height(row, sheet)
    styled = []
    cell_fonts: list[Font] = []  # fonts of the row's cells that do not wrap
    wrapped: list[tuple[Font, int]] = []  # each plain wrapped cell's font and lines
    for p in read:
        cell = grid.cell(i, p)
        # A cell with rich-text runs counts its runs' fonts and its own, even when every run is
        # smaller (captured, runs: 9 pt runs in a 20 pt cell fit the row to 26 pt; §4.16).
        spans = drawn_spans(cell, _UNDRAWN.sub("", str(cell.value))) if cell.runs else None
        own = [span.font for span in spans] if spans else [cell.font]
        # A merged cell's font counts, its wrapped lines do not (captured).
        merged = (row.index, grid.columns[p].index) in grid.merge_of
        turned = cell.kind != "empty" and (cell.style.rotation or cell.style.stacked) and not merged
        if not (turned and _upright(cell.style)):
            # Text at 90 degrees fits the row by its width alone, at least the Normal font's
            # row: its own font's row does not count (captured, row_height_fitted).
            height = max(height, *map(fonts.row_height, own), fonts.row_height(cell.font))
            for font in (*own, cell.font):
                b, n, a = fonts.bottom(font)
                below, inset, ascent = max(below, b), max(inset, n), max(ascent, a)
            b, n, _ = fonts.bottom(cell.font)
            plain_below, plain_inset = max(plain_below, b), max(plain_inset, n)
        if cell.kind != "empty":
            styled += [f for f in own if f.vert_align]
        if turned:
            height = max(height, _turned_height(grid, cell, p, fonts))
            continue
        # Distributed and justified text, across or down, wraps when drawn, and Excel 16.115
        # fits the row to the lines drawn, wrap flag or not (captured, text_style, text_style_more).
        style = cell.style
        wraps = style.wrap or style.h_align in _SPREAD or style.v_align in _SPREAD
        if cell.kind == "text" and wraps and widths[p] and not merged:
            budget = widths[p] - base.left_inset - base.right_inset
            # between the indents, as drawn
            budget -= (
                _indents(cell.style.h_align)
                * cell.style.indent
                * indent_points(fonts.get(sheet.normal_font))
            )
            if spans:
                text = Text.of_spans(spans, fonts.get, breaks=True)
                height = max(height, fitted_height(text, wrap(text.text, budget, text.width)))
            else:
                lines = _wrap(str(cell.value), budget, fonts.get(cell.font))
                height = max(height, wrapped_height(len(lines), cell.font))
                wrapped.append((cell.font, len(lines)))
            continue
        cell_fonts += own
    # The row holds its fonts' largest ascent and largest bottom, an accounting underline's
    # points included (xlformat's row model; captured, row_height_fitted group styled: an
    # accounting underline at 6 to 10 pt under Calibri 11 makes a row of 17 or 18 pt).
    height = max(height, ascent + below)
    # Superscript and subscript grow the row (§4.13); the row fits the unshrunk size of shrunk
    # text, so shrink adds nothing.
    for font, extra in row_extras(styled):
        height = max(height, font_row_height(font) + extra)
    if wrapped:
        height = max(height, _row_model(sheet.normal_font, frozenset(cell_fonts), tuple(wrapped)))
    return min(height, TALLEST_ROW), (below, inset, plain_below, plain_inset)


_INDENT_SPACES = 3
"""Spaces of the Normal font in one indent level (captured, ``text_steps``, §4.6)."""


def indent_points(normal: xlformat.Font | None) -> float:
    """Points one indent level moves text: three spaces of the Normal font (§4.6).

    The cell's own font does not change it.
    """
    return _INDENT_SPACES * xlformat.text_width(" ", font=normal)


def _indents(align: str) -> int:
    """How many indents narrow a wrapped line: both ends when distributed, one left or right."""
    return 2 if align == "distributed" else 1 if align in ("left", "right") else 0


@functools.lru_cache(maxsize=1024)
def _row_model(
    normal: Font, cell_fonts: frozenset[Font], wrapped: tuple[tuple[Font, int], ...]
) -> float:
    """The height xlformat's captured row model gives, or 0 when a font is not measured.

    Wrapped lines take a point each beyond the font's height, one line included, and a larger
    font in the row adds the difference of the faces' line insets (xlformat's
    ``capture/row_heights.py``: Excel's heights for one to six wrapped lines in each face under
    each Normal font, and beside unwrapped cells in other fonts, under "Row heights" in its
    CAPTURING.md; several wrapped cells in one row are not captured). Rows repeat the same
    fonts, so the result is kept by them.
    """

    def spec(font: Font) -> xlformat.Font:
        underline = font.underline if accounting(font) else None  # a single one adds nothing
        return xlformat.Font(font.name, font.size, font.bold, font.italic, underline)

    try:
        return xlformat.row_metrics(
            spec(normal),
            cell_fonts=[spec(f) for f in cell_fonts],
            wrapped=[(spec(f), lines) for f, lines in wrapped],
        ).height
    except LookupError:  # a face or size the model has not measured: the table's heights
        return 0.0


def _turned_height(grid: _Grid, cell: Any, p: int, fonts: _Fonts) -> float:
    """The row height Excel fits to angled or stacked text (§4.17), at least the font's.

    Angled: the text's width times the sine of its angle, its size times the cosine, and
    ``ROTATED_EXTRA``, to the nearest point (captured over three lengths at 15, 45, 75 and 90
    degrees and 9, 11, 14 and 20 pt: within 2 pt below 90 degrees). At exactly 90 or -90 degrees,
    :func:`_upright_height`. Wrapped, justified or distributed angled text leaves the row as it
    is. Stacked: a font line a character and ``STACKED_EXTRA``.
    """
    shown = grid.show(cell, [p])
    content = _UNDRAWN.sub("", (shown.full_text or shown.text).replace("\n", ""))
    spans = drawn_spans(cell, content) if cell.runs else None
    own = [span.font for span in spans] if spans else [cell.font]
    line = max(map(font_row_height, own))
    if cell.style.stacked:
        return max(line, len(content) * line + STACKED_EXTRA)
    style = cell.style
    if style.wrap or style.h_align in _SPREAD or style.v_align in _SPREAD:
        return line  # wrapped angled text does not grow its row (captured, rotation turned)
    if spans:
        width = Text.of_spans(spans, fonts.get, breaks=False).width(0, len(content))
    else:
        width = xlformat.text_width(content, font=fonts.get(cell.font))
    size = max(font.size for font in own)
    if _upright(style):
        return _upright_height(width, size)  # the Normal font's row bounds it (the caller)
    angle = math.radians(cell.style.rotation)
    turned = width * abs(math.sin(angle)) + size * abs(math.cos(angle)) + ROTATED_EXTRA
    return max(line, float(round(turned)))


def _stacked_axis(size: float) -> float:
    """Points from the near edge to the axis stacked left or right text is centered on (§4.17)."""
    constant, share = STACKED_AXIS
    return constant + share * size


def _upright(style: Any) -> bool:
    """Text turned exactly 90 or -90 degrees, on one line: the path ``_upright_height`` fits."""
    return (
        abs(style.rotation) == 90
        and not style.stacked
        and not (style.wrap or style.h_align in _SPREAD or style.v_align in _SPREAD)
    )


def _upright_height(width: float, size: float) -> float:
    """The row Excel fits to text at exactly 90 or -90 degrees (§4.17, fitted rows).

    ``W + max(0, floor(W / 30) - 2) + 7 + 2·floor((S - 9) / 8)``, with ``W`` the width in whole
    advances and ``S`` the font size, at most ``TALLEST_ROW``: past 90 points of text the row
    gains a point every 30. The pad is 7 to 16 pt, 9 at 17 to 24, 11 at 25 to 32 and so on
    (captured, ``row_height_fitted`` group ``turned`` and ``rotation``: every row at 90 and -90
    degrees at the sizes xlformat measures; the 7 held below 9 pt is unverified).
    """
    whole = math.floor(width + 1e-9)
    pad = ROTATED_EXTRA + 2 * max(0, math.floor((size - 9) / 8))
    return min(TALLEST_ROW, whole + max(0, whole // 30 - 2) + pad)


def _points(column: Column, metrics: Any) -> float:
    return 0.0 if _column_hidden(column) else float(metrics(column).points)


def _placed(columns: tuple[RenderedColumn, ...]) -> tuple[RenderedColumn, ...]:
    """Columns with their x positions as running sums of their widths."""
    x, result = 0.0, []
    for c in columns:
        result.append(RenderedColumn(c.index, None, c.hidden, c.added, x, c.width_pt))
        x += c.width_pt or 0.0
    return tuple(result)


class _Points:
    """Places every origin's text in points: runs, clips, spill and crossed boundaries."""

    def __init__(
        self,
        grid: _Grid,
        widths: list[float],
        heights: list[float],
        base: xlformat.ColumnMetrics,
        fonts: _Fonts,
        notes: Any = None,
        normal: Font | None = None,
        bottoms: list[tuple[float, float, float, float]] | None = None,
    ) -> None:
        self.grid = grid
        # each row's (B_row, I_row), with runs' fonts and without (``_fitted``)
        self.bottoms = bottoms
        self.notes = notes
        self.normal = normal
        self.widths = widths
        self.heights = heights
        self.left_inset = base.left_inset
        self.right_inset = base.right_inset
        self.indent = indent_points(fonts.get(normal) if normal is not None else None)
        self.fonts = fonts
        self.xs = _running(widths)
        self.ys = _running(heights)
        self.no_claims: set[tuple[int, int]] = set()  # both competing spills take a cell
        self._metrics: dict[tuple[int, int], tuple[Font, tuple[float, float, float]]] = {}

    def plain_metrics(self, font: Font, lines: int) -> tuple[float, float, float]:
        """``line_metrics`` of text in one font, by the font's identity (see ``_Fonts``)."""
        key = (id(font), lines)
        hit = self._metrics.get(key)
        if hit is None or hit[0] is not font:
            hit = self._metrics[key] = (font, line_metrics([font], lines))
        return hit[1]

    def row_below(self, i: int, own: float, face: Font, wraps: bool, runs: bool = False) -> float:
        """Where a bottom-aligned block's last baseline sits above row ``i``'s bottom (§3.6).

        On the row's shared bottom ``B_row``, at least the line's own (an accounting run), and
        for wrapped text the difference of the row's line inset and its face's higher
        (xlformat's row model). A merge's text takes its last row's (unverified).
        """
        if self.bottoms is None:
            return own
        full_below, full_inset, plain_below, plain_inset = self.bottoms[i]
        below, inset = (plain_below, plain_inset) if runs else (full_below, full_inset)
        below = max(below, own)
        if wraps:
            below += max(0.0, inset - self.fonts.bottom(face)[1])
        return below

    def slant_spans(self, origins: list[Any]) -> list[tuple[int, int]]:
        """The columns each slanted border box reaches, kept as a spill keeps them (§4.17)."""
        found = []
        for origin in origins:
            if len(origin.rows) > 1 or len(origin.cols) > 1 or not slants(origin.cell.style):
                continue
            i, p = origin.rows[0], origin.cols[0]
            shift = slant_shift(origin.cell.style, self.heights[i])
            left = self.xs[p] + min(shift, 0.0)
            right = self.xs[p] + self.widths[p] + max(shift, 0.0)
            low = max(bisect.bisect_right(self.xs, left) - 1, 0)
            high = min(bisect.bisect_left(self.xs, right) - 1, len(self.widths) - 1)
            found.append((low, max(high, p)))
        return found

    def render(self) -> tuple[tuple[RenderedRow, ...], list[int]]:
        grid = self.grid
        origins, shown = grid.origins()
        runs: dict[int, list[TextRun]] = {}
        crossed: set[tuple[int, int]] = set()  # (row position, p): boundary between p and p + 1
        spans: list[tuple[int, int]] = []
        lines_of: dict[int, int] = {}
        spill_of: dict[tuple[int, int], tuple[int, int]] = {}
        for origin in origins:
            placed, low, high, count = self.place(origin)
            if not placed:
                continue
            runs.setdefault(origin.rows[0], []).extend(placed)
            spans.append((low, high))
            for i in origin.rows:
                for p in range(low, high):
                    crossed.add((i, p))
            spill_of[(origin.cell.row, origin.cell.column)] = (
                origin.cols[0] - low,
                high - origin.cols[-1],
            )
            if len(origin.rows) == 1:
                i = origin.rows[0]
                lines_of[i] = max(lines_of.get(i, 1), count)
        spans.extend(self.slant_spans(origins))
        real = [p for p in range(len(grid.columns)) if p not in grid.added]
        reached = [p for low, high in spans for p in (low, high) if p in grid.added]
        kept = list(range(min([*reached, real[0]]), max([*reached, real[-1]]) + 1))
        check_area(grid.sheet, len(grid.rows), len(kept))
        x0 = self.xs[kept[0]]
        result = []
        for i, row in enumerate(grid.rows):
            cells = []
            for p in kept:
                cell = grid.cell(i, p)
                position = (row.index, grid.columns[p].index)
                merge = grid.merge_of.get(position)
                show = shown.get(position, _Shown())
                cells.append(
                    RenderedCell(
                        cell=grid.file_cell(i, p),
                        text=show.text,
                        full_text=show.full_text,
                        overflow=show.overflow,
                        align=show.align,
                        format_color=show.format_color,
                        merge=merge,
                        spill=spill_of.get(position, (0, 0)),
                        lines=(),
                        source=None,
                        fill=grid.anchor_fill(merge) if merge else cell.fill,
                        font=None,
                        color=None,
                        display_style=cell.style,
                    )
                )
            row_runs = tuple(_shifted(run, x0) for run in runs.get(i, ()))
            result.append(
                RenderedRow(
                    row.index,
                    row,
                    lines_of.get(i, 1),
                    tuple(cells),
                    (),
                    self.ys[i],
                    self.heights[i],
                    row_runs,
                    tuple((i, p) in crossed for p in kept[:-1]),
                )
            )
        return tuple(result), kept

    def place(self, origin: _Origin) -> tuple[list[TextRun], int, int, int]:
        """The runs of one origin, the columns its text spans, and its line count."""
        cell, shown = origin.cell, origin.shown
        first, last = origin.cols[0], origin.cols[-1]
        if not shown.text:
            return [], first, last, 0
        if cell.style.stacked:
            return self.place_stacked(origin)
        if cell.style.rotation:
            return self.place_turned(origin)
        x0, x1 = self.xs[first], self.xs[last] + self.widths[last]
        top = self.ys[origin.rows[0]]
        bottom = self.ys[origin.rows[-1]] + self.heights[origin.rows[-1]]
        style = cell.style
        cell_font = cell.font
        font = self.fonts.get(cell_font)  # at the drawn size of superscript and subscript
        size = cell_font.size * vert_scale(cell_font)
        align = shown.align
        budget = (x1 - x0) - self.left_inset - self.right_inset
        v_align = style.v_align
        # Distributed and justified text wraps, horizontally or vertically (§3.6, §4.13).
        wraps = cell.kind == "text" and (
            style.wrap or align in ("justify", "distributed") or v_align in _SPREAD
        )
        indent = style.indent * self.indent
        spans = drawn_spans(cell, shown.text)  # rich-text runs (§4.16)
        single = len(origin.rows) == len(origin.cols) == 1  # a merge is not shrunk (§4.13)
        shrink = 1.0
        if style.shrink and not wraps and single and align in ("left", "right", "center"):
            # Shrink to fit replaces #### with the value, drawn small enough to fit (§4.13).
            if shown.overflow:
                shown.text = shown.full_text
                shown.overflow = False
            if spans:
                whole = Text.of_spans(spans, self.fonts.get, breaks=False)
                step = runs_shrink_step(whole, budget - indent, SHRINK_STEPS)
            else:
                step = shrink_step(shown.text, font, budget - indent)
            shrink = step / SHRINK_STEPS
            font = scaled(font, shrink)
            size *= shrink
        # a line break in a cell without wrapping is not drawn
        if spans:
            text = Text.of_spans(spans, self.fonts.get, breaks=wraps, scale=shrink)
        else:
            content = shown.text.replace("\r", "")
            text = Text.plain(content if wraps else content.replace("\n", ""), cell_font, font)
        # Distributed lines are broken between the indents at both ends (captured, text_style
        # lone); left and right lines inside their one indent (captured, text_style_more).
        line_budget = budget - _indents(align) * indent
        lines = wrap(text.text, line_budget, text.width) if wraps else [(0, len(text.text))]
        n = len(lines)
        if spans:
            metrics = [line_metrics(text.fonts(a, b), n) for a, b in lines]
        else:
            # Shrunk text is placed by its drawn size's metrics (captured, text_style_more
            # group shrunk: top and center at 4/7 and 6/7; 2/7 sits as about 6 pt would)
            placed_font = (
                replace(cell_font, size=cell_font.size * shrink) if shrink < 1 else cell_font
            )
            metrics = [self.plain_metrics(placed_font, n)] * n
        face = line_font(text.fonts(*lines[-1])) if spans else cell_font
        row_below = self.row_below(origin.rows[-1], metrics[-1][2], face, wraps, bool(spans))
        if n == 1 and v_align == "bottom" and metrics[0][0] <= bottom - top:
            # One bottom-aligned line, as most cells are: ``_baselines``' arithmetic, inline.
            baselines = [bottom - row_below]
        else:
            baselines = _baselines(metrics, top, bottom, v_align, row_below)
        shift = 0.0 if spans else vert_shift(cell_font)
        color = shown.format_color or cell_font.color
        direction: TextDirection = (
            "rtl" if style.reading_order != 1 and right_to_left(text.text) else "ltr"
        )
        source = (cell.row, cell.column)
        low, high = first, last
        own = (x0, x1) if cell.kind == "text" else None  # an accounting underline's span
        # Each piece: start, end and copies of the text, x, anchor, baseline, whether first, and
        # the line's width when the piece is the whole line once (measured already).
        placed: list[tuple[int, int, int, float, str, float, bool, float | None]] = []
        for k, ((a, b), baseline) in enumerate(zip(lines, baselines, strict=True)):
            baseline += shift
            width = text.width(a, b)
            words = [m.span() for m in _WORD.finditer(text.text, a, b)] if align in _SPREAD else []
            pieces: list[tuple[int, int, int, float, str]]
            if align == "fill":
                copies = max(1, math.floor(budget / width)) if width else 1
                pieces = [(a, b, copies, x0 + self.left_inset, "left")]
            elif align == "distributed" or (align == "justify" and k < n - 1 and len(words) > 1):
                distributed = align == "distributed"
                # A line broken at a space keeps it as a last gap (captured, text_style lone).
                trailing = b < len(text.text) and text.text[b] == " "
                pieces = self.spread(
                    words, text, x0, x1, distributed, indent if distributed else 0.0, trailing
                )
            elif align == "right" and not shown.overflow:
                pieces = [(a, b, 1, x1 - self.right_inset - indent - width, "right")]
            elif align == "center" or shown.overflow:
                # A # run is centered whatever the alignment (captured, layout rows 130 to 136).
                x = x0 + 1 + math.ceil((x1 - x0 - 1 - width) / 2)
                pieces = [(a, b, 1, x, "center")]
            else:
                pieces = [(a, b, 1, x0 + self.left_inset + indent, "left")]
            if origin.spills and n == 1 and align in ("left", "right", "center"):
                x = pieces[0][3]
                low, high = self.spill(origin, x, x + width, align)
            placed.extend(
                (*piece[:4], piece[4], baseline, j == 0, width if piece[:3] == (a, b, 1) else None)
                for j, piece in enumerate(pieces)
            )
        clip = (self.xs[low] + 1, top + 1, self.xs[high] + self.widths[high], bottom)
        last_line = baselines[-1] + shift if baselines else None
        runs = []
        for a, b, copies, x, anchor, baseline, first_piece, line_width in placed:
            if a == b:
                continue
            piece = text.text[a:b] * copies
            drawn = _spans(text, a, b, copies, shrink, shown.format_color) if spans else ()
            if drawn:
                width = sum(span.width for span in drawn)
                ruled = _span_rules(drawn, x, baseline, own)
                scale = shrink
            else:
                width = (
                    line_width if line_width is not None else xlformat.text_width(piece, font=font)
                )
                ruled = rules(cell_font, size, baseline, (x, x + width), own)
                scale = size / cell_font.size
            if (
                own is not None
                and (not first_piece or baseline != last_line)
                and any(accounting(f) for f in text.fonts(a, b))
            ):
                # An accounting underline under text spans the cell once, under its last line
                # (captured, text_style_more: three wrapped lines draw one rule).
                ruled = tuple(r for r in ruled if r.kind != "underline")
            runs.append(
                TextRun(
                    piece,
                    x,
                    baseline,
                    width,
                    clip,
                    anchor,  # type: ignore[arg-type]
                    source,
                    cell_font,
                    color,
                    scale,
                    ruled,
                    direction,
                    drawn,
                )
            )
        return runs, low, high, n

    @staticmethod
    def _between_slanted(
        x0: float,
        x1: float,
        rise: float,
        align: str,
        angle: int,
        above: float,
        below: float,
        indent: float,
    ) -> float:
        """The baseline's start x for angled text between its slanted sides (§4.17).

        The sides run along the baseline, so Excel lays the line out as in an upright cell turned
        with it: the side its ascent faces is that cell's top, the other its bottom. Text aligned
        to the ascent side puts its ascent ``SLANT_INSETS[0]`` inside it, text aligned to the
        other side its descent ``SLANT_INSETS[1]`` inside that one, centered text its line box
        midway; an indent moves the text horizontally. The baseline's height, ``rise`` from the
        cell's bottom edge (negative above it), is placed as without borders, and the x follows
        from its distance to the side.
        """
        sin, cos = math.sin(math.radians(angle)), math.cos(math.radians(angle))
        ascent_side = x0 if angle > 0 else x1
        across = abs(sin)
        apart = (x1 - x0) * across  # the sides' distance, across the line
        near = "left" if angle > 0 else "right"  # the alignment toward the ascent side
        if align == "center":
            depth = apart / 2 + (above - below) / 2
        elif align == near:
            depth = SLANT_INSETS[0] + above + indent * across
        else:
            depth = apart - SLANT_INSETS[1] - below - indent * across
        # depth: the baseline's distance below the ascent side, across the line; that side
        # passes through the cell's bottom corner on its own side.
        return ascent_side + (depth - rise * cos) / sin

    def _one_line(self, origin: _Origin) -> tuple[Text, tuple[TextSpan, ...] | None]:
        """The origin's text as one line: line breaks are not drawn, and wrapping is reported."""
        cell, shown = origin.cell, origin.shown
        style = cell.style
        wraps = cell.kind == "text" and (
            style.wrap or shown.align in _SPREAD or style.v_align in _SPREAD
        )
        if wraps and self.notes is not None:
            self.notes.add(
                "text.stacked_wrapped",
                "approximated",
                "not_implemented",
                row=cell.row,
                column=cell.column,
                detail="drawn on one line",
            )
        spans = drawn_spans(cell, shown.text)
        if spans:
            return Text.of_spans(spans, self.fonts.get, breaks=False), spans
        content = shown.text.replace("\r", "").replace("\n", "")
        return Text.plain(content, cell.font, self.fonts.get(cell.font)), None

    def place_turned(self, origin: _Origin) -> tuple[list[TextRun], int, int, int]:
        """Angled text: lines turned about their baselines' starts, placed (§4.17).

        A line's box, from its descent to its ascent and along its width, turned, is placed in the
        cell by its upright bounding box: general text by its angle's sign, left for 1 to 89 degrees
        and -90, right for -1 to -89 and 90 (captured at 15 to 89 degrees, their negatives, 90 and
        -90: the switch is exactly at ±90); bottom, top or center by the vertical alignment. A box
        taller than its row stays bottom-aligned and is clipped at the row (captured).

        Wrapped text breaks into lines as long as fit between the row's top and bottom, as a
        fitted row would hold them; the lines step a line apart toward the descent side, each
        placed in height on its own, and the whole is placed across the cell as one box
        (captured, rotation suite, turned).
        """
        cell, shown = origin.cell, origin.shown
        style = cell.style
        first, last = origin.cols[0], origin.cols[-1]
        x0, x1 = self.xs[first], self.xs[last] + self.widths[last]
        top = self.ys[origin.rows[0]]
        bottom = self.ys[origin.rows[-1]] + self.heights[origin.rows[-1]]
        text, spans, wraps = self._turned_text(origin)
        end = len(text.text)
        # The angled model's insets were fitted to the rotation suite's screen captures with
        # these offsets (§4.17), so it keeps them.
        pitch, above, below = print_baseline_offsets(
            line_font(text.fonts(0, end)) if spans else cell.font
        )
        angle = style.rotation
        cos, sin = math.cos(math.radians(angle)), math.sin(math.radians(angle))
        if wraps:
            size = max(font.size for font in text.fonts(0, end)) if spans else cell.font.size
            room = (bottom - top - ROTATED_EXTRA - size * abs(cos)) / abs(sin)
            lines = wrap(text.text, room, text.width)
        else:
            lines = [(0, end)]
        # Each line's start relative to the first's, before the block is placed across.
        placed: list[tuple[int, int, float, float, float]] = []  # a, b, dx, y, width
        box_lo = box_hi = None
        for k, (a, b) in enumerate(lines):
            width = text.width(a, b)
            corners = [
                (s * cos - t * sin, -s * sin - t * cos) for s in (0, width) for t in (-below, above)
            ]
            y_lo, y_hi = min(y for _, y in corners), max(y for _, y in corners)
            room = bottom - top - (y_hi - y_lo)
            if style.v_align == "top" and room >= 0:
                y = top - y_lo + min(TOP_DROP, room)
            elif style.v_align in ("center", "distributed") and y_hi - y_lo <= bottom - top:
                y = top + math.floor((bottom - top - (y_hi - y_lo)) / 2) - y_lo
            else:  # bottom; justify as bottom (unverified); a box taller than the row (captured)
                y = bottom - y_hi
            if k == 0:
                dx = 0.0
                first_y = y
            else:  # on the line k pitches toward the descent side, slid along it to height y
                along = (first_y + k * pitch * cos - y) / sin
                dx = k * pitch * sin + along * cos
            lo = dx + min(x for x, _ in corners)
            hi = dx + max(x for x, _ in corners)
            box_lo = lo if box_lo is None else min(box_lo, lo)
            box_hi = hi if box_hi is None else max(box_hi, hi)
            placed.append((a, b, dx, y, width))
        assert box_lo is not None
        assert box_hi is not None
        align = _turned_align(style.h_align, shown.align, angle)
        indent = style.indent * self.indent
        left_inset, right_inset = ROTATED_INSETS
        if align == "left":
            x = x0 + left_inset + indent - box_lo
        elif align == "right":
            x = x1 - right_inset - indent - box_hi
        else:
            x = (x0 + x1) / 2 - (box_lo + box_hi) / 2
        if len(origin.rows) == len(origin.cols) == 1 and slants(style):
            x = self._between_slanted(x0, x1, first_y - bottom, align, angle, above, below, indent)
        shift = 0.0 if spans else vert_shift(cell.font)
        low, high = first, last
        # Fill does not spill at an angle: its line stops at the cell's edge (captured,
        # rotation_more: 15 and -45 degrees, about 2 pt past the edge in Excel).
        if origin.spills and style.h_align != "fill":
            low, high = self.spill(origin, x + box_lo, x + box_hi, align)
        clip = (self.xs[low] + 1, top + 1, self.xs[high] + self.widths[high], bottom)
        size = cell.font.size * vert_scale(cell.font)
        direction: TextDirection = (
            "rtl" if style.reading_order != 1 and right_to_left(text.text) else "ltr"
        )
        runs = []
        for a, b, dx, y, width in placed:
            # down in the turned frame for superscript and subscript
            rx, ry = x + dx + shift * sin, y + shift * cos
            # An accounting underline runs under the text's width, turned with it (unverified).
            own = (rx, rx + width) if cell.kind == "text" else None
            if spans:
                drawn = _spans(text, a, b, 1, 1.0, shown.format_color)
                ruled = _span_rules(drawn, rx, ry, own)
                scale = 1.0
            else:
                drawn = ()
                ruled = rules(cell.font, size, ry, (rx, rx + width), own)
                scale = size / cell.font.size
            runs.append(
                TextRun(
                    text.text[a:b],
                    rx,
                    ry,
                    width,
                    clip,
                    "left",
                    (cell.row, cell.column),
                    cell.font,
                    shown.format_color or cell.font.color,
                    scale,
                    ruled,
                    direction,
                    drawn,
                    float(angle),
                )
            )
        return runs, low, high, len(lines)

    def _turned_text(self, origin: _Origin) -> tuple[Text, tuple[TextSpan, ...] | None, bool]:
        """Angled text, with its line feeds when it wraps, and whether it wraps."""
        cell, shown = origin.cell, origin.shown
        style = cell.style
        wraps = cell.kind == "text" and (
            style.wrap or shown.align in _SPREAD or style.v_align in _SPREAD
        )
        spans = drawn_spans(cell, shown.text)
        if spans:
            return Text.of_spans(spans, self.fonts.get, breaks=wraps), spans, wraps
        content = shown.text.replace("\r", "")
        if not wraps:
            content = content.replace("\n", "")
        return Text.plain(content, cell.font, self.fonts.get(cell.font)), None, wraps

    def place_stacked(self, origin: _Origin) -> tuple[list[TextRun], int, int, int]:
        """Stacked text: one upright character a line, a font line apart (§4.17).

        Each character is centered on one vertical axis: ``STACKED_AXIS`` from the near edge for
        left and right, the cell's middle for center and general (captured, ``rotation_placement``
        at 9 to 36 pt; Calibri in a 120 pt column only). The
        lines are placed as wrapped lines are by the vertical alignment, clipped at the cell.
        """
        cell, shown = origin.cell, origin.shown
        style = cell.style
        first, last = origin.cols[0], origin.cols[-1]
        x0, x1 = self.xs[first], self.xs[last] + self.widths[last]
        top = self.ys[origin.rows[0]]
        bottom = self.ys[origin.rows[-1]] + self.heights[origin.rows[-1]]
        text, spans = self._one_line(origin)
        n = len(text.text)
        metrics = [
            line_metrics(text.fonts(k, k + 1) if spans else [cell.font], 1) for k in range(n)
        ]
        # stacked text keeps its own bottom, at least the Normal font's (the row's shared
        # bottom is not captured for it)
        own = metrics[-1][2] if metrics else 0.0
        normal = self.fonts.bottom(self.normal)[0] if self.normal is not None else 0.0
        baselines = _baselines(metrics, top, bottom, style.v_align, max(own, normal))
        size = cell.font.size * vert_scale(cell.font)
        shift = 0.0 if spans else vert_shift(cell.font)
        align = shown.align if style.h_align in ("left", "right") else "center"
        clip = (x0 + 1, top + 1, x1, bottom)
        runs = []
        for k, baseline in enumerate(baselines):
            char = text.text[k]
            if not char.strip():
                continue
            width = text.width(k, k + 1)
            if align == "left":
                x = x0 + _stacked_axis(cell.font.size) - width / 2
            elif align == "right":
                x = x1 - _stacked_axis(cell.font.size) - width / 2
            else:
                x = (x0 + x1 - width) / 2
            baseline += shift
            drawn = _spans(text, k, k + 1, 1, 1.0, shown.format_color) if spans else ()
            runs.append(
                TextRun(
                    char,
                    x,
                    baseline,
                    width,
                    clip,
                    align,  # type: ignore[arg-type]
                    (cell.row, cell.column),
                    cell.font,
                    shown.format_color or cell.font.color,
                    1.0 if spans else size / cell.font.size,
                    _span_rules(drawn, x, baseline, None)
                    if drawn
                    else rules(cell.font, size, baseline, (x, x + width), None),
                    "ltr",
                    drawn,
                )
            )
        return runs, first, last, n

    def spread(
        self,
        words: list[tuple[int, int]],
        text: Text,
        x0: float,
        x1: float,
        distributed: bool,
        indent: float = 0.0,
        trailing: bool = False,
    ) -> list[tuple[int, int, int, float, str]]:
        """Words spread across the cell between its insets, less ``indent`` at both ends.

        Each word is its start and end in ``text``; each piece returned is a word's start, end,
        copies, x and anchor. The room left is shared by the gaps between words and, when
        ``trailing`` (the line was broken at a space), one more after the last word, as Excel
        shares it (captured, ``text_style`` group ``lone``). A line without a gap, a single word
        ending its paragraph, is centered when ``distributed`` and left-aligned otherwise.
        """
        lo, hi = x0 + self.left_inset + indent, x1 - self.right_inset - indent
        widths = [text.width(a, b) for a, b in words]
        gaps = len(words) - 1 + trailing
        if not gaps:
            (a, b), width = words[0], widths[0]
            if distributed:
                return [(a, b, 1, lo + math.floor((hi - lo - width) / 2), "center")]
            return [(a, b, 1, lo, "left")]
        gap = (hi - lo - sum(widths)) / gaps
        result, x = [], lo
        for (a, b), width in zip(words, widths, strict=True):
            result.append((a, b, 1, x, "left"))
            x += width + gap
        return result

    def spill(self, origin: _Origin, x: float, end: float, align: str) -> tuple[int, int]:
        """The visible column positions a single line from ``x`` to ``end`` spans (§4.1).

        Left text takes a column to the right once its end passes that column's gridline. Right
        text takes one to the left once its start reaches the far edge of its own leftmost
        gridline: the one-point asymmetry Excel shows (captured). Centered text takes a column
        on both sides at once while it overflows either edge; once one side is blocked, the
        other grows only as far as its own text needs (captured). A cell holding a value, or
        merged, blocks.
        """
        i = origin.rows[0]
        first, last = origin.cols[0], origin.cols[-1]
        left = right = 0
        left_open = right_open = True
        while True:
            lo = self.xs[first - left]
            hi = self.xs[last + right] + self.widths[last + right]
            need_left = x <= lo + 1 if align == "right" else x < lo + 1
            need_right = end > hi
            if not (need_left or need_right):
                break
            left_open = left_open and self.grid.free(i, first - left - 1, self.no_claims)
            right_open = right_open and self.grid.free(i, last + right + 1, self.no_claims)
            centered = align == "center" and left_open and right_open
            grow_left = left_open and (need_left or centered)
            grow_right = right_open and (need_right or centered)
            if not (grow_left or grow_right):
                break
            left += grow_left
            right += grow_right
        return first - left, last + right


def _turned_align(h_align: str, shown: str, angle: int) -> str:
    """Where angled text's box goes across the cell: left, right or center (§4.17).

    General follows the angle and switches sides at exactly 90 and -90 degrees. Fill, justify and
    distributed do not switch: placed as left at positive angles and as right at negative ones,
    90 and -90 included, but distributed at 90 and -90 is centered (captured, rotation,
    rotation_more and rotation_placement from -90 to 90 degrees).
    """
    if h_align == "general":
        return "left" if 0 < angle < 90 or angle == -90 else "right"
    if h_align in ("center", "center_across"):
        return "center"
    if h_align in ("fill", "justify", "distributed"):
        if h_align == "distributed" and abs(angle) == 90:
            return "center"
        return "left" if angle > 0 else "right"
    return "right" if shown == "right" else "left"


def _running(sizes: list[float]) -> list[float]:
    positions, total = [], 0.0
    for size in sizes:
        positions.append(total)
        total += size
    return positions


def _shifted(run: TextRun, dx: float) -> TextRun:
    if not dx:
        return run
    left, top, right, bottom = run.clip
    return replace(
        run,
        x=run.x - dx,
        clip=(left - dx, top, right - dx, bottom),
        rules=tuple(replace(r, x0=r.x0 - dx, x1=r.x1 - dx) for r in run.rules),
    )


def _wrap(text: str, budget: float, font: xlformat.Font) -> list[str]:
    """Lines of ``text`` broken at spaces to fit ``budget`` points; explicit breaks kept.

    A word wider than the budget is broken between characters.
    """
    text = text.replace("\r", "")
    return [
        text[a:b]
        for a, b in wrap(text, budget, lambda a, b: xlformat.text_width(text[a:b], font=font))
    ]


def _baselines(
    metrics: list[tuple[float, float, float]],
    top: float,
    bottom: float,
    v_align: str,
    row_below: float | None = None,
) -> list[float]:
    """The baseline of each line of a block from ``top`` to ``bottom``, placed by ``v_align``.

    ``metrics`` holds each line's pitch, its baseline below its top and its own baseline above
    its bottom (``line_metrics``); ``row_below`` is where the row puts a bottom-aligned last
    baseline above the bottom (``_Points.row_below``), the last line's own when not given.

    - Top: the first baseline ``ceil(S) + 1`` below the top (captured, ``text_style_more``
      group ``faces``).
    - Bottom: the last baseline ``row_below`` above the bottom, shared by the row (captured,
      ``faces``, ``shrunk`` rows 10 to 16, ``rules`` rows 18 to 22).
    - Center: the first baseline ``ceil(S) + 1 + trunc((H - k·F) / 2)`` below the top, the
      line's own font and ``k`` lines of pitch ``F``, the half truncated toward zero (captured,
      ``faces`` and ``overflow``); the row's other fonts do not move it.
    - Vertical distributed and justify: lines go evenly from the top position to the bottom
      one, each a whole number of points below the first, a half point dropped (captured, suite
      ``text_style``, group ``align``; the bottom's ``row_below`` beside larger fonts
      unverified).

    When the lines do not fit, bottom and center text places the first ``k = min(n, ceil(H /
    F))`` lines that reach the bottom as a block of its own, and the rest follow, clipped
    (captured, ``text_style_more`` groups ``shrunk`` and ``overflow``: two and three lines in
    rows of 15 to 32 pt, and a line of 14 or 20 pt in a 15 pt row); top text and the spread
    alignments start at the top (captured, ``overflow``).
    """
    heights = [m[0] for m in metrics]
    total = sum(heights)
    n = len(metrics)
    block = metrics
    if row_below is None:
        row_below = metrics[-1][2]
    if total > bottom - top and v_align in ("bottom", "center"):
        k = 1
        while k < n and sum(heights[:k]) < bottom - top:
            k += 1
        block = metrics[:k]
    elif total > bottom - top:
        v_align = "top"
    first_lines = sum(m[0] for m in block[:-1])
    if v_align not in ("bottom", "center", "distributed", "justify"):
        line_top = top
    elif v_align in _SPREAD and n > 1:
        first = top + metrics[0][1]
        pitch = (bottom - row_below - first) / (n - 1)
        return [first + math.floor(k * pitch) for k in range(n)]
    elif v_align in ("center", "distributed"):
        own = bottom - block[-1][2] - block[-1][1] - first_lines
        line_top = top + math.trunc((own - top) / 2)
    elif v_align == "bottom":
        line_top = bottom - row_below - block[-1][1] - first_lines
    else:  # justify, one line: at the top
        line_top = top
    result = []
    for height, above, _ in metrics:
        result.append(line_top + above)
        line_top += height
    return result


def _spans(
    text: Text, start: int, end: int, copies: int, shrink: float, section: str | None
) -> tuple[TextSpan, ...]:
    """The rich-text spans of a piece of text as drawn: width, size, shift (§4.16).

    The number format's section color replaces each run's color, as it replaces a cell's.
    """
    result = []
    for _ in range(copies):
        for piece, part in text.pieces(start, end):
            font = replace(part.font, color=section) if section else part.font
            result.append(
                TextSpan(
                    piece,
                    font,
                    xlformat.text_width(piece, font=part.measured),
                    vert_scale(part.font) * shrink,
                    vert_shift(part.font),
                )
            )
    return tuple(result)


def _span_rules(
    spans: tuple[TextSpan, ...], x: float, baseline: float, cell: tuple[float, float] | None
) -> tuple[TextRule, ...]:
    """The underlines and strikethrough of each span, each drawn once."""
    found: dict[TextRule, None] = {}
    for span in spans:
        for rule in rules(
            span.font, span.font.size * span.scale, baseline + span.shift, (x, x + span.width), cell
        ):
            found[rule] = None
        x += span.width
    return tuple(found)
