"""``excel`` mode: a sheet laid out in points, as Excel for Mac lays it out.

Columns take xlformat's widths in points and rows the heights of ``ARCHITECTURE.md`` §4.7. Each
cell is formatted at its column's width in dialog units (a merge at its merged width), then
placed by its alignment and spilled into free neighbors by the rules of §4.1, in points. The
result is a list of :class:`~sheetpeek.TextRun` per row, with each run's clip window, and the
column boundaries text crosses. Two spills may take the same free cell; both are kept, as Excel
draws both.
"""

from __future__ import annotations

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
from sheetpeek._model import Column, Row, Sheet
from sheetpeek._rendered import RenderedCell, RenderedColumn, RenderedRow, TextRun
from sheetpeek._rows import (
    fits_cells,
    font_row_height,
    row_height,
    wrapped_height,
)
from sheetpeek._runs import Text, drawn_spans, fitted_height, line_metrics, runs_shrink_step, wrap
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
        sheet, rows, every, [0] * len(every), added, 0, False, _Formatter(sheet, notes, fonts)
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
    heights = [
        _fitted(grid, i, row, sheet, read, widths, base, fonts) for i, row in enumerate(rows)
    ]
    layout = _Points(grid, widths, heights, base, fonts)
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
) -> float:
    """A row's height, fitted to its cells as Excel fits it on opening (§4.7)."""
    height = row_height(row, sheet)
    if not fits_cells(row, sheet):
        return height
    styled = []
    for p in read:
        cell = grid.cell(i, p)
        # A cell with rich-text runs counts its runs' fonts (§4.16).
        spans = drawn_spans(cell, _UNDRAWN.sub("", str(cell.value))) if cell.runs else None
        own = [span.font for span in spans] if spans else [cell.font]
        height = max(height, *map(font_row_height, own))
        if cell.kind != "empty":
            styled += [f for f in own if f.vert_align or accounting(f)]
        # A merged cell's font counts, its wrapped lines do not (captured).
        merged = (row.index, grid.columns[p].index) in grid.merge_of
        if cell.kind == "text" and cell.style.wrap and widths[p] and not merged:
            budget = widths[p] - base.left_inset - base.right_inset
            if spans:
                text = Text.of_spans(spans, fonts.get, breaks=True)
                height = max(height, fitted_height(text, wrap(text.text, budget, text.width)))
            else:
                lines = _wrap(str(cell.value), budget, fonts.get(cell.font))
                height = max(height, wrapped_height(len(lines), cell.font))
    # Superscript, subscript and accounting underlines grow the row (§4.13); the row fits the
    # unshrunk size of shrunk text, so shrink adds nothing.
    for font, extra in row_extras(styled):
        height = max(height, font_row_height(font) + extra)
    return height


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
    ) -> None:
        self.grid = grid
        self.widths = widths
        self.heights = heights
        self.left_inset = base.left_inset
        self.right_inset = base.right_inset
        self.digit = base.digit_width
        self.fonts = fonts
        self.xs = _running(widths)
        self.ys = _running(heights)
        self.no_claims: set[tuple[int, int]] = set()  # both competing spills take a cell

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
                        cell=cell,
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
        x0, x1 = self.xs[first], self.xs[last] + self.widths[last]
        top = self.ys[origin.rows[0]]
        bottom = self.ys[origin.rows[-1]] + self.heights[origin.rows[-1]]
        style = cell.style
        font = self.fonts.get(cell.font)  # at the drawn size of superscript and subscript
        size = cell.font.size * vert_scale(cell.font)
        align = shown.align
        budget = (x1 - x0) - self.left_inset - self.right_inset
        v_align = style.v_align
        # Distributed and justified text wraps, horizontally or vertically (§3.6, §4.13).
        wraps = cell.kind == "text" and (
            style.wrap or align in ("justify", "distributed") or v_align in _SPREAD
        )
        indent = style.indent * self.digit
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
            text = Text.plain(content if wraps else content.replace("\n", ""), cell.font, font)
        lines = wrap(text.text, budget, text.width) if wraps else [(0, len(text.text))]
        n = len(lines)
        if spans:
            metrics = [line_metrics(text.fonts(a, b), n) for a, b in lines]
        else:
            metrics = [line_metrics([cell.font], n)] * n
        baselines = _baselines(metrics, top, bottom, v_align)
        shift = 0.0 if spans else vert_shift(cell.font)
        color = shown.format_color or cell.font.color
        direction: TextDirection = (
            "rtl" if style.reading_order != 1 and right_to_left(text.text) else "ltr"
        )
        source = (cell.row, cell.column)
        low, high = first, last
        own = (x0, x1) if cell.kind == "text" else None  # an accounting underline's span
        # Each piece: start, end and copies of the text, x, anchor, baseline, whether first.
        placed: list[tuple[int, int, int, float, str, float, bool]] = []
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
                pieces = self.spread(
                    words, text, x0, x1, distributed, indent if distributed else 0.0
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
                (*piece[:4], piece[4], baseline, j == 0) for j, piece in enumerate(pieces)
            )
        clip = (self.xs[low] + 1, top + 1, self.xs[high] + self.widths[high], bottom)
        runs = []
        for a, b, copies, x, anchor, baseline, first_piece in placed:
            if a == b:
                continue
            piece = text.text[a:b] * copies
            drawn = _spans(text, a, b, copies, shrink, shown.format_color) if spans else ()
            if drawn:
                width = sum(span.width for span in drawn)
                ruled = _span_rules(drawn, x, baseline, own)
                scale = shrink
            else:
                width = xlformat.text_width(piece, font=font)
                ruled = rules(cell.font, size, baseline, (x, x + width), own)
                scale = size / cell.font.size
            if own is not None and not first_piece and any(accounting(f) for f in text.fonts(a, b)):
                # An accounting underline under text spans the cell, once per line.
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
                    cell.font,
                    color,
                    scale,
                    ruled,
                    direction,
                    drawn,
                )
            )
        return runs, low, high, n

    def spread(
        self,
        words: list[tuple[int, int]],
        text: Text,
        x0: float,
        x1: float,
        distributed: bool,
        indent: float = 0.0,
    ) -> list[tuple[int, int, int, float, str]]:
        """Words spread across the cell between its insets, less ``indent`` at both ends.

        Each word is its start and end in ``text``; each piece returned is a word's start, end,
        copies, x and anchor. A single word is centered there when ``distributed`` and
        left-aligned otherwise.
        """
        lo, hi = x0 + self.left_inset + indent, x1 - self.right_inset - indent
        widths = [text.width(a, b) for a, b in words]
        if len(words) == 1:
            (a, b), width = words[0], widths[0]
            if distributed:
                return [(a, b, 1, lo + math.floor((hi - lo - width) / 2), "center")]
            return [(a, b, 1, lo, "left")]
        gap = (hi - lo - sum(widths)) / (len(words) - 1)
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
    metrics: list[tuple[float, float, float]], top: float, bottom: float, v_align: str
) -> list[float]:
    """The baseline of each line of a block from ``top`` to ``bottom``, placed by ``v_align``.

    ``metrics`` holds each line's height, its baseline below its top and its baseline above its
    bottom (``line_metrics``). Lines that do not fit start at the top: a custom-height row shows
    its first line (captured, layout rows 161 to 168; unverified for centered text). Spread
    lines go evenly from the top position to the bottom one (§4.13).
    """
    heights = [m[0] for m in metrics]
    total = sum(heights)
    n = len(metrics)
    if total > bottom - top or v_align not in ("bottom", "center", "distributed", "justify"):
        line_top = top
    elif v_align in _SPREAD and n > 1:
        first = top + metrics[0][1]
        pitch = (bottom - metrics[-1][2] - first) / (n - 1)
        return [first + k * pitch for k in range(n)]
    elif v_align in ("center", "distributed"):
        line_top = top + math.floor((bottom - top - total) / 2)
    elif v_align == "bottom":
        line_top = bottom - metrics[-1][2] - metrics[-1][1] - sum(heights[:-1])
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
