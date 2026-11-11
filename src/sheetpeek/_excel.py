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
from typing import Any

import xlformat

from sheetpeek._layout import (
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
    WRAPPED_LEADING,
    baseline_offsets,
    fits_cells,
    font_row_height,
    row_height,
    wrapped_height,
)


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
    for p in read:
        cell = grid.cell(i, p)
        height = max(height, font_row_height(cell.font))
        # A merged cell's font counts, its wrapped lines do not (captured).
        merged = (row.index, grid.columns[p].index) in grid.merge_of
        if cell.kind == "text" and cell.style.wrap and widths[p] and not merged:
            budget = widths[p] - base.left_inset - base.right_inset
            lines = _wrap(str(cell.value), budget, fonts.get(cell.font))
            height = max(height, wrapped_height(len(lines), cell.font))
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
        font = self.fonts.get(cell.font)
        align = shown.align
        budget = (x1 - x0) - self.left_inset - self.right_inset
        wraps = cell.kind == "text" and (cell.style.wrap or align in ("justify", "distributed"))
        text = shown.text
        # a line break in a cell without wrapping is not drawn
        lines = _wrap(text, budget, font) if wraps else [text.replace("\r", "").replace("\n", "")]
        line, above, below = baseline_offsets(cell.font)
        n = len(lines)
        if n > 1:
            line += WRAPPED_LEADING  # provisional: the pitch of a fitted wrapped row
        v_align = cell.style.v_align
        if n * line > bottom - top:
            # Lines that do not fit start at the top: a custom-height row shows its first line
            # (captured, layout rows 161 to 168; unverified for centered text).
            first_baseline = top + above
        elif v_align == "center":
            first_baseline = top + math.floor((bottom - top - n * line) / 2) + above
        elif v_align == "bottom":
            first_baseline = bottom - below - (n - 1) * line
        else:
            first_baseline = top + above
        indent = cell.style.indent * self.digit
        color = shown.format_color or cell.font.color
        source = (cell.row, cell.column)
        low, high = first, last
        placed: list[tuple[str, float, str, float]] = []  # (text, x, anchor, baseline)
        for k, line_text in enumerate(lines):
            baseline = first_baseline + k * line
            width = xlformat.text_width(line_text, font=font)
            words = line_text.split()
            if align == "fill":
                copies = max(1, math.floor(budget / width)) if width else 1
                pieces = [(line_text * copies, x0 + self.left_inset, "left")]
            elif align == "distributed" or (align == "justify" and k < n - 1 and len(words) > 1):
                pieces = self.spread(words, x0, x1, font, align == "distributed")
            elif align == "right" and not shown.overflow:
                pieces = [(line_text, x1 - self.right_inset - indent - width, "right")]
            elif align == "center" or shown.overflow:
                # A # run is centered whatever the alignment (captured, layout rows 130 to 136).
                x = x0 + 1 + math.ceil((x1 - x0 - 1 - width) / 2)
                pieces = [(line_text, x, "center")]
            else:
                pieces = [(line_text, x0 + self.left_inset + indent, "left")]
            if origin.spills and n == 1 and align in ("left", "right", "center"):
                x = pieces[0][1]
                low, high = self.spill(origin, x, x + width, align)
            placed.extend((piece, x, anchor, baseline) for piece, x, anchor in pieces)
        clip = (self.xs[low] + 1, top + 1, self.xs[high] + self.widths[high], bottom)
        runs = [
            TextRun(
                piece,
                x,
                baseline,
                xlformat.text_width(piece, font=font),
                clip,
                anchor,  # type: ignore[arg-type]
                source,
                cell.font,
                color,
            )
            for piece, x, anchor, baseline in placed
            if piece
        ]
        return runs, low, high, n

    def spread(
        self, words: list[str], x0: float, x1: float, font: xlformat.Font, distributed: bool
    ) -> list[tuple[str, float, str]]:
        """Words spread across the cell between its insets.

        A single word is centered there when ``distributed`` and left-aligned otherwise.
        """
        lo, hi = x0 + self.left_inset, x1 - self.right_inset
        widths = [xlformat.text_width(w, font=font) for w in words]
        if len(words) == 1:
            if distributed:
                return [(words[0], lo + math.floor((hi - lo - widths[0]) / 2), "center")]
            return [(words[0], lo, "left")]
        gap = (hi - lo - sum(widths)) / (len(words) - 1)
        result, x = [], lo
        for word, width in zip(words, widths, strict=True):
            result.append((word, x, "left"))
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
    return TextRun(
        run.text, run.x - dx, run.y, run.width, (left - dx, top, right - dx, bottom),
        run.align, run.source, run.font, run.color,
    )  # fmt: skip


def _wrap(text: str, budget: float, font: xlformat.Font) -> list[str]:
    """Lines of ``text`` broken at spaces to fit ``budget`` points; explicit breaks kept.

    A word wider than the budget is broken between characters.
    """
    lines: list[str] = []
    for paragraph in text.replace("\r", "").split("\n"):
        line = ""
        for word in paragraph.split(" "):
            candidate = f"{line} {word}" if line else word
            if line and xlformat.text_width(candidate, font=font) > budget:
                lines.append(line)
                candidate = word
            line = candidate
            while len(line) > 1 and xlformat.text_width(line, font=font) > budget:
                cut = len(line) - 1
                while cut > 1 and xlformat.text_width(line[:cut], font=font) > budget:
                    cut -= 1
                lines.append(line[:cut])
                line = line[cut:]
        lines.append(line)
    return lines
