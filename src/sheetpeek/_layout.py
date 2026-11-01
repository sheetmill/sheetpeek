"""Lay out a :class:`~sheetpeek.Sheet` for one width mode: what each visible cell shows.

Each cell is formatted through xlformat at its width. Its text is then placed on a character
canvas per row, aligned, spilled into free neighbors, and clipped. The canvas is cut back into
columns and gaps, so that every column records the characters drawn in it and whose they are.
The rules and the evidence for them are in ``ARCHITECTURE.md`` §4; where characters mode must
differ from Excel, the comments say so.
"""

from __future__ import annotations

import textwrap
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

import xlformat

from sheetpeek._model import Cell, CellRange, CellStyle, Column, Font, Row, Sheet
from sheetpeek._rendered import (
    Alignment,
    Gap,
    Mode,
    RenderedCell,
    RenderedColumn,
    RenderedRow,
    RenderedSheet,
    SpillEdge,
)
from sheetpeek._rows import row_height

# Characters per indent level. Captured (ARCHITECTURE.md §4.6): one level is one digit's width,
# 6 pt at Calibri 11.
_INDENT = 1
_EXPLICIT: dict[str, Alignment] = {
    "left": "left",
    "center": "center",
    "right": "right",
    "fill": "fill",
    "justify": "justify",
    "distributed": "distributed",
    "center_across": "center",
}
_NOBODY = -1
_MAX_COLUMN = 16_384


def layout(
    sheet: Sheet,
    mode: Mode = "characters",
    *,
    show_hidden: bool = False,
    scale: float = 1.0,
    max_width: int | None = None,
    spill_edge: SpillEdge = "extend",
) -> RenderedSheet:
    """Lay out a sheet; see :meth:`sheetpeek.Sheet.layout`."""
    if mode not in ("excel", "characters", "none"):
        raise ValueError(f"unknown width mode {mode!r}; expected 'excel', 'characters' or 'none'")
    if scale <= 0:
        raise ValueError("scale must be positive")
    if spill_edge != "extend":
        raise ValueError(f"unknown spill_edge {spill_edge!r}; only 'extend' is available")
    warnings = list(sheet.warnings)
    fonts = _Fonts(warnings)
    columns = [c for c in sheet.columns if show_hidden or not _column_hidden(c)]
    rows = [r for r in sheet.rows if show_hidden or not _row_hidden(r, sheet)]
    if mode == "excel":
        if scale != 1.0 or max_width is not None:
            raise ValueError("scale and max_width apply to characters mode only")
        from sheetpeek._excel import lay_out

        return lay_out(sheet, columns, rows, warnings, fonts)
    characters = mode == "characters"
    gap = 1 if characters else 0

    normal_font = fonts.get(sheet.normal_font) if characters else None

    def width_of(column: Column) -> int | None:
        if normal_font is None:
            return None
        return _characters(column, sheet, normal_font, scale, max_width)

    if sheet.ref is None or not columns or not rows:
        rendered = tuple(RenderedColumn(c.index, width_of(c), _column_hidden(c)) for c in columns)
        return RenderedSheet(sheet, mode, rendered, (), gap, tuple(warnings))
    before: list[Column] = []
    after: list[Column] = []
    if characters:
        before, after = _edge_columns(sheet, width_of(Column(1)) or 1)
    every = before + columns + after
    widths = [width_of(c) for c in every]
    added = set(range(len(before))) | set(range(len(before) + len(columns), len(every)))
    grid = _Grid(
        sheet, rows, every, widths, added, gap, characters, _Formatter(sheet, warnings, fonts)
    )
    rendered_rows, kept = grid.render()
    rendered_columns = tuple(
        RenderedColumn(every[p].index, widths[p], _column_hidden(every[p]), p in added)
        for p in kept
    )
    return RenderedSheet(sheet, mode, rendered_columns, rendered_rows, gap, tuple(warnings))


def _edge_columns(sheet: Sheet, default: int) -> tuple[list[Column], list[Column]]:
    """Empty columns past the range that spilled text may need, at the default width.

    Only past the edge of the used range, where the sheet is known to be empty: a range that
    cuts through the data was not read beyond its edge, so spill stops there.
    """
    ref, used = sheet.ref, sheet.used_range
    if ref is None or used is None:
        return [], []
    longest = max((len(str(c.value)) for row in sheet.cells for c in row if c.kind == "text"),
                  default=0)  # fmt: skip
    if not longest:
        return [], []
    count = longest // default + 1
    before = after = range(0)
    if ref.min_col <= used.min_col:
        before = range(max(1, ref.min_col - count), ref.min_col)
    if ref.max_col >= used.max_col:
        after = range(ref.max_col + 1, min(_MAX_COLUMN, ref.max_col + count) + 1)
    return [Column(i) for i in before], [Column(i) for i in after]


def _column_hidden(column: Column) -> bool:
    return column.hidden or column.width == 0


def _row_hidden(row: Row, sheet: Sheet) -> bool:
    """Hidden, or 0 pt high by the sheet's format (``zeroHeight``)."""
    return row.hidden or row_height(row, sheet) == 0


def _characters(
    column: Column, sheet: Sheet, normal_font: xlformat.Font, scale: float, max_width: int | None
) -> int:
    """A column's width in characters: Excel's Column Width dialog number, rounded.

    The dialog number is in digits of the Normal font, and Calibri's digits are tabular, so a
    column of width ``w`` holds ``w`` digits in Excel and in a terminal (ARCHITECTURE.md §2.2).
    """
    metrics = xlformat.file_column_metrics(
        column.width,
        base_col_width=sheet.base_column_width,
        default_col_width=sheet.default_column_width,
        font=normal_font,  # the Normal font: xlformat names it font= here, unlike format_cell
    )
    chars = max(1, round(metrics.width * scale))
    return min(chars, max_width) if max_width else chars


@dataclass(slots=True)
class _Shown:
    """What a cell shows at a width, before placement."""

    text: str = ""
    full_text: str = ""
    overflow: bool = False
    align: Alignment = "left"
    format_color: str | None = None


class _Fonts:
    """xlformat's font for each sheetpeek font, built once.

    A family xlformat has no glyph table for is measured as Calibri at the same size and style
    (decided, ARCHITECTURE.md §6.2), with one warning per family, given only where the font is
    measured.
    """

    def __init__(self, warnings: list[str]) -> None:
        self._warnings = warnings
        self._measured = {name.lower() for name in xlformat.measured_fonts()}
        self._fonts: dict[Font, xlformat.Font] = {}
        self._warned: set[str] = set()

    def get(self, font: Font, *, measured: bool = True) -> xlformat.Font:
        """The font to pass to xlformat; ``measured`` when its glyph widths decide the layout."""
        known = font.name.lower() in self._measured
        if measured and not known and font.name not in self._warned:
            self._warned.add(font.name)
            self._warnings.append(
                f"font {font.name!r} has no glyph table in xlformat; measured as Calibri "
                "at the same size"
            )
        found = self._fonts.get(font)
        if found is None:
            name = font.name if known else "Calibri"
            found = self._fonts[font] = xlformat.Font(name, font.size, font.bold, font.italic)
        return found


class _Formatter:
    """Formats cells through xlformat; warns once per refused format and per failing format."""

    def __init__(self, sheet: Sheet, warnings: list[str], fonts: _Fonts) -> None:
        self._date1904 = sheet.date1904
        self._warnings = warnings
        self._refused: set[str] = set()
        self._unformattable: set[str] = set()
        self._fonts = fonts
        self._sheet = sheet

    def show(
        self, cell: Cell, width: int | None = None, column_width: float | None = None
    ) -> _Shown:
        """What a cell shows at ``width`` characters, ``column_width`` dialog units, or neither.

        Only ``column_width`` (excel mode) measures text, so only then do fonts xlformat does not
        measure give a warning.
        """
        explicit = _EXPLICIT.get(cell.style.h_align)
        if cell.kind == "empty":
            return _Shown(align=explicit or "left")
        fmt = cell.number_format or "General"
        measured = column_width is not None
        font = self._fonts.get(cell.font, measured=measured)
        normal = self._fonts.get(self._sheet.normal_font, measured=measured)
        try:
            try:
                formatted = self._format(cell.value, fmt, width, column_width, font, normal)
            except xlformat.FormatError:
                if fmt not in self._refused:
                    self._refused.add(fmt)
                    self._warnings.append(
                        f"number format {fmt!r} is one Excel refuses to open; shown as General"
                    )
                fmt = "General"
                formatted = self._format(cell.value, fmt, width, column_width, font, normal)
        except (ValueError, TypeError) as error:
            # A value xlformat cannot format (an infinite number from a malformed file).
            if fmt not in self._unformattable:
                self._unformattable.add(fmt)
                self._warnings.append(
                    f"number format {fmt!r} cannot show the value {cell.value!r} "
                    f"(first: {cell.coordinate}; {error}); shown as its text"
                )
            text = str(cell.value)
            return _Shown(text, text, False, explicit or "left")
        align: Alignment = explicit or formatted.align
        color = xlformat.DEFAULT_PALETTE.get(formatted.color) if formatted.color else None
        return _Shown(formatted.text, formatted.full_text, formatted.overflow, align, color)

    def _format(
        self,
        value: object,
        fmt: str,
        width: int | None,
        column_width: float | None,
        font: xlformat.Font,
        normal: xlformat.Font,
    ) -> xlformat.FormattedCell:
        return xlformat.format_cell(
            value,
            fmt,
            width=width,
            column_width=column_width,
            font=font,
            normal_font=normal,
            date1904=self._date1904,
        )


@dataclass(slots=True)
class _Origin:
    """A cell whose text is drawn: a plain cell, a merge, or a center-across-selection run."""

    cell: Cell
    shown: _Shown
    rows: list[int]  # visible row positions it occupies
    cols: list[int]  # visible column positions it occupies, contiguous
    spills: bool  # may extend into free neighbors
    lines: list[str] = field(default_factory=list)
    spill: tuple[int, int] = (0, 0)


class _Grid:
    def __init__(
        self,
        sheet: Sheet,
        rows: list[Any],
        columns: list[Column],
        widths: list[int | None],
        added: set[int],
        gap: int,
        characters: bool,
        formatter: _Formatter,
    ) -> None:
        self.sheet = sheet
        self.rows = rows
        self.columns = columns
        self.added = added  # positions of the empty columns added past the range
        self.empty_style = CellStyle("General", sheet.normal_font)
        self.characters = characters
        self.gap = gap
        self.formatter = formatter
        self.widths = [w or 0 for w in widths]
        self.starts: list[int] = []
        x = 0
        for width in self.widths:
            self.starts.append(x)
            x += width + gap
        self.total = max(x - gap, 0)
        self.col_pos = {c.index: p for p, c in enumerate(columns)}
        self.row_pos = {r.index: i for i, r in enumerate(rows)}
        # excel mode: the merged or single width in dialog units of visible column positions,
        # or None for a block without width (a width-0 column shown with show_hidden)
        self.units: Callable[[list[int]], float | None] | None = None
        self.merge_of: dict[tuple[int, int], CellRange] = {}
        ref = sheet.ref
        for merge in sheet.merges:
            for row in range(merge.min_row, merge.max_row + 1):
                for col in range(merge.min_col, merge.max_col + 1):
                    if ref is not None and (row, col) in ref:
                        self.merge_of[(row, col)] = merge

    def cell(self, i: int, p: int) -> Cell:
        if p in self.added:
            return Cell(self.rows[i].index, self.columns[p].index, None, "empty", self.empty_style)
        return self.sheet.cell(self.rows[i].index, self.columns[p].index)

    def block_width(self, cols: list[int]) -> int:
        return sum(self.widths[p] for p in cols) + self.gap * (len(cols) - 1)

    def show(self, cell: Cell, cols: list[int]) -> _Shown:
        """What a cell shows across visible column positions ``cols``, in the grid's mode."""
        if self.characters:
            return self.formatter.show(cell, width=self.block_width(cols))
        if self.units is not None:
            units = self.units(cols)
            if units is None:
                return _Shown(align="left")  # no width, no text
            return self.formatter.show(cell, column_width=units)
        return self.formatter.show(cell)

    # Origins ------------------------------------------------------------------------------

    def origins(self) -> tuple[list[_Origin], dict[tuple[int, int], _Shown]]:
        """Every text source, and what each visible cell shows on its own."""
        shown: dict[tuple[int, int], _Shown] = {}
        origins: list[_Origin] = []
        done_merges: set[CellRange] = set()
        covered_runs: set[tuple[int, int]] = set()
        for i, row in enumerate(self.rows):
            for p, column in enumerate(self.columns):
                position = (row.index, column.index)
                merge = self.merge_of.get(position)
                cell = self.cell(i, p)
                if merge is not None:
                    if merge in done_merges:
                        shown[position] = _Shown()
                        continue
                    done_merges.add(merge)
                    origin = self.merge_origin(merge)
                    if origin is not None:
                        origins.append(origin)
                        anchor = (origin.cell.row, origin.cell.column)
                        shown[anchor] = origin.shown
                    shown.setdefault(position, _Shown())
                    continue
                if position in covered_runs:
                    shown[position] = self.show(cell, [p])
                    continue
                cols = [p]
                if cell.style.h_align == "center_across" and cell.kind != "empty":
                    while cols[-1] + 1 < len(self.columns):
                        nxt = self.cell(i, cols[-1] + 1)
                        if nxt.kind != "empty" or nxt.style.h_align != "center_across":
                            break
                        if (nxt.row, nxt.column) in self.merge_of:
                            break
                        cols.append(cols[-1] + 1)
                        covered_runs.add((nxt.row, nxt.column))
                result = self.show(cell, cols)
                shown[position] = result
                if result.text:
                    # A center-across run spills like centered text; a merge does not
                    # (captured, layout rows 176 to 179).
                    spills = cell.kind == "text"
                    origins.append(_Origin(cell, result, [i], cols, spills))
        return origins, shown

    def merge_origin(self, merge: CellRange) -> _Origin | None:
        rows = [
            self.row_pos[r] for r in range(merge.min_row, merge.max_row + 1) if r in self.row_pos
        ]
        cols = [
            self.col_pos[c] for c in range(merge.min_col, merge.max_col + 1) if c in self.col_pos
        ]
        ref = self.sheet.ref
        if not rows or not cols or ref is None or (merge.min_row, merge.min_col) not in ref:
            return None  # nothing visible, or the anchor lies outside the range read
        anchor = self.sheet.cell(merge.min_row, merge.min_col)
        return _Origin(anchor, self.show(anchor, cols), rows, cols, False)

    # Rendering ----------------------------------------------------------------------------

    def render(self) -> tuple[tuple[RenderedRow, ...], list[int]]:
        """The rows, and the positions of the columns they keep."""
        origins, shown = self.origins()
        if not self.characters:
            return self.render_none(origins, shown), list(range(len(self.columns)))
        # Left to right, top to bottom: where two spills want the same free cell, the first
        # one placed takes it (a choice for characters mode, ARCHITECTURE.md §4.1: Excel draws
        # both texts over each other).
        origins.sort(key=lambda o: (o.rows[0], o.cols[0]))
        for origin in origins:
            origin.lines = self.content_lines(origin)
        canvas = _Canvas(self, self.heights(origins))
        claimed: set[tuple[int, int]] = set()
        for number, origin in enumerate(origins):
            self.place(canvas, number, origin, claimed)
        # Keep the added columns that spilled text reached, and those between them and the range.
        # A column is reached when text is drawn in it or in a gap beside it.
        reached = [
            p
            for p in self.added
            if canvas.used(2 * p) or canvas.used(2 * p + 1) or (p > 0 and canvas.used(2 * p - 1))
        ]
        real = [p for p in range(len(self.columns)) if p not in self.added]
        low = min([*reached, real[0]])
        high = max([*reached, real[-1]])
        kept = list(range(low, high + 1))
        return canvas.rows(origins, shown, kept), kept

    def render_none(
        self, origins: list[_Origin], shown: dict[tuple[int, int], _Shown]
    ) -> tuple[RenderedRow, ...]:
        source = {(o.cell.row, o.cell.column): o for o in origins}
        result = []
        for i, row in enumerate(self.rows):
            cells = []
            for p, column in enumerate(self.columns):
                cell = self.cell(i, p)
                position = (row.index, column.index)
                merge = self.merge_of.get(position)
                anchor = (merge.min_row, merge.min_col) if merge else position
                show = shown.get(position, _Shown())
                origin = source.get(position) if anchor == position else None
                text = origin.shown.text if origin else ""
                fill = self.anchor_fill(merge) if merge else cell.fill
                cells.append(
                    RenderedCell(
                        cell=cell,
                        text=show.text,
                        full_text=show.full_text,
                        overflow=show.overflow,
                        align=show.align,
                        format_color=show.format_color,
                        merge=merge,
                        spill=(0, 0),
                        lines=(text,),
                        source=position if origin else None,
                        fill=fill,
                        font=cell.font if origin else None,
                        color=(show.format_color or cell.font.color) if origin else None,
                    )
                )
            result.append(RenderedRow(row.index, row, 1, tuple(cells), ()))
        return tuple(result)

    def anchor_fill(self, merge: CellRange) -> str | None:
        ref = self.sheet.ref
        if ref is None or (merge.min_row, merge.min_col) not in ref:
            return None
        return self.sheet.cell(merge.min_row, merge.min_col).fill

    def content_lines(self, origin: _Origin) -> list[str]:
        text = origin.shown.text
        width = self.block_width(origin.cols)
        wraps = origin.cell.kind == "text" and (
            origin.cell.style.wrap or origin.shown.align in ("justify", "distributed")
        )
        if not wraps:
            # Captured (ARCHITECTURE.md §4.1): a line break in a cell without wrapping is not drawn.
            return [text.replace("\r", "").replace("\n", "")]
        origin.spills = False
        lines: list[str] = []
        for paragraph in text.replace("\r", "").split("\n"):
            lines.extend(textwrap.wrap(paragraph, max(width, 1)) or [""])
        return lines

    def heights(self, origins: list[_Origin]) -> list[int]:
        heights = [1] * len(self.rows)
        for origin in origins:
            if len(origin.rows) == 1:
                i = origin.rows[0]
                heights[i] = max(heights[i], len(origin.lines))
        for origin in origins:
            if len(origin.rows) > 1:
                have = sum(heights[i] for i in origin.rows)
                if len(origin.lines) > have:
                    heights[origin.rows[-1]] += len(origin.lines) - have
        return heights

    def free(self, i: int, p: int, claimed: set[tuple[int, int]]) -> bool:
        """Whether spilled text may enter visible column ``p`` of row ``i``.

        Hidden columns between are crossed whatever they hold (captured, ARCHITECTURE.md §4.5).
        """
        if not 0 <= p < len(self.columns) or (i, p) in claimed:
            return False
        cell = self.cell(i, p)
        return cell.kind == "empty" and (cell.row, cell.column) not in self.merge_of

    def place(
        self, canvas: _Canvas, number: int, origin: _Origin, claimed: set[tuple[int, int]]
    ) -> None:
        first, last = origin.cols[0], origin.cols[-1]
        x0 = self.starts[first]
        x1 = self.starts[last] + self.widths[last]
        width = x1 - x0
        align = origin.shown.align
        indent = origin.cell.style.indent * _INDENT
        lines = origin.lines
        total = sum(canvas.heights[i] for i in origin.rows)
        v_align = origin.cell.style.v_align
        top = (
            0
            if v_align == "top"
            else (total - len(lines)) // 2
            if v_align == "center"
            else total - len(lines)
        )
        spills = origin.spills and len(lines) == 1 and align in ("left", "right", "center")
        for k, text in enumerate(lines):
            n = len(text)
            if align == "fill" and n:
                text = text * max(1, width // n)  # captured: whole copies only (§4.1)
                n = len(text)
                x = x0
            elif align == "right":
                x = x1 - indent - n
            elif align == "center":
                x = x0 + (width - n) // 2
            else:
                x = x0 + indent
            low, high = x0, x1
            if spills:
                # Captured (§4.1): left-aligned text takes columns to the right and right-aligned
                # to the left, one whole column at a time, as far as it needs. Centered text
                # takes a column on each side at once while it overflows either edge.
                # Captured: once one side is blocked, the other grows only as far as its own text
                # needs.
                i = origin.rows[0]
                left = right = 0
                left_open = right_open = True
                while x < low or x + n > high:
                    left_open = left_open and self.free(i, first - left - 1, claimed)
                    right_open = right_open and self.free(i, last + right + 1, claimed)
                    centered = align == "center" and left_open and right_open
                    grow_left = left_open and (x < low or centered)
                    grow_right = right_open and (x + n > high or centered)
                    if not (grow_left or grow_right):
                        break
                    if grow_left:
                        left += 1
                        claimed.add((i, first - left))
                        low = self.starts[first - left]
                    if grow_right:
                        right += 1
                        claimed.add((i, last + right))
                        high = self.starts[last + right] + self.widths[last + right]
                origin.spill = (left, right)
                if left or right:
                    # Excel draws no gridlines across the columns a spill takes.
                    span = list(range(first - left, last + right + 1))
                    canvas.claim_gaps(origin.rows, span, number)
            canvas.put(origin.rows, top + k, x, text, number, low, high)
        if len(origin.cols) > 1:
            # A merge or a center-across run is one area: no separators inside it.
            canvas.claim_gaps(origin.rows, origin.cols, number)


class _Canvas:
    """Characters and their owners, per visible row and line."""

    def __init__(self, grid: _Grid, heights: list[int]) -> None:
        self.grid = grid
        self.heights = heights
        width = grid.total
        self.chars = [[[" "] * width for _ in range(h)] for h in heights]
        slots = 2 * len(grid.widths)
        self.owner = [[[_NOBODY] * slots for _ in range(h)] for h in heights]
        self.slot = [0] * width
        for p, start in enumerate(grid.starts):
            for x in range(start, start + grid.widths[p]):
                self.slot[x] = 2 * p
            for x in range(start + grid.widths[p], min(start + grid.widths[p] + grid.gap, width)):
                self.slot[x] = 2 * p + 1

    def put(
        self, rows: list[int], line: int, x: int, text: str, number: int, low: int, high: int
    ) -> None:
        for i in rows:  # find the row holding this line of the block
            if line < self.heights[i]:
                break
            line -= self.heights[i]
        else:
            return
        chars, owner, slot = self.chars[i][line], self.owner[i][line], self.slot
        for k, char in enumerate(text):
            at = x + k
            if low <= at < high:
                chars[at] = char
                owner[slot[at]] = number

    def claim_gaps(self, rows: list[int], cols: list[int], number: int) -> None:
        for i in rows:
            for owners in self.owner[i]:
                for p in cols[:-1]:
                    if owners[2 * p + 1] == _NOBODY:
                        owners[2 * p + 1] = number

    def used(self, slot: int) -> bool:
        """Whether any line of any row draws text in a column or gap slot."""
        return any(owners[slot] != _NOBODY for lines in self.owner for owners in lines)

    def rows(
        self, origins: list[_Origin], shown: dict[tuple[int, int], _Shown], kept: list[int]
    ) -> tuple[RenderedRow, ...]:
        grid = self.grid
        by_cell = {(o.cell.row, o.cell.column): o for o in origins}
        result = []
        for i, row in enumerate(grid.rows):
            cells: list[RenderedCell] = []
            for p in kept:
                column = grid.columns[p]
                cell = grid.cell(i, p)
                position = (row.index, column.index)
                merge = grid.merge_of.get(position)
                show = shown.get(position, _Shown())
                start, width = grid.starts[p], grid.widths[p]
                lines = tuple("".join(chars[start : start + width]) for chars in self.chars[i])
                owner = self._owner(i, 2 * p)
                origin = origins[owner] if owner != _NOBODY else None
                own = by_cell.get(position)
                cells.append(
                    RenderedCell(
                        cell=cell,
                        text=show.text,
                        full_text=show.full_text,
                        overflow=show.overflow,
                        align=show.align,
                        format_color=show.format_color,
                        merge=merge,
                        spill=own.spill if own else (0, 0),
                        lines=lines,
                        source=(origin.cell.row, origin.cell.column) if origin else None,
                        fill=grid.anchor_fill(merge) if merge else cell.fill,
                        font=origin.cell.font if origin else None,
                        color=_ink(origin),
                    )
                )
            gaps = []
            for k in range(len(cells) - 1):
                p = kept[k]
                start = grid.starts[p] + grid.widths[p]
                owner = self._owner(i, 2 * p + 1)
                origin = origins[owner] if owner != _NOBODY else None
                gap_lines = tuple(
                    "".join(chars[start : start + grid.gap])
                    if owners[2 * p + 1] != _NOBODY
                    else None
                    for chars, owners in zip(self.chars[i], self.owner[i], strict=True)
                )
                left, right = cells[k].fill, cells[k + 1].fill
                gaps.append(
                    Gap(
                        gap_lines,
                        left if left == right else None,
                        origin.cell.font if origin else None,
                        _ink(origin),
                    )
                )
            result.append(RenderedRow(row.index, row, self.heights[i], tuple(cells), tuple(gaps)))
        return tuple(result)

    def _owner(self, i: int, slot: int) -> int:
        for owners in self.owner[i]:
            if owners[slot] != _NOBODY:
                return owners[slot]
        return _NOBODY


def _ink(origin: _Origin | None) -> str | None:
    if origin is None:
        return None
    return origin.shown.format_color or origin.cell.font.color
