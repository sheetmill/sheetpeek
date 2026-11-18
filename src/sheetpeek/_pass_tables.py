"""The tables pass: the style each table cell is displayed in, before the cells are laid out.

A table draws its style's elements over regions of its range (§4.20): the whole table, row and
column stripes over the data rows, the first and last column, the header and totals rows and
their corner cells. The elements stack in ECMA-376's order (:data:`ELEMENT_ORDER`), a later one
winning property by property and edge segment by edge segment, whatever a line's weight.
Stripes count visible rows and columns only. A region's outer sides draw on its outline and its
``horizontal`` and ``vertical`` sides on its inner boundaries; each stripe band is a region. A
cell's own side on a segment, from either cell, replaces the table's line there; between the
cells the borders pass's rule then decides as for any two sides.

The result, ``Draft.display_styles``, maps a position to the style :func:`_dxf.stack` gives;
positions it leaves out are displayed in their own style. It runs before the cells, since a
table's bold changes text widths and spill. The screen suite ``tables`` captures all of it.
"""

from __future__ import annotations

import itertools
from typing import TYPE_CHECKING

from sheetpeek._dxf import stack
from sheetpeek._fill_border_model import Borders, Side
from sheetpeek._model import CellRange, CellStyle, Sheet
from sheetpeek._table_model import (
    ELEMENT_ORDER,
    DifferentialFormat,
    Table,
    TableStyle,
    TableStyleElementType,
)

if TYPE_CHECKING:
    from sheetpeek._layout import Draft

Position = tuple[int, int]
_Key = tuple[int, tuple[int, ...], bool, tuple[Side | None, ...]]


def tables_pass(draft: Draft) -> None:
    """Fill ``draft.display_styles`` for the cells the sheet's tables style."""
    sheet = draft.sheet
    ref = sheet.ref
    if ref is None or not sheet.tables:
        return
    from sheetpeek._layout import _column_hidden as column_hidden
    from sheetpeek._layout import _row_hidden as row_hidden

    hidden = {row.index for row in sheet.rows if row_hidden(row, sheet)}
    hidden_columns = {column.index for column in sheet.columns if column_hidden(column)}
    formats: dict[Position, list[DifferentialFormat]] = {}
    above: dict[Position, Side] = {}  # the line along the top of the cell at a position
    left: dict[Position, Side] = {}  # the line along its left side
    for table in sheet.tables:
        style = _style_of(table, sheet)
        if style is None:
            continue
        for kind, region in regions(table, style, hidden, hidden_columns):
            element = style.element(kind)
            if element is None:
                continue
            _paint(element.format, region, ref, formats, above, left)
    draft.display_styles.update(_stacked(sheet, ref, formats, above, left))


def _style_of(table: Table, sheet: Sheet) -> TableStyle | None:
    """The style a table draws, or ``None``: no style info, no name, or a name not found."""
    info = table.style
    if info is None or not info.name:
        return None
    return sheet.table_style(info.name)


def regions(
    table: Table,
    style: TableStyle,
    hidden: set[int],
    hidden_columns: frozenset[int] | set[int] = frozenset(),
) -> list[tuple[TableStyleElementType, CellRange]]:
    """Each element's regions on the table, in the order that stacks them.

    ``hidden`` and ``hidden_columns`` hold the hidden rows and columns, which stripes do not
    count.
    """
    info = table.style
    assert info is not None
    ref = table.ref
    first, last = ref.min_col, ref.max_col
    data_top, data_bottom = table.data_rows
    found: dict[TableStyleElementType, list[CellRange]] = {}

    def size(kind: TableStyleElementType) -> int:
        element = style.element(kind)
        return max(1, element.size) if element is not None else 1

    if data_top <= data_bottom:
        if info.show_column_stripes:
            bands = _bands(
                [c for c in range(first, last + 1) if c not in hidden_columns],
                size("firstColumnStripe"),
                size("secondColumnStripe"),
            )
            for kind, (low, high) in zip(itertools.cycle(_COLUMN_STRIPES), bands):
                found.setdefault(kind, []).append(CellRange(data_top, low, data_bottom, high))
        if info.show_row_stripes:
            visible = [r for r in range(data_top, data_bottom + 1) if r not in hidden]
            bands = _bands(visible, size("firstRowStripe"), size("secondRowStripe"))
            for kind, (low, high) in zip(itertools.cycle(_ROW_STRIPES), bands):
                found.setdefault(kind, []).append(CellRange(low, first, high, last))
    found["wholeTable"] = [ref]
    if info.show_last_column:
        found["lastColumn"] = [CellRange(ref.min_row, last, ref.max_row, last)]
    if info.show_first_column:
        found["firstColumn"] = [CellRange(ref.min_row, first, ref.max_row, first)]
    if table.header_row_count > 0:
        bottom = ref.min_row + table.header_row_count - 1
        found["headerRow"] = [CellRange(ref.min_row, first, bottom, last)]
        if info.show_first_column:
            found["firstHeaderCell"] = [CellRange(ref.min_row, first, bottom, first)]
        if info.show_last_column:
            found["lastHeaderCell"] = [CellRange(ref.min_row, last, bottom, last)]
    if table.totals_row_count > 0:
        top = ref.max_row - table.totals_row_count + 1
        found["totalRow"] = [CellRange(top, first, ref.max_row, last)]
        if info.show_first_column:
            found["firstTotalCell"] = [CellRange(top, first, ref.max_row, first)]
        if info.show_last_column:
            found["lastTotalCell"] = [CellRange(top, last, ref.max_row, last)]
    return [(kind, region) for kind in ELEMENT_ORDER for region in found.get(kind, ())]


_ROW_STRIPES: tuple[TableStyleElementType, ...] = ("firstRowStripe", "secondRowStripe")
_COLUMN_STRIPES: tuple[TableStyleElementType, ...] = ("firstColumnStripe", "secondColumnStripe")


def _bands(indexes: list[int], first: int, second: int) -> list[tuple[int, int]]:
    """The bands over ``indexes``, ``first`` and ``second`` of them by turns: (first, last)."""
    bands: list[tuple[int, int]] = []
    at, sizes = 0, (first, second)
    while at < len(indexes):
        band = indexes[at : at + sizes[len(bands) % 2]]
        bands.append((band[0], band[-1]))
        at += len(band)
    return bands


def _paint(
    found: DifferentialFormat,
    region: CellRange,
    ref: CellRange,
    formats: dict[Position, list[DifferentialFormat]],
    above: dict[Position, Side],
    left: dict[Position, Side],
) -> None:
    """Lay one element's format over its region, clipped to the range read."""
    rows = range(max(region.min_row, ref.min_row), min(region.max_row, ref.max_row) + 1)
    columns = range(max(region.min_col, ref.min_col), min(region.max_col, ref.max_col) + 1)
    sides = any(
        s is not None
        for s in (
            found.left,
            found.right,
            found.top,
            found.bottom,
            found.horizontal,
            found.vertical,
        )
    )
    paints = found.fill is not None or any(
        p is not None
        for p in (found.color, found.bold, found.italic, found.underline, found.strike)
    )
    for row in rows:
        for column in columns:
            if paints:
                formats.setdefault((row, column), []).append(found)
            if not sides:
                continue
            top = found.top if row == region.min_row else found.horizontal
            bottom = found.bottom if row == region.max_row else found.horizontal
            west = found.left if column == region.min_col else found.vertical
            east = found.right if column == region.max_col else found.vertical
            if top is not None:
                above[(row, column)] = top
            if bottom is not None:
                above[(row + 1, column)] = bottom
            if west is not None:
                left[(row, column)] = west
            if east is not None:
                left[(row, column + 1)] = east


def _stacked(
    sheet: Sheet,
    ref: CellRange,
    formats: dict[Position, list[DifferentialFormat]],
    above: dict[Position, Side],
    left: dict[Position, Side],
) -> dict[Position, CellStyle]:
    """The displayed style of every position a table format or line reaches."""
    normal_color = sheet._normal_color
    done: dict[_Key, CellStyle] = {}
    result: dict[Position, CellStyle] = {}
    positions = set(formats)
    for row, column in (*above, *left):
        for position in ((row, column), (row - 1, column), (row, column - 1)):
            if position in ref:
                positions.add(position)
    for position in positions:
        row, column = position
        if position not in ref:
            continue
        own = sheet.cell(row, column).style
        sides = (
            _segment(sheet, ref, left, (row, column), across=False),
            _segment(sheet, ref, left, (row, column + 1), across=False),
            _segment(sheet, ref, above, (row, column), across=True),
            _segment(sheet, ref, above, (row + 1, column), across=True),
        )
        found = formats.get(position, [])
        own_color = normal_color.get(id(own)) is not own
        key = (id(own), tuple(id(f) for f in found), own_color, sides)
        drawn = done.get(key)
        if drawn is None:
            beneath = Borders(*sides) if any(s is not None for s in sides) else None
            drawn = done[key] = stack(own, found, own_color=own_color, sides=beneath)
        if drawn is not own:
            result[position] = drawn
    return result


def _segment(
    sheet: Sheet, ref: CellRange, lines: dict[Position, Side], at: Position, *, across: bool
) -> Side | None:
    """The table's line on one segment, unless a cell's own side lies on it.

    The segment runs along the top (``across``) or the left of the cell at ``at``; an own side
    of that cell or of the one above or left of it, facing the segment, replaces the line.
    """
    line = lines.get(at)
    if line is None:
        return None
    row, column = at
    if (row, column) in ref:
        borders = sheet.cell(row, column).style.borders
        if borders is not None and (borders.top if across else borders.left) is not None:
            return None
    other = (row - 1, column) if across else (row, column - 1)
    if other in ref:
        borders = sheet.cell(*other).style.borders
        if borders is not None and (borders.bottom if across else borders.right) is not None:
            return None
    return line
