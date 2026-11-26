"""A worksheet's structure: what the layout reads besides the cells.

xlio reads the sheet format, column widths, row heights and hidden flags, the merged ranges and
the gridline settings (:func:`fill`), as Excel takes them. The reader's one pass of xlio's cell
stream (``_read._stream``) gives the bounds of every cell, which with the merged ranges give the
used range, the rows the file lists, and the marks the notices need: formulas without a cached
value, strings with runs, shown phonetic guides. The ``<dimension>`` element is ignored: some
writers omit it or get it wrong.
"""

from __future__ import annotations

import dataclasses
import math
from dataclasses import dataclass, field
from typing import Any

import xlio

from sheetpeek._model import CellRange
from sheetpeek._notices import CellSet


@dataclass(frozen=True, slots=True)
class ColumnSpec:
    """One ``<col>`` element: a run of columns sharing a width and a hidden flag."""

    min: int
    max: int
    width: float | None
    hidden: bool
    raw_width: str | None = None
    """``width`` as written; xlio derives ``width`` at Calibri 11's 6 pt digit."""


@dataclass(frozen=True, slots=True)
class Structure:
    """What the layout reads of one worksheet part besides its cells."""

    used_range: CellRange | None
    uncached_formulas: int
    """Formula cells without a cached value (a file written by a library, never opened by Excel)."""
    first_uncached_formula: str | None
    merges: tuple[CellRange, ...] = ()
    columns: tuple[ColumnSpec, ...] = ()
    rows: dict[int, tuple[float | None, bool, bool]] = field(default_factory=dict)
    """Row index → (height in points or ``None``, hidden, autofit), for rows that set any: a
    height without ``customHeight`` is no height and autofit (``Row.autofit``)."""
    default_column_width: float | None = None
    base_column_width: int = 8
    default_row_height: float | None = None
    show_gridlines: bool = True
    listed_rows: frozenset[int] = frozenset()
    """Rows the file has a ``<row>`` element for, which decides some heights (§4.7)."""
    has_sheet_format: bool = True
    """Whether the part has ``<sheetFormatPr>``."""
    custom_row_height: bool = False
    """``<sheetFormatPr customHeight>``: ``defaultRowHeight`` is the rows' height."""
    zero_row_height: bool = False
    """``<sheetFormatPr zeroHeight>``: rows without a ``<row>`` element are hidden."""
    thick_top: bool = False
    thick_bottom: bool = False
    show_headings: bool = True
    """The first ``<sheetView>``'s ``showRowColHeaders``."""
    grid_color_id: int | None = None
    """The first ``<sheetView>``'s ``colorId`` under ``defaultGridColor="0"``, else ``None``."""
    uncached: CellSet = field(default_factory=CellSet)
    """Formula cells without a cached value."""
    runs: CellSet = field(default_factory=CellSet)
    """Cells whose string has rich-text runs."""
    phonetic: CellSet = field(default_factory=CellSet)
    """Cells that show a phonetic guide their string has."""
    row_styles: tuple[tuple[int, int], ...] = ()
    """(row, style index) of rows with a format of their own (``customFormat``)."""
    column_styles: tuple[tuple[int, int, int], ...] = ()
    """(first, last, style index) of ``<col>`` elements with a style."""
    extra: dict[str, Any] = field(default_factory=dict)
    """What features found (links, the outline), by feature."""


def fill(found: Structure, sheet: xlio.Sheet) -> Structure:
    """``found`` with what xlio reads of ``sheet``: format, columns, rows, merges, gridlines.

    Each value as Excel takes it (xlio's captures). A merged range occupies its cells even where
    the file has no ``<c>`` for them, so the used range grows to hold every merge.
    """
    merges = tuple(CellRange.parse(ref) for ref in sheet.merges)
    used = found.used_range
    for merge in merges:
        used = merge if used is None else _union(used, merge)
    form = sheet.format
    custom_height, zero_height = form.custom_height, form.zero_height
    if custom_height and form.default_row_height == 0.0:
        # An enormous default acts as zeroHeight: a row with an element keeps its height
        # (captured).
        custom_height, zero_height = False, True
    views = sheet.views
    view = views[0] if views else None
    rows = sheet.rows
    return dataclasses.replace(
        found,
        used_range=used,
        merges=merges,
        columns=tuple(
            ColumnSpec(column.min, column.max, column.width, column.hidden, column.raw_width)
            for column in sheet.columns
        ),
        rows={
            index: (
                row.height if row.custom_height else None,
                row.hidden,
                row.height is not None and not row.custom_height,
            )
            for index, row in rows.items()
            if row.height is not None or row.hidden
        },
        default_column_width=form.default_col_width,
        base_column_width=form.base_col_width,
        default_row_height=form.default_row_height,
        has_sheet_format=form.source is not None,
        custom_row_height=custom_height,
        zero_row_height=zero_height,
        thick_top=form.thick_top,
        thick_bottom=form.thick_bottom,
        show_gridlines=view.show_grid_lines if view else True,
        show_headings=view.show_row_col_headers if view else True,
        grid_color_id=view.color_id if view and not view.default_grid_color else None,
        row_styles=tuple(
            (index, row.style) for index, row in rows.items() if row.custom_format and row.style
        ),
        column_styles=tuple(
            (column.min, column.max, column.style) for column in sheet.columns if column.style
        ),
    )


def _union(a: CellRange, b: CellRange) -> CellRange:
    return CellRange(
        min(a.min_row, b.min_row),
        min(a.min_col, b.min_col),
        max(a.max_row, b.max_row),
        max(a.max_col, b.max_col),
    )


XLIO_DIGIT = 6.0
"""The digit, in points, at which xlio derives a column's width: Calibri 11's."""
_WIDTH_WRAP = 65_536  # points: Excel keeps a column's width in 16 bits
_INT_RANGE = 2**31


def width_at_digit(spec: ColumnSpec, digit: float) -> float | None:
    """A column's width as Excel takes it under a Normal font whose digit is ``digit`` points.

    Excel keeps a width in points modulo 65,536 at the Normal font's digit, a negative width with
    one more point, and a width whose points pass 2**31 as one point (captured ``package`` at
    Calibri 11, ``package_more`` at Calibri 14: 1e6 is 53,184 pt and -1 is 65,530 at a 7 pt
    digit). xlio does the same at a 6 pt digit, and its 0 (a hidden column: NaN, text, 8e305 and
    past) stands at any digit.
    """
    width, raw = spec.width, spec.raw_width
    if digit == XLIO_DIGIT or not width or raw is None:
        return width
    try:
        value = float(raw)
    except ValueError:
        return width
    points = value * digit
    if 0 <= points < _WIDTH_WRAP:
        return value
    if abs(points) >= _INT_RANGE:
        return 1 / digit
    wrapped = math.fmod(points, _WIDTH_WRAP) % _WIDTH_WRAP
    return (wrapped + (1 if value < 0 else 0)) / digit
