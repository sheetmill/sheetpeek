"""A worksheet's structure: what the layout reads besides the cells.

xlio reads the sheet format, column widths, row heights and hidden flags, the merged ranges and
the gridline settings (:func:`fill`), as Excel takes them. One ``iterparse`` pass over the
part (:func:`scan`) records the bounds of every ``<c>`` element, which with the merged ranges
give the used range, the rows the file lists, and what openpyxl's read-only mode does not expose
about cells. Cell contents are not read, apart from noticing a formula without a cached value.
The ``<dimension>`` element is ignored: some writers omit it or get it wrong.
"""

from __future__ import annotations

import dataclasses
import xml.etree.ElementTree as ET
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import IO, Any

import xlio

from sheetpeek._model import CellRange
from sheetpeek._notices import CellSet

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_SHEET_DATA = f"{_NS}sheetData"
_ROW = f"{_NS}row"
_CELL = f"{_NS}c"
_FORMULA = f"{_NS}f"
_VALUE = f"{_NS}v"
_COL = f"{_NS}col"
_SHEET_VIEW = f"{_NS}sheetView"
_INLINE = f"{_NS}is"
_RUN = f"{_NS}r"
RICH = 1
"""A shared string's mark: it has runs."""
PHONETIC = 2
"""A shared string's mark: it has a phonetic guide."""


@dataclass(frozen=True, slots=True)
class ColumnSpec:
    """One ``<col>`` element: a run of columns sharing a width and a hidden flag."""

    min: int
    max: int
    width: float | None
    hidden: bool


@dataclass(frozen=True, slots=True)
class Structure:
    """What the scan found in one worksheet part."""

    used_range: CellRange | None
    uncached_formulas: int
    """Formula cells without a cached value (a file written by a library, never opened by Excel)."""
    first_uncached_formula: str | None
    merges: tuple[CellRange, ...] = ()
    columns: tuple[ColumnSpec, ...] = ()
    rows: dict[int, tuple[float | None, bool]] = field(default_factory=dict)
    """Row index → (height in points or ``None``, hidden), for rows that set either."""
    default_column_width: float | None = None
    base_column_width: int = 8
    default_row_height: float | None = None
    show_gridlines: bool = True
    empty_strings: frozenset[tuple[int, int]] = frozenset()
    """(row, column) of formula cells whose cached result is the empty string."""
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
    root: str = "worksheet"
    """The part's root element, without namespace: ``worksheet``, ``dialogsheet``, ..."""
    children: tuple[ET.Element, ...] = ()
    """The root's children other than ``<sheetData>``, for the reader's detectors."""
    uncached: CellSet = field(default_factory=CellSet)
    """Formula cells without a cached value."""
    runs: CellSet = field(default_factory=CellSet)
    """Cells whose string has rich-text runs."""
    inline_runs: CellSet = field(default_factory=CellSet)
    """Cells whose inline string has rich-text runs, which openpyxl's read-only reader drops."""
    phonetic: CellSet = field(default_factory=CellSet)
    """Cells that show a phonetic guide their string has."""
    row_styles: tuple[tuple[int, int], ...] = ()
    """(row, style index) of rows with a format of their own (``customFormat``)."""
    column_styles: tuple[tuple[int, int, int], ...] = ()
    """(first, last, style index) of ``<col>`` elements with a style."""
    extra: dict[str, Any] = field(default_factory=dict)
    """What features' own handlers found, by feature."""


Handler = Callable[["Scan", ET.Element], None]
"""Called with the scan's state and an element, when the element ends."""

HANDLERS: dict[str, list[Handler]] = {}
"""Handlers by tag. A feature adds its own with :func:`handler` and keeps what it finds in
``Scan.extra`` under its own key, so features do not meet in this module."""


def handler(tag: str) -> Callable[[Handler], Handler]:
    """Register a handler for elements with ``tag`` (namespace included)."""

    def register(function: Handler) -> Handler:
        HANDLERS.setdefault(tag, []).append(function)
        return function

    return register


class Scan:
    """What the scan has found so far; the handlers read and add to it."""

    def __init__(self, marks: bytes | None) -> None:
        self.marks = marks
        self.min_row = self.min_col = 1 << 30
        self.max_row = self.max_col = 0
        self.row_index = 0
        self.uncached = CellSet()
        self.first_uncached: str | None = None
        self.runs = CellSet()
        self.inline_runs = CellSet()
        self.phonetic = CellSet()
        self.root = "worksheet"
        self.children: list[ET.Element] = []
        self.empty_strings: set[tuple[int, int]] = set()
        self.listed: set[int] = set()
        self.extra: dict[str, Any] = {}

    def finish(self) -> Structure:
        """What the scan found; :func:`fill` adds the rest."""
        found = self.max_row
        used = CellRange(self.min_row, self.min_col, self.max_row, self.max_col) if found else None
        return Structure(
            used_range=used,
            uncached_formulas=self.uncached.count,
            first_uncached_formula=self.first_uncached,
            empty_strings=frozenset(self.empty_strings),
            listed_rows=frozenset(self.listed),
            root=self.root,
            children=tuple(self.children),
            uncached=self.uncached,
            runs=self.runs,
            inline_runs=self.inline_runs,
            phonetic=self.phonetic,
            extra=dict(self.extra),
        )


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
            ColumnSpec(column.min, column.max, column.width, column.hidden)
            for column in sheet.columns
        ),
        rows={
            index: (row.height, row.hidden)
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


def scan(
    source: IO[bytes],
    marks: bytes | None = None,
) -> Structure:
    """Scan one worksheet part.

    Args:
        source: The worksheet part, opened for reading.
        marks: Per shared string, :data:`RICH` and :data:`PHONETIC` bits, or ``None`` when no
            shared string has either.
    """
    state = Scan(marks)
    depth = 0
    sheet_data: ET.Element | None = None
    for event, element in ET.iterparse(source, events=("start", "end")):
        tag = element.tag
        if tag in (_CELL, _VALUE):
            continue  # cells are handled at the end of their row
        if event == "start":
            depth += 1
            if depth == 1:
                state.root = tag.rpartition("}")[2]
            elif tag == _SHEET_DATA:
                sheet_data = element
            continue
        depth -= 1
        if depth == 1 and tag != _SHEET_DATA:
            state.children.append(element)
        for function in HANDLERS.get(tag, ()):
            function(state, element)
        if tag == _ROW and sheet_data is not None:
            sheet_data.clear()  # the row and its cells are done
    return state.finish()


@handler(_ROW)
def _row(state: Scan, element: ET.Element) -> None:
    """A row: that the file lists it, the used range, and its cells' marks."""
    r = element.get("r")
    row_index = state.row_index = int(r) if r else state.row_index + 1
    state.listed.add(row_index)
    if len(element):
        first_col, last_col = _column_extent(element)
        state.min_row = min(state.min_row, row_index)
        state.max_row = max(state.max_row, row_index)
        state.min_col = min(state.min_col, first_col)
        state.max_col = max(state.max_col, last_col)
        _cells(state, element, row_index)


def _cells(state: Scan, row: ET.Element, row_index: int) -> None:
    """Uncached formulas, empty-string results, and strings with runs or phonetic guides."""
    marks = state.marks
    column = 0
    for cell in row:
        ref = cell.get("r")
        column = _position(ref)[1] if ref else column + 1
        kind = cell.get("t")
        # The common cell is one <v>; only others can be formulas without a value.
        if (len(cell) != 1 or cell[0].tag == _FORMULA) and (
            cell.find(_FORMULA) is not None and not cell.findtext(_VALUE)
        ):
            if kind == "str":
                # A formula whose cached result is the empty string. openpyxl reads it as an
                # empty cell, but it is not empty: it stops spilled text.
                if ref:
                    state.empty_strings.add(_position(ref))
            else:
                state.uncached.add(row_index, column)
                if state.first_uncached is None:
                    state.first_uncached = ref or f"row {row_index}"
        elif kind == "s" and marks is not None:
            index = cell.findtext(_VALUE) or ""
            mark = marks[int(index)] if index.isdigit() and int(index) < len(marks) else 0
            if mark & RICH:
                state.runs.add(row_index, column)
            if mark & PHONETIC and _flag(cell.get("ph")):
                state.phonetic.add(row_index, column)
        elif kind == "inlineStr":
            inline = cell.find(_INLINE)
            if inline is not None and inline.find(_RUN) is not None:
                state.runs.add(row_index, column)
                state.inline_runs.add(row_index, column)


def _column_extent(row: ET.Element) -> tuple[int, int]:
    """First and last column of a row's cells, which the schema orders by column."""
    first, last = row[0].get("r"), row[-1].get("r")
    if first and last:
        return _position(first)[1], _position(last)[1]
    # Some cells omit r: each then follows its predecessor.
    col = 0
    low, high = 1 << 30, 0
    for cell in row:
        r = cell.get("r")
        col = _position(r)[1] if r else col + 1
        low, high = min(low, col), max(high, col)
    return low, high


def _position(ref: str) -> tuple[int, int]:
    """Row and column of an A1 reference such as ``"AB12"``, without validation."""
    col = 0
    for i, char in enumerate(ref):
        code = ord(char)
        if 65 <= code <= 90:
            col = col * 26 + code - 64
        elif 97 <= code <= 122:
            col = col * 26 + code - 96
        elif char != "$":
            return int(ref[i:].replace("$", "")), col
    raise ValueError(f"not a cell reference: {ref!r}")


def _flag(value: str | None, *, default: bool = False) -> bool:
    if value is None:
        return default
    return value in ("1", "true")
