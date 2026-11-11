"""The structure scan: what openpyxl's read-only mode does not expose, read from the sheet XML.

One ``iterparse`` pass over a worksheet part records the sheet view, the sheet format, column
widths, row heights and hidden flags, merged ranges, and the bounds of every ``<c>``
element, which with the merged ranges give the used range. Cell contents are not read, apart
from noticing a formula without a cached value. The ``<dimension>`` element is ignored: some
writers omit it or get it wrong.
"""

from __future__ import annotations

import math
import xml.etree.ElementTree as ET
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import IO, Any

from sheetpeek._model import CellRange
from sheetpeek._notices import CellSet
from sheetpeek._package import PackageError

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_SHEET_DATA = f"{_NS}sheetData"
_ROW = f"{_NS}row"
_CELL = f"{_NS}c"
_FORMULA = f"{_NS}f"
_VALUE = f"{_NS}v"
_COL = f"{_NS}col"
_MERGE = f"{_NS}mergeCell"
_SHEET_VIEW = f"{_NS}sheetView"
_SHEET_FORMAT = f"{_NS}sheetFormatPr"
_INLINE = f"{_NS}is"
_RUN = f"{_NS}r"
RICH = 1
"""A shared string's mark: it has runs."""
PHONETIC = 2
"""A shared string's mark: it has a phonetic guide."""
# The numbers a file can give a column or a row, as Excel takes them (ARCHITECTURE.md §5.5).
_DIGIT_POINTS = 6  # at Calibri 11, the only font this was captured with
_WIDTH_WRAP = 65_536  # points: Excel keeps a column's width in 16 bits
_WIDEST_FILE_WIDTH = 10_922.0  # the widest that does not wrap
_ONE_POINT_WIDTH = 0.1  # a file width of one point: what Excel makes of 1e9 to 1e30
_INT_RANGE = 2**31
_ENORMOUS = 1e30  # past this, a width or a height is 0; 1e30 and 1e308 are captured
_TALLEST = 409.5
_TWIPS = 20
_HEIGHT_WRAP = 65_536  # twips
_WIDEST_DEFAULT = 255.0  # 300 is refused; between them is not captured
_WIDEST_BASE = 255


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
    merges: tuple[CellRange, ...]
    columns: tuple[ColumnSpec, ...]
    rows: dict[int, tuple[float | None, bool]]
    """Row index → (height in points or ``None``, hidden), for rows that set either."""
    default_column_width: float | None
    base_column_width: int
    default_row_height: float | None
    show_gridlines: bool
    uncached_formulas: int
    """Formula cells without a cached value (a file written by a library, never opened by Excel)."""
    first_uncached_formula: str | None
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
    root: str = "worksheet"
    """The part's root element, without namespace: ``worksheet``, ``dialogsheet``, ..."""
    children: tuple[ET.Element, ...] = ()
    """The root's children other than ``<sheetData>``, for the reader's detectors."""
    uncached: CellSet = field(default_factory=CellSet)
    """Formula cells without a cached value."""
    runs: CellSet = field(default_factory=CellSet)
    """Cells whose string has rich-text runs."""
    phonetic: CellSet = field(default_factory=CellSet)
    """Cells that show a phonetic guide their string has."""
    outline_rows: int = 0
    """Rows with an outline level."""
    outline_columns: int = 0
    """Columns with an outline level."""
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

    def __init__(self, rels: dict[str, tuple[str, bool]], marks: bytes | None) -> None:
        self.rels = rels
        self.marks = marks
        self.min_row = self.min_col = 1 << 30
        self.max_row = self.max_col = 0
        self.row_index = 0
        self.rows: dict[int, tuple[float | None, bool]] = {}
        self.columns: list[ColumnSpec] = []
        self.merges: list[CellRange] = []
        self.uncached = CellSet()
        self.first_uncached: str | None = None
        self.runs = CellSet()
        self.phonetic = CellSet()
        self.show_headings = True
        self.root = "worksheet"
        self.children: list[ET.Element] = []
        self.outline_rows = self.outline_columns = 0
        self.row_styles: list[tuple[int, int]] = []
        self.column_styles: list[tuple[int, int, int]] = []
        self.empty_strings: set[tuple[int, int]] = set()
        self.default_column_width: float | None = None
        self.base_column_width = 8
        self.default_row_height: float | None = None
        self.show_gridlines = True
        self.listed: set[int] = set()
        self.sheet_format: dict[str, str] | None = None
        self.seen_view = False
        self.extra: dict[str, Any] = {}

    def finish(self) -> Structure:
        """The structure found."""
        min_row, max_row, min_col, max_col = self.min_row, self.max_row, self.min_col, self.max_col
        # A merged range occupies its cells even where the file has no <c> for them.
        for merge in self.merges:
            min_row, max_row = min(min_row, merge.min_row), max(max_row, merge.max_row)
            min_col, max_col = min(min_col, merge.min_col), max(max_col, merge.max_col)
        used = CellRange(min_row, min_col, max_row, max_col) if max_row else None
        sheet_format = self.sheet_format or {}
        custom_height = _flag(sheet_format.get("customHeight"))
        zero_height = _flag(sheet_format.get("zeroHeight"))
        if custom_height and self.default_row_height == 0.0:
            # An enormous default acts as zeroHeight: a row with an element keeps its height
            # (captured).
            custom_height, zero_height = False, True
        return Structure(
            used_range=used,
            merges=tuple(self.merges),
            columns=tuple(self.columns),
            rows=self.rows,
            default_column_width=self.default_column_width,
            base_column_width=self.base_column_width,
            default_row_height=self.default_row_height,
            show_gridlines=self.show_gridlines,
            uncached_formulas=self.uncached.count,
            first_uncached_formula=self.first_uncached,
            empty_strings=frozenset(self.empty_strings),
            listed_rows=frozenset(self.listed),
            has_sheet_format=self.sheet_format is not None,
            custom_row_height=custom_height,
            zero_row_height=zero_height,
            thick_top=_flag(sheet_format.get("thickTop")),
            thick_bottom=_flag(sheet_format.get("thickBottom")),
            show_headings=self.show_headings,
            root=self.root,
            children=tuple(self.children),
            uncached=self.uncached,
            runs=self.runs,
            phonetic=self.phonetic,
            outline_rows=self.outline_rows,
            outline_columns=self.outline_columns,
            row_styles=tuple(self.row_styles),
            column_styles=tuple(self.column_styles),
            extra=dict(self.extra),
        )


def scan(
    source: IO[bytes],
    rels: dict[str, tuple[str, bool]] | None = None,
    marks: bytes | None = None,
) -> Structure:
    """Scan one worksheet part.

    Args:
        source: The worksheet part, opened for reading.
        rels: The part's relationships (id → target, external), for hyperlink targets.
        marks: Per shared string, :data:`RICH` and :data:`PHONETIC` bits, or ``None`` when no
            shared string has either.
    """
    state = Scan(rels or {}, marks)
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
    """A row: its height and hidden flag, the used range, and its cells' marks."""
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
    if _level(element):
        state.outline_rows += 1
    if _flag(element.get("customFormat")) and (style := _style(element.get("s"))):
        state.row_styles.append((row_index, style))
    height = element.get("ht")
    hidden = _flag(element.get("hidden"))
    if height is not None or hidden:
        state.rows[row_index] = (_row_height(height) if height is not None else None, hidden)


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


@handler(_COL)
def _col(state: Scan, element: ET.Element) -> None:
    """A run of columns: width, hidden flag, outline level, style."""
    first = int(element.get("min", "1"))
    last = int(element.get("max", element.get("min", "1")))
    width = element.get("width")
    if _level(element):
        state.outline_columns += 1
    if style := _style(element.get("style")):
        state.column_styles.append((first, last, style))
    state.columns.append(
        ColumnSpec(
            first,
            last,
            _column_width(width) if width is not None else None,
            _flag(element.get("hidden")),
        )
    )


@handler(_MERGE)
def _merge(state: Scan, element: ET.Element) -> None:
    ref = element.get("ref")
    if ref:
        state.merges.append(CellRange.parse(ref))


@handler(_SHEET_VIEW)
def _view(state: Scan, element: ET.Element) -> None:
    """The first sheet view only: the main window's."""
    if state.seen_view:
        return
    state.seen_view = True
    state.show_gridlines = _flag(element.get("showGridLines"), default=True)
    state.show_headings = _flag(element.get("showRowColHeaders"), default=True)


@handler(_SHEET_FORMAT)
def _sheet_format(state: Scan, element: ET.Element) -> None:
    state.sheet_format = dict(element.attrib)
    if (value := element.get("defaultColWidth")) is not None:
        state.default_column_width = _default_width(value)
    if (value := element.get("baseColWidth")) is not None:
        state.base_column_width = _base_width(value)
    if (value := element.get("defaultRowHeight")) is not None:
        state.default_row_height = _default_height(value)


def _column_width(text: str) -> float:
    """A ``<col width>`` as Excel takes it, in points modulo 65,536 (captured)."""
    width = float(text)
    if not math.isfinite(width) or abs(width) > _ENORMOUS:
        return 0.0
    if 0 <= width <= _WIDEST_FILE_WIDTH:
        return width
    points = width * _DIGIT_POINTS
    if abs(points) >= _INT_RANGE:
        return _ONE_POINT_WIDTH
    wrapped = math.fmod(points, _WIDTH_WRAP) % _WIDTH_WRAP
    return (wrapped + (1 if width < 0 else 0)) / _DIGIT_POINTS


def _row_height(text: str) -> float | None:
    """A ``<row ht>`` as Excel takes it, or ``None`` when it ignores it (captured)."""
    height = abs(float(text))
    if not math.isfinite(height):
        return None
    if height > _ENORMOUS:
        return 0.0
    if height <= _TALLEST:
        return height
    twips = height * _TWIPS
    if twips >= _INT_RANGE:
        return _TALLEST
    return min(_TALLEST, math.floor((twips % _HEIGHT_WRAP) / _TWIPS * 4) / 4)


def _default_width(text: str) -> float | None:
    """``defaultColWidth`` as Excel takes it, or ``None`` when it ignores it (captured)."""
    width = float(text)
    if math.isnan(width) or width == math.inf:
        return None
    if width < 0 or _ENORMOUS >= width > _WIDEST_DEFAULT:
        raise PackageError(f"defaultColWidth of {text}, which Excel refuses")
    return 0.0 if width > _ENORMOUS else width


def _base_width(text: str) -> int:
    """``baseColWidth`` as Excel takes it (captured): out of an int's range or negative, ignored."""
    width = int(text)
    if width < 0 or width >= _INT_RANGE:
        return 8
    if width > _WIDEST_BASE:
        raise PackageError(f"baseColWidth of {text}, which Excel refuses")
    return width


def _default_height(text: str) -> float:
    """``defaultRowHeight``: Excel refuses NaN and infinity (captured)."""
    height = _row_height(text)
    if height is None:
        raise PackageError(f"defaultRowHeight of {text}, which Excel refuses")
    return height


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


def _level(element: ET.Element) -> bool:
    """Whether a ``<row>`` or ``<col>`` has an outline level above 0."""
    value = element.get("outlineLevel")
    return value is not None and value.strip() not in ("0", "")


def _style(value: str | None) -> int:
    """A style index, or 0 when absent or not a number."""
    return int(value) if value and value.isdigit() else 0


def _flag(value: str | None, *, default: bool = False) -> bool:
    if value is None:
        return default
    return value in ("1", "true")
