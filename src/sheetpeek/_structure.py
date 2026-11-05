"""The structure scan: what openpyxl's read-only mode does not expose, read from the sheet XML.

One ``iterparse`` pass over a worksheet part records the sheet view, the sheet format, column
widths, row heights and hidden flags, merged ranges, hyperlinks, and the bounds of every ``<c>``
element, which with the merged ranges give the used range. Cell contents are not read, apart
from noticing a formula without a cached value. The ``<dimension>`` element is ignored: some
writers omit it or get it wrong.
"""

from __future__ import annotations

import math
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import IO

from sheetpeek._model import CellRange
from sheetpeek._package import PackageError

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_SHEET_DATA = f"{_NS}sheetData"
_ROW = f"{_NS}row"
_CELL = f"{_NS}c"
_FORMULA = f"{_NS}f"
_VALUE = f"{_NS}v"
_COL = f"{_NS}col"
_MERGE = f"{_NS}mergeCell"
_HYPERLINK = f"{_NS}hyperlink"
_SHEET_VIEW = f"{_NS}sheetView"
_SHEET_FORMAT = f"{_NS}sheetFormatPr"
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
    hyperlinks: tuple[tuple[CellRange, str], ...]
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


def scan(source: IO[bytes], rels: dict[str, tuple[str, bool]] | None = None) -> Structure:
    """Scan one worksheet part.

    Args:
        source: The worksheet part, opened for reading.
        rels: The part's relationships (id → target, external), for hyperlink targets.
    """
    rels = rels or {}
    min_row = min_col = 1 << 30
    max_row = max_col = 0
    row_index = 0
    rows: dict[int, tuple[float | None, bool]] = {}
    columns: list[ColumnSpec] = []
    merges: list[CellRange] = []
    hyperlinks: list[tuple[CellRange, str]] = []
    uncached = 0
    first_uncached: str | None = None
    empty_strings: set[tuple[int, int]] = set()
    default_column_width: float | None = None
    base_column_width = 8
    default_row_height: float | None = None
    show_gridlines = True
    listed: set[int] = set()
    sheet_format: dict[str, str] | None = None
    seen_view = False
    sheet_data: ET.Element | None = None

    for event, element in ET.iterparse(source, events=("start", "end")):
        tag = element.tag
        if tag in (_CELL, _VALUE):
            continue  # cells are handled at the end of their row
        if event == "start":
            if tag == _SHEET_DATA:
                sheet_data = element
            continue
        if tag == _ROW:
            r = element.get("r")
            row_index = int(r) if r else row_index + 1
            listed.add(row_index)
            if len(element):
                first_col, last_col = _column_extent(element)
                min_row = min(min_row, row_index)
                max_row = max(max_row, row_index)
                min_col = min(min_col, first_col)
                max_col = max(max_col, last_col)
                for cell in element:
                    # The common cell is one <v>; only others can be formulas without a value.
                    if (len(cell) != 1 or cell[0].tag == _FORMULA) and (
                        cell.find(_FORMULA) is not None and not cell.findtext(_VALUE)
                    ):
                        ref = cell.get("r")
                        if cell.get("t") == "str":
                            # A formula whose cached result is the empty string. openpyxl reads
                            # it as an empty cell, but it is not empty: it stops spilled text.
                            if ref:
                                empty_strings.add(_position(ref))
                        else:
                            uncached += 1
                            if first_uncached is None:
                                first_uncached = ref or f"row {row_index}"
            height = element.get("ht")
            hidden = _flag(element.get("hidden"))
            if height is not None or hidden:
                rows[row_index] = (_row_height(height) if height is not None else None, hidden)
            if sheet_data is not None:
                sheet_data.clear()  # the row and its cells are done
        elif tag == _COL:
            width = element.get("width")
            columns.append(
                ColumnSpec(
                    int(element.get("min", "1")),
                    int(element.get("max", element.get("min", "1"))),
                    _column_width(width) if width is not None else None,
                    _flag(element.get("hidden")),
                )
            )
        elif tag == _MERGE:
            ref = element.get("ref")
            if ref:
                merges.append(CellRange.parse(ref))
        elif tag == _HYPERLINK:
            link = _hyperlink(element, rels)
            ref = element.get("ref")
            if link and ref:
                hyperlinks.append((CellRange.parse(ref), link))
        elif tag == _SHEET_VIEW and not seen_view:
            seen_view = True
            show_gridlines = _flag(element.get("showGridLines"), default=True)
        elif tag == _SHEET_FORMAT:
            sheet_format = dict(element.attrib)
            if (value := element.get("defaultColWidth")) is not None:
                default_column_width = _default_width(value)
            if (value := element.get("baseColWidth")) is not None:
                base_column_width = _base_width(value)
            if (value := element.get("defaultRowHeight")) is not None:
                default_row_height = _default_height(value)

    # A merged range occupies its cells even where the file has no <c> for them.
    for merge in merges:
        min_row, max_row = min(min_row, merge.min_row), max(max_row, merge.max_row)
        min_col, max_col = min(min_col, merge.min_col), max(max_col, merge.max_col)
    used = CellRange(min_row, min_col, max_row, max_col) if max_row else None
    custom_height = _flag((sheet_format or {}).get("customHeight"))
    zero_height = _flag((sheet_format or {}).get("zeroHeight"))
    if custom_height and default_row_height == 0.0:
        # An enormous default acts as zeroHeight: a row with an element keeps its height (captured).
        custom_height, zero_height = False, True
    return Structure(
        used_range=used,
        merges=tuple(merges),
        columns=tuple(columns),
        rows=rows,
        default_column_width=default_column_width,
        base_column_width=base_column_width,
        default_row_height=default_row_height,
        show_gridlines=show_gridlines,
        hyperlinks=tuple(hyperlinks),
        uncached_formulas=uncached,
        first_uncached_formula=first_uncached,
        empty_strings=frozenset(empty_strings),
        listed_rows=frozenset(listed),
        has_sheet_format=sheet_format is not None,
        custom_row_height=custom_height,
        zero_row_height=zero_height,
        thick_top=_flag((sheet_format or {}).get("thickTop")),
        thick_bottom=_flag((sheet_format or {}).get("thickBottom")),
    )


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


def _flag(value: str | None, *, default: bool = False) -> bool:
    if value is None:
        return default
    return value in ("1", "true")


def _hyperlink(element: ET.Element, rels: dict[str, tuple[str, bool]]) -> str | None:
    target = ""
    rel_id = element.get(_REL_ID)
    if rel_id and rel_id in rels:
        target = rels[rel_id][0]
    location = element.get("location")
    if location:
        target = f"{target}#{location}"
    return target or None
