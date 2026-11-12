"""The source model: what a worksheet says, before any layout.

Everything here is immutable. Coordinates are 1-based, as in openpyxl and Excel. Colors are
``RRGGBB`` strings, or ``None`` for Excel's automatic color.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Literal, TypeAlias

from xlformat import ExcelError

from sheetpeek._fill_border_model import Borders, Fill
from sheetpeek._hyperlink import Hyperlink
from sheetpeek._outline_model import Outline

if TYPE_CHECKING:
    from sheetpeek._notices import Notice
    from sheetpeek._rendered import Mode, RenderedSheet, SpillEdge
    from sheetpeek._runs_model import TextSpan

CellKind: TypeAlias = Literal["empty", "number", "text", "bool", "error"]
HorizontalAlignment: TypeAlias = Literal[
    "general", "left", "center", "right", "fill", "justify", "distributed", "center_across"
]
VerticalAlignment: TypeAlias = Literal["top", "center", "bottom", "justify", "distributed"]
SheetState: TypeAlias = Literal["visible", "hidden", "very_hidden"]
VerticalTextPosition: TypeAlias = Literal["superscript", "subscript"]
ReadingOrder: TypeAlias = Literal[0, 1, 2]


_NOT_ON_A_LINE = re.compile("[\x00-\x1f\x7f-\x9f\ud800-\udfff]")


def one_line(text: str) -> str:
    """Text from a workbook, a file name or a client, safe to print on one line of a terminal.

    No control character, C0 or C1, which a terminal would take as a command, and no line break,
    which could forge a line of output. Lone surrogates go too, since they cannot be encoded.
    """
    return _NOT_ON_A_LINE.sub("", text)


CellValue: TypeAlias = bool | int | float | str | ExcelError | None
"""A cached cell value. Dates are serial numbers; an error is an :class:`xlformat.ExcelError`, a
``str`` enum that equals its text (compare errors with ``==``, not ``is``)."""

_REF = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")


@dataclass(frozen=True, slots=True, order=True)
class CellRange:
    """A rectangle of cells, given by its first and last row and column (inclusive)."""

    min_row: int
    min_col: int
    max_row: int
    max_col: int

    def __post_init__(self) -> None:
        if not (1 <= self.min_row <= self.max_row and 1 <= self.min_col <= self.max_col):
            raise ValueError(f"not a cell range: {self!r}")

    @classmethod
    def parse(cls, ref: str) -> CellRange:
        """Parse an A1-style reference such as ``"B2"`` or ``"A1:H40"``.

        Raises:
            ValueError: ``ref`` is not a cell or a range of cells.
        """
        first, sep, last = ref.strip().partition(":")
        start = _parse_cell(first, ref)
        end = _parse_cell(last, ref) if sep else start
        return cls(
            min(start[0], end[0]),
            min(start[1], end[1]),
            max(start[0], end[0]),
            max(start[1], end[1]),
        )

    def __str__(self) -> str:
        first = f"{column_letter(self.min_col)}{self.min_row}"
        if (self.min_row, self.min_col) == (self.max_row, self.max_col):
            return first
        return f"{first}:{column_letter(self.max_col)}{self.max_row}"

    @property
    def rows(self) -> int:
        """Number of rows."""
        return self.max_row - self.min_row + 1

    @property
    def columns(self) -> int:
        """Number of columns."""
        return self.max_col - self.min_col + 1

    def __contains__(self, position: object) -> bool:
        if not isinstance(position, tuple) or len(position) != 2:
            return False
        row, col = position
        return bool(self.min_row <= row <= self.max_row and self.min_col <= col <= self.max_col)

    def intersects(self, other: CellRange) -> bool:
        """Whether the two ranges share at least one cell."""
        return not (
            other.max_row < self.min_row
            or other.min_row > self.max_row
            or other.max_col < self.min_col
            or other.min_col > self.max_col
        )

    def intersection(self, other: CellRange) -> CellRange | None:
        """The cells the two ranges share, or ``None``."""
        if not self.intersects(other):
            return None
        return CellRange(
            max(self.min_row, other.min_row),
            max(self.min_col, other.min_col),
            min(self.max_row, other.max_row),
            min(self.max_col, other.max_col),
        )


def column_letter(index: int) -> str:
    """The letters of a 1-based column index: 1 is ``A``, 27 is ``AA``."""
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def column_index(letters: str) -> int:
    """The 1-based index of a column's letters, in either case."""
    index = 0
    for char in letters.upper():
        index = index * 26 + ord(char) - 64
    return index


def _parse_cell(text: str, ref: str) -> tuple[int, int]:
    match = _REF.match(text.strip())
    if match is None:
        raise ValueError(f"not a cell reference: {ref!r}")
    row = int(match[2])
    col = column_index(match[1])
    if not (1 <= row <= 1_048_576 and 1 <= col <= 16_384):
        raise ValueError(f"outside the sheet: {ref!r}")
    return row, col


@dataclass(frozen=True, slots=True)
class Font:
    """A cell's font. ``color`` is ``RRGGBB``, or ``None`` for automatic."""

    name: str
    size: float
    bold: bool = False
    italic: bool = False
    underline: str | None = None
    """``None``, or the file's underline style: ``single``, ``double``, ``singleAccounting``,
    ``doubleAccounting``."""
    strike: bool = False
    color: str | None = None
    vert_align: VerticalTextPosition | None = None
    """``superscript`` or ``subscript``: the whole cell's text drawn at two thirds of the size,
    raised or lowered (``ARCHITECTURE.md`` §4.13); ``None`` on the baseline."""
    scheme: Literal["major", "minor"] | None = None
    """The font scheme the file gives, ``major`` (headings) or ``minor`` (body) for a theme font,
    whose ``name`` is then the theme's typeface (§4.14); ``None`` for a font named for itself."""
    stored_name: str | None = None
    """The name the file stores, when Excel draws another (a theme font); else ``None``."""


@dataclass(frozen=True, slots=True)
class CellStyle:
    """Everything a cell's style index resolves to; shared by every cell with that style."""

    number_format: str
    font: Font
    fill: str | None = None
    fill_spec: Fill | None = None
    """The fill itself (:class:`Pattern` or :class:`Gradient`); ``fill`` is its one color."""
    borders: Borders | None = None
    """The cell's border sides and diagonals, as the file states them."""
    h_align: HorizontalAlignment = "general"
    v_align: VerticalAlignment = "bottom"
    wrap: bool = False
    indent: int = 0
    shrink: bool = False
    """Whether text too wide for the cell is drawn smaller, in steps (shrink to fit, §4.13)."""
    reading_order: ReadingOrder = 0
    """The file's ``readingOrder``, 0 from the text (context), 1 left to right, 2 right to left."""
    justify_last_line: bool = False
    """``justifyLastLine`` of distributed text; Excel draws no difference (§4.13)."""


@dataclass(frozen=True, slots=True)
class Cell:
    """One cell as the file states it: the cached value and the resolved style."""

    row: int
    column: int
    value: CellValue
    kind: CellKind
    style: CellStyle
    hyperlink: Hyperlink | None = None
    """The cell's link, from a ``<hyperlink>`` element or a ``HYPERLINK()`` formula."""
    runs: tuple[TextSpan, ...] | None = None
    """The text's rich-text runs, each in its own font, or ``None`` when the text has none or
    every run is in the cell's font; ``value`` holds their text joined (§4.16)."""

    @property
    def number_format(self) -> str:
        """The cell's number format string."""
        return self.style.number_format

    @property
    def font(self) -> Font:
        """The cell's font."""
        return self.style.font

    @property
    def fill(self) -> str | None:
        """Background color of a pattern fill, ``RRGGBB``, or ``None``."""
        return self.style.fill

    @property
    def coordinate(self) -> str:
        """The A1-style coordinate, such as ``"C7"``."""
        return f"{column_letter(self.column)}{self.row}"


@dataclass(frozen=True, slots=True)
class Column:
    """A column's width as stored in the file, and whether it is hidden.

    ``width`` is in the file's units (ECMA-376 ``<col width>``, padding included), or ``None``
    when the column has the sheet's default width.
    """

    index: int
    width: float | None = None
    hidden: bool = False
    outline_level: int = 0
    """The column's outline level, 0 (not grouped) to 7."""
    collapsed: bool = False
    """The file marks it as the summary column of a collapsed group."""

    @property
    def letter(self) -> str:
        """The column letter, such as ``"C"``."""
        return column_letter(self.index)


@dataclass(frozen=True, slots=True)
class Row:
    """A row's height in points, or ``None`` for the sheet's default, and whether it is hidden."""

    index: int
    height: float | None = None
    hidden: bool = False
    listed: bool = False
    """Whether the file has a ``<row>`` element for it, which decides its default height in some
    sheet formats (``ARCHITECTURE.md`` §4.7)."""
    outline_level: int = 0
    """The row's outline level, 0 (not grouped) to 7."""
    collapsed: bool = False
    """The file marks it as the summary row of a collapsed group."""


@dataclass(frozen=True, slots=True)
class Sheet:
    """A worksheet as the file states it, before any layout.

    ``cells`` is a dense row-major grid over ``ref``: ``cells[i][j]`` is the cell at row
    ``ref.min_row + i`` and column ``ref.min_col + j``. ``columns`` and ``rows`` cover ``ref``
    in the same order. An empty sheet has ``ref`` and ``used_range`` of ``None`` and no cells.
    """

    name: str
    state: SheetState
    ref: CellRange | None
    used_range: CellRange | None
    cells: tuple[tuple[Cell, ...], ...]
    columns: tuple[Column, ...]
    rows: tuple[Row, ...]
    merges: tuple[CellRange, ...]
    normal_font: Font
    date1904: bool = False
    default_column_width: float | None = None
    """``<sheetFormatPr defaultColWidth>`` in file units, or ``None`` when the file omits it."""
    base_column_width: int = 8
    """``<sheetFormatPr baseColWidth>``: the default width in digits, before padding."""
    default_row_height: float | None = None
    """``<sheetFormatPr defaultRowHeight>``; Excel uses it only with ``custom_row_height``."""
    has_sheet_format: bool = True
    """Whether the file has ``<sheetFormatPr>``; without it an unlisted row is 16 pt."""
    custom_row_height: bool = False
    """``<sheetFormatPr customHeight>``: rows without a height take ``default_row_height``."""
    zero_row_height: bool = False
    """``<sheetFormatPr zeroHeight>``: rows without a ``<row>`` element are hidden."""
    thick_top: bool = False
    """``<sheetFormatPr thickTop>``: unlisted rows are 1 pt taller."""
    thick_bottom: bool = False
    """``<sheetFormatPr thickBottom>``: unlisted rows are 1 pt taller."""
    show_gridlines: bool = True
    show_headings: bool = True
    """The file's ``showRowColHeaders``: whether Excel shows row and column headings."""
    grid_color: str | None = None
    """The gridlines' color, ``RRGGBB`` from the workbook palette (``<sheetView
    defaultGridColor="0" colorId>``), or ``None`` for Excel's automatic gray (§4.14)."""
    outline: Outline | None = None
    """The outline's settings when any row or column has an outline level, else ``None``."""
    notices: tuple[Notice, ...] = field(default=())
    """What the file uses that sheetpeek does not read yet (:class:`~sheetpeek.Notice`)."""
    max_cells: int | None = field(default=None, repr=False, compare=False)
    """The largest grid its layout may build (:class:`~sheetpeek.ReadLimits`), or ``None``."""

    @property
    def warnings(self) -> tuple[str, ...]:
        """The messages of the notices that were warnings before notices existed.

        Deprecated: use :attr:`notices`.
        """
        from sheetpeek._notices import LEGACY

        return tuple(n.message for n in self.notices if n.code in LEGACY)

    def cell(self, row: int, column: int) -> Cell:
        """The cell at a sheet position inside ``ref``.

        Raises:
            IndexError: the position is outside ``ref``.
        """
        if self.ref is None or (row, column) not in self.ref:
            raise IndexError(f"({row}, {column}) is outside {self.ref}")
        return self.cells[row - self.ref.min_row][column - self.ref.min_col]

    def layout(
        self,
        mode: Mode = "characters",
        *,
        show_hidden: bool = False,
        scale: float = 1.0,
        max_width: int | None = None,
        spill_edge: SpillEdge = "extend",
    ) -> RenderedSheet:
        """Lay the sheet out for one width mode: what each visible cell shows.

        Args:
            mode: ``"excel"`` for Excel's own geometry in points (HTML and other drawn
                output): column and row positions, and each text as a
                :class:`~sheetpeek.TextRun` with its baseline and clip window;
                ``"characters"`` for fixed-width output
                (terminals, plain text), where every column is a whole number of characters
                and values that do not fit show ``#``; ``"none"`` for output without column
                widths (Markdown, CSV), where every cell shows its full text.
            show_hidden: Keep hidden rows and columns.
            scale: Multiplies every column's width in characters (``characters`` mode only).
                A column is never wider than 255 characters, Excel's own limit.
            max_width: Caps every column's width in characters (``characters`` mode only).
            spill_edge: What text spilling past the edge of the range does. ``"extend"`` (the
                default, as Excel shows it) adds empty columns at the default width past the
                range, marked ``added``, as far as the spill reaches. Columns are added only past
                the used range, where the sheet is known to be empty; a range that cuts through
                the data clips the spill at its edge.

        Raises:
            ValueError: an unknown mode, a scale that is not positive, or ``scale`` or
                ``max_width`` in ``excel`` mode.
            LimitError: The rows times the columns, with those added for spilled text, exceed
                the ``max_cells`` the sheet was read with.
        """
        from sheetpeek._layout import layout

        return layout(
            self,
            mode,
            show_hidden=show_hidden,
            scale=scale,
            max_width=max_width,
            spill_edge=spill_edge,
        )

    def __iter__(self) -> Iterator[tuple[Cell, ...]]:
        return iter(self.cells)

    def __rich_console__(self, console: Any, options: Any) -> Iterator[Any]:
        """Draw the sheet with ``rich.print(sheet)``, laid out in ``characters`` mode."""
        from sheetpeek.backends.rich import RichBackend

        yield from RichBackend().lines(self.layout("characters"))
