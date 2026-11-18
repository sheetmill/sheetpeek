"""Tables and their styles: the table parts a sheet holds and the formats they draw (§4.20).

The names follow ECMA-376's, which xlio and openpyxl use too, and Excel's object model where it
has one: a :class:`Table` (Excel's ``ListObject``) with its :class:`TableStyleInfo`, and the
:class:`TableStyle` it names, each of whose elements (:class:`TableStyleElement`) carries a
:class:`DifferentialFormat`: a partial format, of which only what it sets applies.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, TypeAlias

from sheetpeek._fill_border_model import Fill, Side
from sheetpeek._model import CellRange

TableStyleElementType: TypeAlias = Literal[
    "wholeTable",
    "firstColumnStripe",
    "secondColumnStripe",
    "firstRowStripe",
    "secondRowStripe",
    "lastColumn",
    "firstColumn",
    "headerRow",
    "totalRow",
    "firstHeaderCell",
    "lastHeaderCell",
    "firstTotalCell",
    "lastTotalCell",
]
"""The elements that apply to a table, in the order that stacks them (ECMA-376 Part 1,
§18.8.41): a later element's property wins over an earlier one's."""

ELEMENT_ORDER: tuple[TableStyleElementType, ...] = TableStyleElementType.__args__  # type: ignore[attr-defined]


@dataclass(frozen=True, slots=True)
class TableStyleInfo:
    """The style a table names and the four options that turn its elements on."""

    name: str | None = None
    """The style's name, a preset's (``TableStyleMedium2``) or the workbook's own; ``None`` draws
    no style."""
    show_first_column: bool = False
    show_last_column: bool = False
    show_row_stripes: bool = False
    show_column_stripes: bool = False


@dataclass(frozen=True, slots=True)
class Table:
    """A table on the sheet (``<tablePart>``)."""

    name: str
    """The table's display name, as formulas refer to it."""
    ref: CellRange
    """The whole table, header and totals rows included."""
    header_row_count: int = 1
    """1 with a header row, 0 without."""
    totals_row_count: int = 0
    """1 with a totals row, 0 without."""
    style: TableStyleInfo | None = None
    """The style the table draws; ``None`` (no ``tableStyleInfo``) draws none."""

    @property
    def data_rows(self) -> tuple[int, int]:
        """The first and last data row, the range less its header and totals rows."""
        return self.ref.min_row + self.header_row_count, self.ref.max_row - self.totals_row_count


@dataclass(frozen=True, slots=True)
class DifferentialFormat:
    """A partial format (ECMA-376's differential formatting, ``<dxf>``).

    Each field is ``None`` where the format leaves the property to what is beneath it. Colors
    are ``RRGGBB``, resolved through the workbook's theme. A solid fill is drawn in the format's
    ``bgColor``, unlike a cell's own.
    """

    bold: bool | None = None
    italic: bool | None = None
    underline: str | None = None
    """``single``, ``double``, ``singleAccounting``, ``doubleAccounting`` or ``none``."""
    strike: bool | None = None
    color: str | None = None
    """The font color."""
    fill: Fill | None = None
    left: Side | None = None
    right: Side | None = None
    top: Side | None = None
    bottom: Side | None = None
    horizontal: Side | None = None
    """The lines between the rows a format covers, inside the region it applies to."""
    vertical: Side | None = None
    """The lines between its columns."""


@dataclass(frozen=True, slots=True)
class TableStyleElement:
    """One element of a table style: where it applies and what it draws there."""

    type: TableStyleElementType
    format: DifferentialFormat
    size: int = 1
    """For a stripe, the rows or columns in each band (1 to 9)."""


@dataclass(frozen=True, slots=True)
class TableStyle:
    """A table style, a preset or the workbook's own (``<tableStyle>``)."""

    name: str
    elements: tuple[TableStyleElement, ...] = ()

    def element(self, type: TableStyleElementType) -> TableStyleElement | None:  # noqa: A002
        """The element of that type, or ``None``."""
        for found in self.elements:
            if found.type == type:
                return found
        return None
