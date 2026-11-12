"""The rendered model: what each visible cell shows, laid out for one width mode.

A backend draws a :class:`RenderedSheet` column by column. In ``characters`` mode every
:class:`RenderedCell` carries the exact characters drawn in its column (``lines``), which may be
its own text, part of a neighbor's spilled text, or part of a merged cell's text, and the
:class:`Gap` between two columns carries the characters drawn across it. The background of each
column is that cell's own ``fill``; the text drawn there takes the ``font`` and ``color`` of the
cell it comes from. This is how Excel paints spilled text over a neighbor's fill.
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Literal, TypeAlias

from sheetpeek._fill_border_model import RenderedBorders
from sheetpeek._model import Cell, CellRange, Font, Row, column_letter
from sheetpeek._outline_model import RenderedOutline
from sheetpeek._runs_model import TextSpan
from sheetpeek._text_model import TextDirection, TextRule

if TYPE_CHECKING:
    from sheetpeek._census import Presence
    from sheetpeek._model import Sheet
    from sheetpeek._notices import Notice

Mode: TypeAlias = Literal["excel", "characters", "none"]
SpillEdge: TypeAlias = Literal["extend"]
"""What spilled text does at the edge of the range read. Only ``"extend"`` exists so far; the
planned alternatives are listed in ``ARCHITECTURE.md`` ("Planned options")."""
Alignment: TypeAlias = Literal["left", "center", "right", "fill", "justify", "distributed"]


@dataclass(frozen=True, slots=True)
class RenderedColumn:
    """A visible column: its sheet index and its width, in characters or in points by mode."""

    index: int
    width: int | None
    """Width in characters in ``characters`` mode, else ``None``."""
    hidden: bool = False
    """Hidden in the sheet; present only when the layout was asked to show hidden columns."""
    added: bool = False
    """An empty column past the range read, added at the default width to show spilled text."""
    x_pt: float | None = None
    """In ``excel`` mode, the x of the column's left gridline in points, from the first visible
    column's; else ``None``."""
    width_pt: float | None = None
    """In ``excel`` mode, the width in points, gridline included; else ``None``."""

    @property
    def letter(self) -> str:
        """The column letter, such as ``"C"``."""
        return column_letter(self.index)


@dataclass(frozen=True, slots=True)
class RenderedCell:
    """One visible cell as Excel shows it.

    ``text`` is what the cell itself displays at its width (``####``, a reduced ``General``,
    a ``*`` fill), before any spill or clipping. ``lines`` is what is drawn in this cell's column:
    in ``characters`` mode one string of exactly the column's width per line of the row; in
    ``none`` mode a single line holding ``text`` (empty for a covered merged cell).
    """

    cell: Cell
    text: str
    full_text: str
    overflow: bool
    align: Alignment
    format_color: str | None
    """The number format's section color, ``RRGGBB``."""
    merge: CellRange | None
    """The merged range this cell belongs to, if any."""
    spill: tuple[int, int]
    """How many visible columns this cell's text reaches to the left and to the right."""
    lines: tuple[str, ...]
    source: tuple[int, int] | None
    """(row, column) of the cell whose text is drawn in ``lines``, or ``None`` when blank."""
    fill: str | None
    """The background of this column, which is the cell's own fill or the merge anchor's."""
    font: Font | None
    """Font of the text drawn here (the source cell's), or ``None`` when blank."""
    color: str | None
    """The color of the text drawn here, the source's section color or else its font color."""

    @property
    def row(self) -> int:
        """Sheet row."""
        return self.cell.row

    @property
    def column(self) -> int:
        """Sheet column."""
        return self.cell.column

    @property
    def is_merge_anchor(self) -> bool:
        """Whether this cell is the top-left cell of a merged range."""
        return self.merge is not None and (self.merge.min_row, self.merge.min_col) == (
            self.cell.row,
            self.cell.column,
        )


@dataclass(frozen=True, slots=True)
class Gap:
    """What is drawn between two adjacent visible columns in ``characters`` mode.

    ``lines`` holds, per line of the row, the characters of text that crosses the gap (spilled
    or merged text), or ``None`` where the backend draws its column separator.
    """

    lines: tuple[str | None, ...]
    fill: str | None
    """The background shared by both neighbors (or their merge), else ``None``."""
    font: Font | None
    color: str | None


@dataclass(frozen=True, slots=True)
class TextRun:
    """One line of a cell's text as Excel draws it, in ``excel`` mode.

    Every length is in points, in sheet coordinates: x to the right of the first visible
    column's left gridline, y down from the first visible row's top gridline. Glyph ``k`` starts
    at ``x`` plus the sum of the first ``k`` of ``xlformat.glyph_widths(text, font=...)``, since
    Excel for Mac sets glyphs on whole points.
    """

    text: str
    """The characters drawn, one line of the cell's text or a fill's whole copies."""
    x: float
    """The left edge of the first glyph's advance."""
    y: float
    """The baseline, provisional since it comes from Excel's print layout at Calibri 11."""
    width: float
    """The sum of Excel's glyph advances (``xlformat.text_width``)."""
    clip: tuple[float, float, float, float]
    """Left, top, right and bottom of what shows; text outside is cut, through a glyph."""
    align: Literal["left", "center", "right"]
    """The point that stays where Excel put it when a different font draws the text."""
    source: tuple[int, int]
    """(row, column) of the cell whose text it is."""
    font: Font
    color: str | None
    """The number format's section color, else the font color; ``None`` is automatic."""
    scale: float = 1.0
    """The drawn size over ``font.size``: 2/3 for superscript and subscript, times the step
    shrink to fit chose (§4.13). ``width`` and the glyph widths are measured at the drawn size."""
    rules: tuple[TextRule, ...] = ()
    """The underlines and strikethrough drawn with the run (§4.13)."""
    direction: TextDirection = "ltr"
    """``rtl`` for a right-to-left paragraph: text whose first strong character is
    right-to-left, unless the cell's reading order is left to right (§4.13)."""
    spans: tuple[TextSpan, ...] = ()
    """The run's text in rich-text runs, each with its own font, width, size and baseline shift,
    for a cell with :attr:`~sheetpeek.Cell.runs` (§4.16); empty for text in one font. With
    spans, ``scale`` is the step of shrink to fit alone; each span carries its drawn size."""


@dataclass(frozen=True, slots=True)
class RenderedRow:
    """A visible row: its cells, one per visible column, and the gaps between them."""

    index: int
    source: Row
    height: int
    """Lines of text the row needs; 1 except for wrapped text in ``characters`` mode."""
    cells: tuple[RenderedCell, ...]
    gaps: tuple[Gap, ...]
    """``len(cells) - 1`` gaps in ``characters`` mode; none in the other modes."""
    y_pt: float | None = None
    """In ``excel`` mode, the y of the row's top gridline in points, from the first visible
    row's; else ``None``."""
    height_pt: float | None = None
    """In ``excel`` mode, the height in points, gridline included; else ``None``."""
    runs: tuple[TextRun, ...] = ()
    """In ``excel`` mode, every text drawn from this row's cells (a merge's from its first
    visible row), including text spilling into neighbors; two may overlap, as in Excel."""
    crossed: tuple[bool, ...] = ()
    """In ``excel`` mode, per boundary between two visible cells, whether text or a merge
    crosses it, so that no vertical gridline is drawn there."""


@dataclass(frozen=True, slots=True)
class RenderedSheet:
    """A sheet laid out for one width mode. Build it with :meth:`Sheet.layout`."""

    sheet: Sheet
    mode: Mode
    columns: tuple[RenderedColumn, ...]
    rows: tuple[RenderedRow, ...]
    gap: int
    """Characters between two columns in ``characters`` mode; 0 in ``none`` mode."""
    notices: tuple[Notice, ...]
    """The sheet's notices on the cells laid out, followed by the layout's. What a particular
    output leaves out is added by :func:`sheetpeek.notices`."""
    outline: RenderedOutline | None = None
    """The outline symbols, when the sheet has an outline whose symbols the file shows."""
    borders: RenderedBorders | None = None
    """In ``excel`` mode, the borders resolved to one side per edge (§4.12), or ``None`` when no
    laid-out cell has one. Every output reads these, never two cells' own sides."""
    extra: dict[str, Any] = field(default_factory=dict, compare=False)
    """What features' own layout passes add, by feature (``Draft.extra``), for features that
    keep their result here rather than in a field of their own."""

    @property
    def warnings(self) -> tuple[str, ...]:
        """The messages of the notices that were warnings before notices existed.

        Deprecated: use :attr:`notices`.
        """
        from sheetpeek._notices import LEGACY

        return tuple(n.message for n in self.notices if n.code in LEGACY)

    @property
    def features(self) -> dict[str, Presence]:
        """The census the outputs report from, by feature code.

        Each entry says how many cells (or things) use the feature, and for cell-level
        features which cells.
        """
        from sheetpeek._census import census

        return census(self)

    def cell(self, row: int, column: int) -> RenderedCell:
        """The visible cell at a sheet position.

        Raises:
            IndexError: the position is outside the range or hidden.
        """
        for rendered_row in self.rows:
            if rendered_row.index == row:
                for rendered in rendered_row.cells:
                    if rendered.column == column:
                        return rendered
        raise IndexError(f"({row}, {column}) is not a visible cell")

    def __iter__(self) -> Iterator[RenderedRow]:
        return iter(self.rows)

    def __rich_console__(self, console: Any, options: Any) -> Iterator[Any]:
        """Draw with ``rich.print``; a sheet not in ``characters`` mode is laid out again."""
        from sheetpeek.backends.rich import RichBackend

        sheet = self if self.mode == "characters" else self.sheet.layout("characters")
        yield from RichBackend().lines(sheet)

    @property
    def width_pt(self) -> float:
        """Total width in points in ``excel`` mode; 0 in the other modes."""
        return sum(c.width_pt or 0 for c in self.columns)

    @property
    def width(self) -> int:
        """Total width in characters of a row in ``characters`` mode, gaps included."""
        widths = [c.width or 0 for c in self.columns]
        return sum(widths) + self.gap * max(len(widths) - 1, 0)
