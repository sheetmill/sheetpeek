"""Read a worksheet into a :class:`~sheetpeek.Sheet`.

Everything is read through xlio: the cells from one pass of its stream (``Book._stream``), which
also gives the used range and the marks the notices need, and the structure, styles and parts
from its models. Dates stay the serial numbers Excel stores. ``ARCHITECTURE.md`` §5.4 has the
measurements behind this.
"""

from __future__ import annotations

import contextlib
import dataclasses
import functools
import io
import os
import re
import zipfile
from array import array
from collections.abc import Iterator, Sequence
from types import MappingProxyType, TracebackType
from typing import IO, Any, overload

import xlformat
import xlio
import xlrecalc

from sheetpeek import _style as styles
from sheetpeek._colors import SYSTEM_BACKGROUND, grid_color, resolve, workbook_palette
from sheetpeek._dates import iso_serial
from sheetpeek._detect import (
    book_notices,
    report_cells,
    report_style,
    sheet_notices,
    visible_style,
    workbook_info,
)
from sheetpeek._hyperlink import Hyperlink
from sheetpeek._model import (
    Cell,
    CellKind,
    CellRange,
    CellStyle,
    CellValue,
    Column,
    Font,
    Row,
    Sheet,
    SheetState,
)
from sheetpeek._notices import CellSet, Collector, Notice
from sheetpeek._package import (
    LimitError,
    PackageError,
    ReadLimits,
    open_package,
    open_workbook,
    worksheet_parts,
)
from sheetpeek._page import MAX_COLUMN, MAX_ROW, manual_breaks, print_areas
from sheetpeek._read_filters import read_autofilters
from sheetpeek._read_links import Found as Links
from sheetpeek._read_links import add_formula, cell_links, with_elements, with_formulas
from sheetpeek._read_links import report as report_links
from sheetpeek._read_outline import read_outline, with_levels
from sheetpeek._read_tables import read_tables
from sheetpeek._read_view import sheet_view
from sheetpeek._recalc import KEPT, Recalc, RecalcOnLoad, recalc_mode, shown, why_not
from sheetpeek._runs_model import TextSpan
from sheetpeek._sheets import (
    ListedSheet,
    NotAWorksheetError,
    SheetInfo,
    SheetKind,
    StatedColor,
    default_sheet,
    from_xlio,
)
from sheetpeek._structure import XLIO_DIGIT, Structure, fill, width_at_digit
from sheetpeek._style import Code
from sheetpeek._style import font as font_aspect
from sheetpeek._view_model import SheetView

Source = str | os.PathLike[str] | IO[bytes]

_NOT_SHOWN = {
    "chartsheet": "is a chartsheet; charts are not shown",
    "dialogsheet": "is a dialog sheet (a legacy form); it is not shown",
}
_EXACT = 2**53  # whole numbers below this are exact as floats
_SURROGATE = re.compile("[\ud800-\udfff]")


def book_file_name(book: Book) -> str | None:
    """The file name a book was opened from, without its folder; ``None`` for a file object."""
    return os.path.basename(book._path) if book._path else None


def active_sheet(book: Book) -> SheetInfo | None:
    """The sheet Excel opens on (``activeTab``), if the workbook lists it; any kind or state."""
    active = book._listing[1]
    if active is None:
        return None
    return next((info for info in book.sheets if info.name == active.name), None)


def book_limits(book: Book) -> ReadLimits:
    """The limits a book reads under."""
    return book._limits


_HEX6 = re.compile(r"[0-9A-Fa-f]{6}")
_KIND: dict[str, SheetKind] = {
    "worksheet": "worksheet",
    "chartsheet": "chartsheet",
    "dialogsheet": "dialogsheet",
    "macrosheet": "macrosheet",
}
_STATE: dict[str, SheetState] = {
    "visible": "visible",
    "hidden": "hidden",
    "veryHidden": "very_hidden",
}


def xlio_theme_colors(theme: xlio.Theme | None) -> tuple[str, ...]:
    """The twelve theme colors as ``RRGGBB`` in Excel's index order, from xlio's reading.

    No theme, a color missing, or one that is not six hex digits gives no theme colors: the
    theme's text reaches CSS, so nothing but a color may leave here.
    """
    if theme is None or theme.colors is None:
        return ()
    found = []
    for index in range(12):
        color = theme.colors.by_index(index)
        rgb = (color.rgb or "") if color is not None else ""
        if not _HEX6.fullmatch(rgb):
            return ()
        found.append(rgb.upper())
    return tuple(found)


_NO_THEME_FONTS = {"minor": "Aptos Narrow"}
"""Without a theme part Excel draws minor-scheme text, the Normal style's included, in its own
default theme's minor font (``theme_fonts_palette``); the major font is not captured."""


def xlio_theme_fonts(theme: xlio.Theme | None) -> dict[str, str]:
    """The theme's Latin typefaces, ``{"major": ..., "minor": ...}``, those it names."""
    if theme is None:
        return dict(_NO_THEME_FONTS)
    pairs = (("major", theme.major_fonts), ("minor", theme.minor_fonts))
    return {key: fonts.latin for key, fonts in pairs if fonts is not None and fonts.latin}


class Book:
    """An open workbook, for reading several sheets with one load. Use :func:`open`.

    Use a book from one thread at a time: its readers share one open file.
    """

    def __init__(
        self,
        source: Source,
        *,
        limits: ReadLimits | None = None,
        recalc_on_load: RecalcOnLoad = "excel",
    ) -> None:
        self._limits = limits or ReadLimits()
        self._mode = recalc_mode(recalc_on_load)
        self._recalc_state: Recalc | None = None
        self._path: str | None = None
        name: str | None
        handle: IO[bytes]
        if isinstance(source, str | os.PathLike):
            # One open file for the check and for every read after it, so that what was checked
            # is what is read. A missing or unreadable path raises its own OSError here.
            self._path = name = os.fspath(source)
            handle = io.open(self._path, "rb")  # noqa: SIM115, UP020 - kept open; open() here is this module's
        else:
            given = getattr(source, "name", None)
            name = given if isinstance(given, str) else None
            handle = io.BytesIO(source.read())
        self._handle = handle
        self._book: xlio.Workbook | None = None
        self._archive: zipfile.ZipFile | None = None
        try:
            self._archive = open_package(handle, self._limits, name)
            self._book = open_workbook(handle, self._limits)
            self._parts = worksheet_parts(self._archive)
            self._date1904 = bool(self._book.properties.date1904)
            self._theme = xlio_theme_colors(self._book.theme)
            custom = self._book.styles.indexed_colors or None
            self._palette: Sequence[str] = workbook_palette(custom)
            self._own_palette = tuple(self._palette) if custom else None
            self._theme_fonts = xlio_theme_fonts(self._book.theme)
            self._normal_font = self._read_normal_font()
            self._normal_digit = _normal_digit(self._normal_font)
            self._info = workbook_info(self._book.parts)
            self._context = styles.Context(
                self._cell_color, self._normal_font, self._theme_fonts, self._palette
            )
        except PackageError:
            self.close()
            raise
        except xlio.LimitError as error:
            self.close()
            raise LimitError(str(error)) from None
        except Exception as error:  # a malformed part can raise more than xlio's own errors
            self.close()
            raise self._unreadable(error) from error
        except BaseException:
            self.close()
            raise
        self._default_style = CellStyle("General", self._normal_font)
        self._styles: dict[Any, tuple[CellStyle, tuple[Code, ...], tuple[Code, ...]]] = {}
        self._normal_color: dict[int, CellStyle] = {}
        """Styles whose font color is the Normal style's, by ``id`` (``Sheet._normal_color``)."""
        self._normal_key = (
            font_aspect.color_key(self._normal_font_part()) if self._book is not None else None
        )
        # A position without a <c> shows the Normal style, as a table's empty cells do.
        self._normal_color[id(self._default_style)] = self._default_style
        self._run_cache: dict[int, tuple[Any, Font, tuple[TextSpan, ...] | None]] = {}

    def _unreadable(self, error: Exception) -> PackageError:
        """The refusal for a failure to read: its reason, without where the file lives."""
        reason = str(error.__cause__ or error).split("\n")[0]
        if self._path:
            reason = reason.replace(self._path, os.path.basename(self._path))
        return PackageError(f"cannot read this workbook: {reason}")

    @contextlib.contextmanager
    def _xlio_errors(self) -> Iterator[None]:
        """Translate the refusals xlio raises while reading after opening into sheetpeek's."""
        try:
            yield
        except xlio.LimitError as error:
            raise LimitError(str(error)) from None
        except xlio.PackageError as error:
            raise self._unreadable(error) from error

    @property
    def _recalc(self) -> Recalc:
        """Recalculation on opening, set up on the first sheet read."""
        if self._recalc_state is None:
            assert self._book is not None
            self._recalc_state = Recalc(self._mode, self._book)
        return self._recalc_state

    def _check_area(self, ref: CellRange) -> None:
        area = (ref.max_row - ref.min_row + 1) * (ref.max_col - ref.min_col + 1)
        if area > self._limits.max_cells:
            raise LimitError(
                f"{ref} is {area} cells, more than the limit of {self._limits.max_cells}: "
                "read a smaller range"
            )

    def __enter__(self) -> Book:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        self.close()

    def close(self) -> None:
        """Close the file."""
        if self._book is not None:
            self._book.close()
        if self._archive is not None:
            self._archive.close()
        self._handle.close()

    @property
    def sheet_names(self) -> list[str]:
        """Names of all sheets, hidden ones included, in workbook order."""
        return [listed.name for listed, _ in self._listing[0]]

    @functools.cached_property
    def sheets(self) -> tuple[SheetInfo, ...]:
        """Every sheet, hidden ones and those sheetpeek does not draw included, in workbook order.

        Each with its kind, state and tab color. A chartsheet or dialog sheet is listed here, and
        :meth:`sheet` raises :class:`NotAWorksheetError` for it.
        """
        return tuple(
            SheetInfo(listed.name, listed.kind, listed.state, self._color(color))
            for listed, color in self._listing[0]
        )

    @functools.cached_property
    def _listing(self) -> tuple[list[tuple[ListedSheet, StatedColor | None]], ListedSheet | None]:
        """The workbook's sheets (the first of a name) with their tab colors, and the active one.

        A sheet whose part is missing is not listed (xlio). The active one is the first book
        view's ``activeTab``, which xlio gives as an index into its own list (a sheet listed
        without a part counts only as written).
        """
        assert self._book is not None
        listed = [
            (ListedSheet(s.name, _KIND.get(s.kind, "worksheet"), _STATE[s.state], None), s)
            for s in self._book
        ]
        kept: dict[str, tuple[ListedSheet, StatedColor | None]] = {}
        with self._xlio_errors():
            for sheet, source in listed:
                if sheet.name not in kept:
                    kept[sheet.name] = (
                        sheet,
                        from_xlio(source.properties.tab_color),
                    )
            views = self._book.views
        active = views[0].active_tab if views else None
        chosen = listed[active][0] if active is not None and 0 <= active < len(listed) else None
        found = any(chosen is sheet for sheet, _ in kept.values())
        return list(kept.values()), chosen if found else None

    @property
    def notices(self) -> tuple[Notice, ...]:
        """Workbook-level notices: the sheets that cannot be shown (chartsheets, dialog sheets)."""
        collector = Collector(None)
        book_notices(collector, self._info)
        return collector.notices()

    @property
    def date1904(self) -> bool:
        """Whether the workbook uses the 1904 date system."""
        return self._date1904

    def sheet(self, name: str | None = None, *, ref: str | CellRange | None = None) -> Sheet:
        """Read one sheet.

        Args:
            name: The sheet's name. By default, the sheet Excel opens on: the active sheet, or
                the first visible worksheet when the active one is hidden or not a worksheet.
            ref: The range to read, such as ``"A1:H40"``. By default, the used range.

        Raises:
            KeyError: There is no sheet with that name.
            NotAWorksheetError: The sheet is a chartsheet or a dialog sheet (a
                :class:`ValueError`).
            ValueError: ``ref`` is not a range.
            LimitError: ``ref``, or else the sheet's used range, is over ``max_cells``.
            PackageError: The sheet's part cannot be read.
        """
        return self._sheet(name, ref)

    def _print_area_sheets(self, name: str | None) -> tuple[tuple[Sheet, ...], str | None]:
        """A sheet's print areas, each read as a sheet of its own, in the file's order.

        An area of whole columns or whole rows is cut to the sheet's used range, as Excel
        prints it. A sheet without a print area gives the whole sheet, as :meth:`sheet` does.

        Returns:
            The sheets, and why a part of the print area was left out (``#REF!``, a formula,
            another sheet), else ``None``.
        """
        if name is None:
            name = self._default_sheet_name()
        areas, problem = self._page_layout(name)[:2]
        if not areas and problem is None:
            return (self._sheet(name, None),), None
        sheets = tuple(s for area in areas if (s := self._sheet(name, area, clip=True)) is not None)
        return sheets, problem

    def _view(self, name: str) -> SheetView:
        """The sheet view, from xlio."""
        assert self._book is not None
        with self._xlio_errors():
            return sheet_view(self._book.sheet(name).views)

    def _page_layout(
        self, name: str
    ) -> tuple[tuple[CellRange, ...], str | None, tuple[int, ...], tuple[int, ...]]:
        """The print area, its problem, and the manual row and column breaks, from xlio."""
        assert self._book is not None
        with self._xlio_errors():
            source = self._book.sheet(name)
            areas, problem = print_areas(source.print_area)
            rows = manual_breaks(source.row_breaks, MAX_ROW)
            columns = manual_breaks(source.column_breaks, MAX_COLUMN)
        return areas, problem, rows, columns

    @overload
    def _sheet(self, name: str | None, ref: str | CellRange | None) -> Sheet: ...
    @overload
    def _sheet(self, name: str | None, ref: CellRange, *, clip: bool) -> Sheet | None: ...
    def _sheet(
        self, name: str | None, ref: str | CellRange | None, *, clip: bool = False
    ) -> Sheet | None:
        """Read one sheet, as :meth:`sheet` does.

        With ``clip``, whole columns or rows of ``ref`` are cut to the used range first, and
        ``None`` is given when nothing of it is left.
        """
        if name is None:
            name = self._default_sheet_name()
        listed = _find_sheet([s for s, _ in self._listing[0]], name)
        if listed is None:
            raise KeyError(f"no sheet named {name!r}; the sheets are {self.sheet_names}")
        name = listed.name  # the file's own spelling
        kind = self._info.kinds.get(name, "worksheet")
        if kind in ("chartsheet", "dialogsheet"):
            raise NotAWorksheetError(f"{name!r} {_NOT_SHOWN[kind]}")
        if listed.kind not in ("worksheet", "macrosheet") or name not in self._parts:
            raise NotAWorksheetError(f"{name!r} is not a worksheet")
        if isinstance(ref, str):
            ref = CellRange.parse(ref)
        if ref is not None and not clip:
            self._check_area(ref)
        assert self._book is not None
        try:
            with self._xlio_errors():
                source = self._book.sheet(name)
                cell_notes = Collector(name)
                structure, kept, formulas = self._stream(
                    source, name if listed.kind == "worksheet" else None, ref, cell_notes
                )
                structure = with_levels(with_elements(fill(structure, source), source), source)
                structure = with_formulas(structure, structure.extra.pop(_LINKS))
        except PackageError:
            raise
        except Exception as error:
            raise self._unreadable(error) from error
        if ref is None:
            ref = structure.used_range
            if ref is not None:
                self._check_area(ref)
        elif clip:
            ref = _clip(ref, structure.used_range)
            if ref is None:
                return None
            self._check_area(ref)

        notes = Collector(name)
        try:
            with self._xlio_errors():
                sheet_notices(notes, self._book.parts, source, self._info, name, ref)
                autofilters = read_autofilters(source, ref)
                tables, table_styles = read_tables(
                    source, ref, self._book.styles.table_style, self._context
                )
                notes.extend(cell_notes.notices())  # after the sheet's own, as they were read
                empty = self._recompute(name, ref, kept, formulas, notes) if ref else CellSet()
                cells = self._grid(kept, ref, structure, notes) if ref else ()
        except PackageError:
            raise
        except Exception as error:
            raise self._unreadable(error) from error
        self._report_structure(notes, structure, ref, empty)
        areas, area_problem, row_breaks, column_breaks = self._page_layout(name)
        columns = _columns(structure, ref, self._normal_digit)
        rows, columns, outline = read_outline(structure, ref, _rows(structure, ref), columns, notes)
        return Sheet(
            name=name,
            state=listed.state,
            ref=ref,
            used_range=structure.used_range,
            cells=cells,
            columns=columns,
            rows=rows,
            merges=tuple(m for m in structure.merges if ref and m.intersects(ref)),
            normal_font=self._normal_font,
            date1904=self.date1904,
            default_column_width=structure.default_column_width,
            base_column_width=structure.base_column_width,
            default_row_height=structure.default_row_height,
            has_sheet_format=structure.has_sheet_format,
            custom_row_height=structure.custom_row_height,
            zero_row_height=structure.zero_row_height,
            thick_top=structure.thick_top,
            thick_bottom=structure.thick_bottom,
            show_gridlines=structure.show_gridlines,
            show_headings=structure.show_headings,
            grid_color=grid_color(structure.grid_color_id, self._palette),
            theme_colors=self._theme,
            palette=self._own_palette,
            outline=outline,
            view=self._view(name),
            print_areas=areas,
            print_area_problem=area_problem,
            row_breaks=row_breaks,
            column_breaks=column_breaks,
            autofilters=autofilters,
            tables=tables,
            table_styles=table_styles,
            notices=notes.notices(),
            max_cells=self._limits.max_cells,
            _normal_color=MappingProxyType(dict(self._normal_color)) if tables else {},
        )

    def _report_structure(
        self,
        notes: Collector,
        structure: Structure,
        ref: CellRange | None,
        empty: CellSet | None,
    ) -> None:
        """Notices from what the cell stream found: uncached formulas, runs, row styles.

        The formulas without a cached value reported are those that show as empty: ``empty``,
        the range's that xlrecalc did not compute, or every one of the sheet's when
        recalculation could not be done (``None``).
        """
        uncached = structure.uncached if empty is None else empty
        if uncached.count:
            ranges, truncated = uncached.ranges()
            top = min(ranges, key=lambda r: (r.min_row, r.min_col)) if ranges else None
            first = (
                structure.first_uncached_formula
                if empty is None or top is None
                else f"{CellRange(top.min_row, top.min_col, top.min_row, top.min_col)}"
            )
            notes.extend(
                [
                    Notice(
                        "value.no_cached_value",
                        "omitted",
                        "not_implemented",
                        f"{uncached.count} formula cell(s) have no cached value and "
                        f"show as empty (first: {first}); a file "
                        "saved by Excel stores the values",
                        notes.sheet,
                        ranges,
                        uncached.count,
                        context={"truncated": 1} if truncated else {},
                    )
                ]
            )
        if ref is None:
            return
        assert self._book is not None
        ranges, _ = structure.phonetic.ranges()
        report_cells(notes, "text.phonetic", (r for r in map(ref.intersection, ranges) if r))
        rows = sum(
            1
            for row, style in structure.row_styles
            if ref.min_row <= row <= ref.max_row and visible_style(self._book.styles, style)
        )
        columns = sum(
            max(0, min(last, ref.max_col) - max(first, ref.min_col) + 1)
            for first, last, style in structure.column_styles
            if visible_style(self._book.styles, style)
        )
        if rows or columns:
            notes.add(
                "style.row_column",
                "omitted",
                "not_implemented",
                count=rows + columns,
                detail="fills and borders of whole rows and columns, on cells the file has no "
                "element for",
                context={"rows": rows, "columns": columns},
            )

    def _default_sheet_name(self) -> str:
        found = default_sheet([sheet for sheet, _ in self._listing[0]], self._listing[1])
        if found is None:
            raise NotAWorksheetError("the workbook has no worksheet to show")
        return found.name

    def _stream(
        self, sheet: xlio.Sheet, name: str | None, ref: CellRange | None, notes: Collector
    ) -> tuple[Structure, dict[int, list[Cell]], _Formulas]:
        """One pass of xlio's cell stream: the structure's marks, and the cells, built as read.

        The used range is the bounds of every cell the stream gives (the merges are added by
        :func:`fill`); the rows with cells are listed. Also found: formulas without a cached
        value, strings with runs, shown phonetic guides, ``HYPERLINK()`` formulas. The cells
        inside ``ref`` are built, every cell without one (they all lie inside the used range),
        each kept in its row's list in file order with no hyperlink yet (:meth:`_grid` adds
        them); what their styles use and is not read goes to ``notes``. No ``xlio.Cell`` is
        kept, and a row's cells share one row number. A worksheet's cells pass through the
        recalculation's loader; the formula cells inside ``ref`` and the ranges of array
        formulas are kept for :meth:`_recompute`. ``name`` is ``None`` for a macro sheet, which
        xlrecalc does not compute: its cells keep their cached values.
        """
        recalc = self._recalc
        formulas = _Formulas()
        min_row = min_col = 1 << 30
        max_row = max_col = 0
        listed: set[int] = set()
        uncached, runs, phonetic = CellSet(), CellSet(), CellSet()
        first_uncached: str | None = None
        links = Links()
        rows: dict[int, list[Cell]] = {}
        line: list[Cell] = []
        current = 0
        stream = recalc.cells(sheet) if name is not None else sheet.cells(rich=True)
        for source in stream:
            row, column = source.row, source.column
            if row == current:
                row = current  # one int per row: each cell's own would cost 32 bytes a cell
            else:
                current = row
                if row not in listed:
                    listed.add(row)
                    min_row = min(min_row, row)
                    max_row = max(max_row, row)
                line = rows.setdefault(row, [])
            if column < min_col:
                min_col = column
            if column > max_col:
                max_col = column
            formula = source.formula
            if formula is not None:
                cached = source.type != "none"
                if not cached:
                    uncached.add(row, column)
                    if first_uncached is None:
                        first_uncached = f"{CellRange(row, column, row, column)}"
                if name is not None and formula.kind in ("array", "dataTable") and formula.ref:
                    formulas.add_array(row, column, formula.ref, ref)
                add_formula(links, formula, row, column)
            rich = source.rich
            if rich is not None:
                if rich.runs:
                    runs.add(row, column)
                if rich.phonetic and source.phonetic and source.type == "s":
                    phonetic.add(row, column)
            if ref is not None and not (
                ref.min_row <= row <= ref.max_row and ref.min_col <= column <= ref.max_col
            ):
                continue
            value, kind = self._value(source)
            style, every, text = self._style(source.style)
            for code in every:
                report_style(notes, code, row, column)
            if text and value is not None:
                for code in text:
                    report_style(notes, code, row, column)
            if formula is not None and name is not None:
                formulas.add(row, len(line))
            line.append(
                Cell(
                    row,
                    column,
                    value,
                    kind,
                    style,
                    None,
                    self._runs(rich, style.font, source.value) if kind == "text" else None,
                )
            )
        used = CellRange(min_row, min_col, max_row, max_col) if max_row else None
        listed.update(sheet.rows)
        structure = Structure(
            used_range=used,
            uncached_formulas=uncached.count,
            first_uncached_formula=first_uncached,
            listed_rows=frozenset(listed),
            uncached=uncached,
            runs=runs,
            phonetic=phonetic,
            extra={_LINKS: links},
        )
        return structure, rows, formulas

    def _recompute(
        self,
        name: str,
        ref: CellRange,
        rows: dict[int, list[Cell]],
        formulas: _Formulas,
        notes: Collector,
    ) -> CellSet | None:
        """The formula cells of ``ref`` as Excel shows them on opening, in ``rows`` in place.

        Nothing changes when the check's answer is ``"none"``. Otherwise the engine gives each
        formula cell's value, and each cell an array formula fills; a cell xlrecalc does not
        compute keeps the value the file holds and is reported, with xlrecalc's reason.

        Returns:
            The formula cells of ``ref`` without a cached value that xlrecalc did not compute;
            ``None`` when recalculation could not be done (xlrecalc refused the book), and every
            formula cell of ``ref`` keeps its cached value and is reported.
        """
        empty = CellSet()
        if not formulas.marks and not formulas.arrays:
            return empty
        recalc = self._recalc
        answer = recalc.answer()
        if answer == "none":
            return empty
        engine = recalc.engine() if answer is not None else None
        if engine is None:
            detail = recalc.error or "xlrecalc could not read the workbook"
            for row, indexes in formulas.marks.items():
                for index in indexes:
                    cell = rows[row][index]
                    notes.add(
                        "value.not_computed",
                        "approximated",
                        "not_implemented",
                        row=row,
                        column=cell.column,
                        detail=detail,
                    )
            return None

        def value(row: int, column: int, formula: bool) -> tuple[CellValue, CellKind] | None:
            """The engine's value, or ``None`` where the file's value stays (and is reported)."""
            where = xlrecalc.CellRef(name, row, column)
            try:
                result = shown(engine.value(where))
                status = engine.status(where)  # a filled cell's is its array's anchor's
            except Exception as error:  # a cell xlrecalc fails on keeps its value
                reason = f"an error inside xlrecalc ({type(error).__name__})"
            else:
                if status.state not in KEPT:
                    if result is not None:
                        return result
                    reason = "a value sheetpeek does not show"
                else:
                    reason = why_not(status)
            if formula:
                notes.add(
                    "value.not_computed",
                    "approximated",
                    "not_implemented",
                    row=row,
                    column=column,
                    detail=reason,
                )
            return None

        for row, indexes in formulas.marks.items():
            line = rows[row]
            for index in indexes:
                cell = line[index]
                found = value(row, cell.column, True)
                if found is None and cell.kind == "empty":
                    empty.add(row, cell.column)
                elif found is not None and found != (cell.value, cell.kind):
                    line[index] = dataclasses.replace(
                        cell, value=found[0], kind=found[1], runs=None
                    )
        # The cells an array fills: an array xlrecalc does not compute keeps their values.
        for anchor_row, anchor_column, filled in formulas.arrays:
            for row in range(filled.min_row, filled.max_row + 1):
                line = rows.setdefault(row, [])
                at = {cell.column: index for index, cell in enumerate(line)}
                for column in range(filled.min_col, filled.max_col + 1):
                    if (row, column) == (anchor_row, anchor_column):
                        continue
                    found = value(row, column, False)
                    if found is None:
                        continue
                    found_at = at.get(column)
                    if found_at is None:
                        if found[1] != "empty":
                            line.append(Cell(row, column, found[0], found[1], self._default_style))
                    elif found != (line[found_at].value, line[found_at].kind):
                        line[found_at] = dataclasses.replace(
                            line[found_at], value=found[0], kind=found[1], runs=None
                        )
        return empty

    def _grid(
        self,
        rows: dict[int, list[Cell]],
        ref: CellRange,
        structure: Structure,
        notes: Collector,
    ) -> tuple[tuple[Cell, ...], ...]:
        """The cells of ``ref`` from the stream's rows, with their links; others are empty.

        Each row's list is let go once its row is built. A cell the file repeats is the last one.
        """
        links = cell_links(structure, ref)
        first, width = ref.min_col, ref.max_col - ref.min_col + 1

        def line(index: int) -> tuple[Cell, ...]:
            placed: list[Cell | None] = [None] * width
            for cell in rows.pop(index, ()):
                if first <= cell.column <= ref.max_col:
                    if links and (found := links.get((cell.row, cell.column))) is not None:
                        cell = dataclasses.replace(cell, hyperlink=found)
                    placed[cell.column - first] = cell
            return tuple(
                cell or self._empty(index, first + offset, links)
                for offset, cell in enumerate(placed)
            )

        grid = tuple(line(index) for index in range(ref.min_row, ref.max_row + 1))
        report_links(notes, structure, ref, grid)
        return grid

    def _empty(
        self, row: int, column: int, links: dict[tuple[int, int], Hyperlink] | None = None
    ) -> Cell:
        """A position without a ``<c>``: empty, in the default style, with its link if any."""
        link = links.get((row, column)) if links else None
        return Cell(row, column, None, "empty", self._default_style, link)

    def _value(self, source: xlio.Cell) -> tuple[CellValue, CellKind]:
        """A cell's value and kind as Excel reads it (xlio), ``_xHHHH_`` escapes decoded."""
        value = source.value
        if value is None:
            return None, "empty"
        if source.type == "e" or isinstance(value, xlio.ErrorCode):
            try:
                return xlformat.ExcelError(str(value)), "error"
            except ValueError:  # not shaped like an error code: a malformed file
                return str(value), "text"
        if isinstance(value, bool):
            return value, "bool"
        if isinstance(value, float):
            # A whole number is an int, as the file's digits gave it before the move to xlio.
            if value.is_integer() and -_EXACT < value < _EXACT:
                return int(value), "number"
            return value, "number"
        if source.type == "d":  # ISO 8601 text: Excel reads it as a date serial
            return iso_serial(value, date1904=self._date1904), "number"
        return whole_characters(value), "text"

    def _runs(
        self, rich: xlio.RichString | None, font: Font, value: Any
    ) -> tuple[TextSpan, ...] | None:
        """A string's runs in their fonts; ``None`` when every run is in ``font``.

        A run without properties takes the cell's font. A run with properties stands alone: what
        they leave out is the default, not the cell's, and a missing name or size the
        Normal font's (both unverified). Adjacent runs in one font are joined.
        """
        if rich is None or not rich.runs:
            return None
        if "".join(run.text for run in rich.runs) != value:
            return None  # a string with text of its own outside its runs shows that text alone
        cached = self._run_cache.get(id(rich))
        if cached is not None and cached[0] is rich and cached[1] == font:
            return cached[2]
        spans: list[TextSpan] = []
        for run in rich.runs:
            own = (
                font
                if run.font is None
                else font_aspect.to_font(
                    run.font, self._normal_font, self._cell_color, self._theme_fonts
                )
            )
            if not run.text:
                continue
            text = whole_characters(run.text)
            if spans and _drawn(spans[-1].font) == _drawn(own):
                spans[-1] = TextSpan(spans[-1].text + text, own)
            else:
                spans.append(TextSpan(text, own))
        runs = tuple(spans) if any(_drawn(span.font) != _drawn(font) for span in spans) else None
        self._run_cache[id(rich)] = (rich, font, runs)  # the string keeps its id from reuse
        return runs

    def _style(self, index: int) -> tuple[CellStyle, tuple[Code, ...], tuple[Code, ...]]:
        """Cell format ``index``'s style, and what it uses that is not read (any cell, values)."""
        found = self._styles.get(index)
        if found is None:
            assert self._book is not None
            # An index past the cell formats raises IndexError: the sheet is unreadable.
            cell_format = self._book.styles.cell_formats[index]
            code = cell_format.number_format
            if code is None:
                code = _builtin_format(cell_format.number_format_id)
            cell_format = dataclasses.replace(cell_format, number_format=code)
            found = self._styles[index] = styles.resolve(cell_format, self._context)
            # By identity: a style whose font color is the Normal style's and one with another
            # color can compare equal (colors stated differently resolve alike). Interning
            # equal styles would mark both; test_table_layout's equal-style test fails if so.
            if font_aspect.color_key(cell_format.font) == self._normal_key:
                self._normal_color[id(found[0])] = found[0]
        return found

    def _cell_color(self, color: Any) -> str | None:
        """An ``xlio.Color`` as ``RRGGBB`` or ``None``.

        A color element without attributes is read as ``00000000``, black; a fill's background
        in it is white (``_style.fill.color``). Indexed 65, the system background, is white in a
        font, a border and a fill; 64, the system foreground, is automatic (captured,
        ``auto_colors``).
        """
        stated = from_xlio(color)
        if stated is None and color is not None:
            stated = StatedColor("rgb", rgb="00000000", tint=color.tint)
        if stated is not None and stated.type == "indexed" and stated.indexed == SYSTEM_BACKGROUND:
            return "FFFFFF"
        return self._color(stated)

    def _color(self, color: Any) -> str | None:
        """A :class:`StatedColor` (or an object of its shape) as ``RRGGBB`` or ``None``."""
        if color is None:
            return None
        kind = getattr(color, "type", None)
        # Only the field named by ``type`` is meaningful; the others are not read.
        value = getattr(color, kind, None) if kind in ("rgb", "indexed", "theme") else None
        return resolve(kind, value, getattr(color, "tint", 0.0) or 0.0, self._theme, self._palette)

    def _read_normal_font(self) -> Font:
        """The Normal style's font; else the first font; else Calibri 11."""
        font = self._normal_font_part()
        if font is None:
            return font_aspect.FALLBACK
        return font_aspect.to_font(font, font_aspect.FALLBACK, self._cell_color, self._theme_fonts)

    def _normal_font_part(self) -> Any:
        """The Normal style's ``xlio.Font``; else the first font; else ``None``."""
        assert self._book is not None
        found = self._book.styles
        font = next(
            (
                style.format.font
                for style in found.cell_styles
                if style.name == "Normal" and style.format is not None
            ),
            None,
        )
        if font is None and found.fonts:
            font = found.fonts[0]
        return font


def whole_characters(text: str) -> str:
    """``text`` with a lone UTF-16 surrogate as the replacement character.

    xlio joins an escaped surrogate pair (``_xD83D__xDE00_``) into the character it spells and
    keeps a lone one (``_xD800_``), which text cannot be encoded for output with.
    """
    if _SURROGATE.search(text):
        return text.encode("utf-16", "surrogatepass").decode("utf-16", "replace")
    return text


def _drawn(font: Font) -> Font:
    """A font as drawn, black for automatic.

    What the file stored for a theme font (scheme, name) and the theme color a color comes from
    do not count.
    """
    return dataclasses.replace(
        font, scheme=None, stored_name=None, theme_color=None, color=font.color or "000000"
    )


def open(  # noqa: A001 - the module-level open() of gzip, tarfile, shelve
    source: Source,
    *,
    limits: ReadLimits | None = None,
    recalc_on_load: RecalcOnLoad = "excel",
) -> Book:
    """Open a workbook to read several sheets with one load.

    Args:
        source: A path, or a binary file object (read into memory).
        limits: What reading may cost. By default, ``ReadLimits()``; lower them for files from
            people you do not trust.
        recalc_on_load: Which formulas show a computed value rather than the one the file
            holds, as xlrecalc's option of that name: ``"excel"``, those Excel recalculates on
            opening the file (all of them in a file another program wrote or an older Excel
            saved; in a file Excel saved, those it marked, its volatile cells and their
            dependents); ``"always"``, every formula; ``"never"``, none but those without a
            cached value and the volatile ones.

    Raises:
        OSError: A path that cannot be opened.
        ValueError: ``recalc_on_load`` is not one of the three.
        LimitError: The file exceeds ``limits``.
        PackageError: The file is not a readable xlsx workbook, or it is refused.

    Example:
        ``with sheetpeek.open("book.xlsx") as book: sheet = book.sheet("Data")``
    """
    return Book(source, limits=limits, recalc_on_load=recalc_on_load)


def read(
    source: Source,
    sheet: str | None = None,
    *,
    ref: str | CellRange | None = None,
    limits: ReadLimits | None = None,
    recalc_on_load: RecalcOnLoad = "excel",
) -> Sheet:
    """Read one sheet of a workbook.

    Args:
        source: A path, or a binary file object (read into memory).
        sheet: The sheet's name. By default, the sheet Excel opens on.
        ref: The range to read, such as ``"A1:H40"``. By default, the used range.
        limits: What reading may cost. By default, ``ReadLimits()``; lower them for files from
            people you do not trust.
        recalc_on_load: Which formulas show a computed value, as :func:`open` takes it.

    Raises:
        OSError: A path that cannot be opened.
        KeyError: There is no sheet with that name.
        NotAWorksheetError: The sheet is a chartsheet or a dialog sheet (a :class:`ValueError`).
        ValueError: ``ref`` is not a range, or ``recalc_on_load`` not one of the three.
        LimitError: The file, or the range, exceeds ``limits``.
        PackageError: The file is not a readable xlsx workbook, or it is refused.
    """
    with Book(source, limits=limits, recalc_on_load=recalc_on_load) as book:
        return book.sheet(sheet, ref=ref)


def _clip(area: CellRange, used: CellRange | None) -> CellRange | None:
    """A print area of whole columns or rows cut to the used range, ``None`` if nothing is left."""
    max_row, max_col = area.max_row, area.max_col
    if area.min_row == 1 and max_row == MAX_ROW:
        max_row = min(max_row, used.max_row) if used is not None else 0
    if area.min_col == 1 and max_col == MAX_COLUMN:
        max_col = min(max_col, used.max_col) if used is not None else 0
    if max_row < area.min_row or max_col < area.min_col:
        return None
    return CellRange(area.min_row, area.min_col, max_row, max_col)


def _normal_digit(font: Font) -> float:
    """The Normal font's digit in points, as the layout measures it.

    A face xlformat has not measured keeps xlio's 6 pt (the layout measures it as Calibri, §4.7).
    """
    if font.name.lower() not in {name.lower() for name in xlformat.measured_fonts()}:
        return XLIO_DIGIT
    measured = xlformat.Font(font.name, font.size, font.bold, font.italic)
    return xlformat.file_column_metrics(None, normal_font=measured).digit_width


def _columns(structure: Structure, ref: CellRange | None, digit: float) -> tuple[Column, ...]:
    if ref is None:
        return ()
    found: dict[int, Column] = {}
    for spec in structure.columns:
        width = width_at_digit(spec, digit)
        for index in range(max(spec.min, ref.min_col), min(spec.max, ref.max_col) + 1):
            found[index] = Column(index, width, spec.hidden)
    # A default width of 0 hides every column without a width of its own (captured).
    width = 0.0 if structure.default_column_width == 0 else None
    return tuple(found.get(i, Column(i, width)) for i in range(ref.min_col, ref.max_col + 1))


def _rows(structure: Structure, ref: CellRange | None) -> tuple[Row, ...]:
    if ref is None:
        return ()
    result = []
    for index in range(ref.min_row, ref.max_row + 1):
        height, hidden = structure.rows.get(index, (None, False))
        result.append(Row(index, height, hidden, index in structure.listed_rows))
    return tuple(result)


def _builtin_format(number_format_id: int) -> str:
    """The code of a built-in format id, as Excel shows it (xlformat's captured table).

    A custom id the file does not define reads as General (captured, suite ``odd_values``).
    """
    try:
        return xlformat.builtin_format(number_format_id)
    except ValueError:
        return "General"


_LINKS = "_stream_links"
"""Where :meth:`Book._stream` leaves the ``HYPERLINK()`` formulas, until the elements are read."""


def _find_sheet(sheets: list[ListedSheet], name: str) -> ListedSheet | None:
    """The sheet called ``name``, compared without case as Excel compares sheet names."""
    exact = next((s for s in sheets if s.name == name), None)
    if exact is not None:
        return exact
    folded = name.casefold()
    return next((s for s in sheets if s.name.casefold() == folded), None)


class _Formulas:
    """The formula cells a sheet read keeps for :meth:`Book._recompute`."""

    __slots__ = ("arrays", "marks")

    def __init__(self) -> None:
        self.marks: dict[int, array[int]] = {}
        """Row → the indexes, in that row's list of built cells, of its formula cells."""
        self.arrays: list[tuple[int, int, CellRange]] = []
        """Each array formula or data table that fills part of the range: its first cell and the
        part of its range inside the range read."""

    def add(self, row: int, index: int) -> None:
        marks = self.marks.get(row)
        if marks is None:
            self.marks[row] = marks = array("I")
        marks.append(index)

    def add_array(self, row: int, column: int, text: str, ref: CellRange | None) -> None:
        try:
            filled = CellRange.parse(text)
        except ValueError:
            return
        if ref is not None:
            inside = ref.intersection(filled)
            if inside is None:
                return
            filled = inside
        self.arrays.append((row, column, filled))
