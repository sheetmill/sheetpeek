"""Read a worksheet into a :class:`~sheetpeek.Sheet`.

Cells come from openpyxl's read-only mode, which streams the sheet in constant memory, with its
date conversion switched off so that dates stay the serial numbers Excel stores. What read-only
mode does not expose (merges, column widths, row heights, the used range) comes from the
structure scan in ``_structure.py``. ``ARCHITECTURE.md`` §5 has the measurements behind this.
"""

from __future__ import annotations

import contextlib
import dataclasses
import datetime as dt
import functools
import io
import os
import re
import warnings
import zipfile
from collections.abc import Iterator, Sequence
from contextlib import contextmanager
from types import TracebackType
from typing import IO, Any, overload

import xlformat
import xlio
from openpyxl import load_workbook
from openpyxl.styles.numbers import BUILTIN_FORMATS
from openpyxl.utils.datetime import CALENDAR_MAC_1904, to_excel

from sheetpeek import _style as styles
from sheetpeek._colors import grid_color, resolve, workbook_palette
from sheetpeek._detect import (
    book_notices,
    report_cells,
    report_style,
    sheet_notices,
    string_marks,
    visible_style,
    workbook_info,
)
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
from sheetpeek._notices import Collector, Notice
from sheetpeek._package import (
    LimitError,
    PackageError,
    ReadLimits,
    open_package,
    open_workbook,
    worksheet_parts,
)
from sheetpeek._page import MAX_COLUMN, MAX_ROW, manual_breaks, print_areas
from sheetpeek._runs_model import TextSpan
from sheetpeek._scan_links import cell_links, with_elements
from sheetpeek._scan_links import report as report_links
from sheetpeek._scan_outline import read_outline
from sheetpeek._scan_view import sheet_view
from sheetpeek._sheets import (
    ListedSheet,
    NotAWorksheetError,
    SheetInfo,
    SheetKind,
    StatedColor,
    default_sheet,
    from_xlio,
)
from sheetpeek._structure import Structure, fill, scan
from sheetpeek._style import Code
from sheetpeek._style import font as font_aspect
from sheetpeek._view_model import SheetView

Source = str | os.PathLike[str] | IO[bytes]

DATE_FALLBACK_WARNING = (
    "dates were read through openpyxl's date conversion, because this openpyxl version does not "
    "allow switching it off: times may be off by up to a millisecond, and serial 60 "
    "(1900-02-29) shows as 1900-02-28"
)
_STATES: dict[str, SheetState] = {
    "visible": "visible",
    "hidden": "hidden",
    "veryHidden": "very_hidden",
}
_NOT_SHOWN = {
    "chartsheet": "is a chartsheet; charts are not shown",
    "dialogsheet": "is a dialog sheet (a legacy form); it is not shown",
}
_DATE_TYPES = (dt.datetime, dt.date, dt.time, dt.timedelta)
_ESCAPE = re.compile(r"_x([0-9A-Fa-f]{4})_")
_SURROGATE = re.compile("[\ud800-\udfff]")


def unescape(text: str) -> str:
    """Decode OOXML's ``_xHHHH_`` escapes, as Excel does; openpyxl leaves them (§5.6).

    ``_x000D_`` is a carriage return and ``_x005F_x0041_`` the text ``_x0041_``: the first escape
    is consumed whole, so the rest is not read as another. Escaped UTF-16 surrogates become the
    character they spell, and a lone one the replacement character, since text holding one
    cannot be encoded for output.
    """
    if "_x" not in text:
        return text
    decoded = _ESCAPE.sub(lambda m: chr(int(m[1], 16)), text)
    if _SURROGATE.search(decoded):
        return decoded.encode("utf-16", "surrogatepass").decode("utf-16", "replace")
    return decoded


def disable_date_conversion(workbook: Any) -> bool:
    """Make a read-only workbook yield date cells as serial numbers; report whether it worked.

    openpyxl converts a number to a date while parsing a sheet when the cell's style is in
    ``workbook._date_formats``. Read-only mode parses lazily, so emptying that set (and
    ``_timedelta_formats``) before iterating leaves every number as stored. The attributes are
    private, so their presence is checked; without them the caller falls back to converting
    openpyxl's dates back to serials.
    """
    formats = getattr(workbook, "_date_formats", None)
    timedeltas = getattr(workbook, "_timedelta_formats", None)
    if not isinstance(formats, set) or not isinstance(timedeltas, set):
        return False
    workbook._date_formats = set()
    workbook._timedelta_formats = set()
    return True


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


def xlio_theme_fonts(theme: xlio.Theme | None) -> dict[str, str]:
    """The theme's Latin typefaces, ``{"major": ..., "minor": ...}``, those it names."""
    if theme is None:
        return {}
    pairs = (("major", theme.major_fonts), ("minor", theme.minor_fonts))
    return {key: fonts.latin for key, fonts in pairs if fonts is not None and fonts.latin}


class Book:
    """An open workbook, for reading several sheets with one load. Use :func:`open`.

    Use a book from one thread at a time: its readers share one open file.
    """

    def __init__(self, source: Source, *, limits: ReadLimits | None = None) -> None:
        self._limits = limits or ReadLimits()
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
        self._workbook: Any = None
        self._book: xlio.Workbook | None = None
        self._archive: zipfile.ZipFile | None = None
        try:
            self._archive = open_package(handle, self._limits, name)
            self._book = open_workbook(handle, self._limits)
            handle.seek(0)
            self._workbook = load_workbook(
                handle, read_only=True, data_only=True, keep_links=False, rich_text=True
            )
            self._raw_dates = disable_date_conversion(self._workbook)
            self._parts = worksheet_parts(self._archive)
            self._epoch = self._workbook.epoch
            self._theme = xlio_theme_colors(self._book.theme)
            self._palette: Sequence[str] = workbook_palette(
                self._book.styles.indexed_colors or None
            )
            self._theme_fonts = xlio_theme_fonts(self._book.theme)
            self._normal_font = self._read_normal_font()
            self._info = workbook_info(self._book.parts)
            self._marks = string_marks(self._book.parts, self._info.strings_part)
            self._context = styles.Context(
                self._cell_color, self._normal_font, self._theme_fonts, self._palette
            )
        except PackageError:
            self.close()
            raise
        except xlio.LimitError as error:
            self.close()
            raise LimitError(str(error)) from None
        except Exception as error:  # openpyxl raises many kinds on a file it cannot read
            self.close()
            raise self._unreadable(error) from error
        except BaseException:
            self.close()
            raise
        self._default_style = CellStyle("General", self._normal_font)
        self._styles: dict[Any, tuple[CellStyle, tuple[Code, ...], tuple[Code, ...]]] = {}
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
        if self._workbook is not None:
            self._workbook.close()
        if self._archive is not None:
            self._archive.close()
        self._handle.close()

    @property
    def sheet_names(self) -> list[str]:
        """Names of all sheets, hidden ones included, in workbook order."""
        return list(self._workbook.sheetnames)

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

        Only sheets openpyxl lists too, while it reads the cells; it skips a sheet whose part is
        missing. The active one is the first book view's ``activeTab``, which xlio gives as an
        index into its own list (a sheet listed without a part counts only as written).
        """
        assert self._book is not None
        present = set(self._workbook.sheetnames)
        listed = [
            (ListedSheet(s.name, _KIND.get(s.kind, "worksheet"), _STATE[s.state], None), s)
            for s in self._book
        ]
        kept: dict[str, tuple[ListedSheet, StatedColor | None]] = {}
        with self._xlio_errors():
            for sheet, source in listed:
                if sheet.name in present and sheet.name not in kept:
                    kept[sheet.name] = (sheet, from_xlio(source.properties.tab_color))
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
        return bool(self._epoch == CALENDAR_MAC_1904)

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
        if name not in self._workbook.sheetnames:
            raise KeyError(f"no sheet named {name!r}; the sheets are {self.sheet_names}")
        kind = self._info.kinds.get(name, "worksheet")
        if kind in ("chartsheet", "dialogsheet"):
            raise NotAWorksheetError(f"{name!r} {_NOT_SHOWN[kind]}")
        worksheet = self._workbook[name]
        if not hasattr(worksheet, "iter_rows") or name not in self._parts:
            raise NotAWorksheetError(f"{name!r} is not a worksheet")
        part = self._parts[name]
        if isinstance(ref, str):
            ref = CellRange.parse(ref)
        if ref is not None and not clip:
            self._check_area(ref)
        assert self._book is not None
        try:
            with self._xlio_errors():
                with self._book.parts.open(part) as handle:
                    structure = scan(handle, self._marks)
                source = self._book.sheet(name)
                structure = with_elements(fill(structure, source), source)
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
                sheet_notices(notes, self._book.parts, part, structure, self._info, name, ref)
            with _quiet_extensions():
                cells = self._read_cells(worksheet, ref, structure, notes) if ref else ()
        except PackageError:
            raise
        except Exception as error:
            raise self._unreadable(error) from error
        self._report_structure(notes, structure, ref)
        areas, area_problem, row_breaks, column_breaks = self._page_layout(name)
        rows, columns, outline = read_outline(
            structure, ref, _rows(structure, ref), _columns(structure, ref), notes
        )
        return Sheet(
            name=name,
            state=_STATES.get(worksheet.sheet_state, "visible"),
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
            outline=outline,
            view=self._view(name),
            print_areas=areas,
            print_area_problem=area_problem,
            row_breaks=row_breaks,
            column_breaks=column_breaks,
            notices=notes.notices(),
            max_cells=self._limits.max_cells,
        )

    def _report_structure(
        self, notes: Collector, structure: Structure, ref: CellRange | None
    ) -> None:
        """Notices from what the structure scan found: uncached formulas, runs, row styles."""
        if structure.uncached_formulas:
            ranges, truncated = structure.uncached.ranges()
            notes.extend(
                [
                    Notice(
                        "value.no_cached_value",
                        "omitted",
                        "not_implemented",
                        f"{structure.uncached_formulas} formula cell(s) have no cached value and "
                        f"show as empty (first: {structure.first_uncached_formula}); a file "
                        "saved by Excel stores the values",
                        notes.sheet,
                        ranges,
                        structure.uncached_formulas,
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

    def _read_cells(
        self, worksheet: Any, ref: CellRange, structure: Structure, notes: Collector
    ) -> tuple[tuple[Cell, ...], ...]:
        links = cell_links(structure, ref)
        converted_dates = False
        grid: list[tuple[Cell, ...]] = []
        row_index = ref.min_row
        for source_row in worksheet.iter_rows(
            min_row=ref.min_row, max_row=ref.max_row, min_col=ref.min_col, max_col=ref.max_col
        ):
            row: list[Cell] = []
            col_index = ref.min_col - 1
            for col_index, source in enumerate(source_row, ref.min_col):
                value, kind, from_date = self._value(source)
                converted_dates |= from_date
                if value is None and (row_index, col_index) in structure.empty_strings:
                    value, kind = "", "text"
                style, every, text = self._style(source)
                for code in every:
                    report_style(notes, code, row_index, col_index)
                if text and value is not None:
                    for code in text:
                        report_style(notes, code, row_index, col_index)
                row.append(
                    Cell(
                        row_index,
                        col_index,
                        value,
                        kind,
                        style,
                        links.get((row_index, col_index)),
                        self._runs(getattr(source, "value", None), style.font),
                    )
                )
            row.extend(self._empty(row_index, c) for c in range(col_index + 1, ref.max_col + 1))
            grid.append(tuple(row))
            row_index += 1
        # Read-only iteration stops after the last row holding data, whatever max_row says.
        grid.extend(
            tuple(self._empty(index, c) for c in range(ref.min_col, ref.max_col + 1))
            for index in range(row_index, ref.max_row + 1)
        )
        if converted_dates and not self._raw_dates:
            notes.add(
                "value.date_converted",
                "approximated",
                "needs_dependency",
                message=DATE_FALLBACK_WARNING,
                dependency="openpyxl",
            )
        report_links(notes, structure, ref, grid)
        _report_lost_runs(notes, structure, ref)
        return tuple(grid)

    def _empty(self, row: int, column: int) -> Cell:
        return Cell(row, column, None, "empty", self._default_style)

    def _value(self, source: Any) -> tuple[CellValue, CellKind, bool]:
        """Value, kind, and whether it was converted from an openpyxl date."""
        value = getattr(source, "value", None)
        if value is None:
            return None, "empty", False
        if getattr(source, "data_type", None) == "e":
            try:
                return xlformat.ExcelError(value), "error", False
            except ValueError:  # not shaped like an error code: a malformed file
                return str(value), "text", False
        if isinstance(value, bool):
            return value, "bool", False
        if isinstance(value, int | float):
            return value, "number", False
        if isinstance(value, str):
            return unescape(value), "text", False
        if isinstance(value, list):  # openpyxl's CellRichText: rich-text runs (§4.16)
            return "".join(text for text, _ in _parts(value)), "text", False
        if isinstance(value, _DATE_TYPES):
            # ISO 8601 cells (t="d") always arrive as dates; other date cells only when the
            # conversion could not be switched off.
            return to_excel(value, self._epoch), "number", True
        return str(value), "text", False

    def _runs(self, value: Any, font: Font) -> tuple[TextSpan, ...] | None:
        """A rich-text value's runs in their fonts; ``None`` when every run is in ``font``.

        A run without properties takes the cell's font. A run with properties stands alone: what
        they leave out is the default, not the cell's (Excel 16.114), and a missing name or
        size the Normal font's (unverified). Adjacent runs in one font are joined.
        """
        if not isinstance(value, list):
            return None
        cached = self._run_cache.get(id(value))
        if cached is not None and cached[0] is value and cached[1] == font:
            return cached[2]
        spans: list[TextSpan] = []
        for text, properties in _parts(value):
            own = (
                font
                if properties is None
                else font_aspect.to_font(
                    _RunFont(properties), self._normal_font, self._cell_color, self._theme_fonts
                )
            )
            if not text:
                continue
            if spans and _drawn(spans[-1].font) == _drawn(own):
                spans[-1] = TextSpan(spans[-1].text + text, own)
            else:
                spans.append(TextSpan(text, own))
        runs = tuple(spans) if any(_drawn(span.font) != _drawn(font) for span in spans) else None
        self._run_cache[id(value)] = (value, font, runs)  # the value keeps its id from reuse
        return runs

    def _style(self, source: Any) -> tuple[CellStyle, tuple[Code, ...], tuple[Code, ...]]:
        """The cell's style, and what it uses that is not read: (for any cell, for values)."""
        index = getattr(source, "_style_id", None)  # openpyxl's read-only cell: its xf index
        if index is None:  # an EmptyCell: no <c> element at this position
            return self._default_style, (), ()
        found = self._styles.get(index)
        if found is None:
            assert self._book is not None
            # An index past the cell formats raises IndexError: the sheet is unreadable.
            cell_format = self._book.styles.cell_formats[index]
            code = cell_format.number_format
            if code is None:
                code = BUILTIN_FORMATS.get(cell_format.number_format_id, "General")
            cell_format = dataclasses.replace(cell_format, number_format=code)
            found = self._styles[index] = styles.resolve(cell_format, self._context)
        return found

    def _cell_color(self, color: Any) -> str | None:
        """An ``xlio.Color`` as ``RRGGBB`` or ``None``.

        A color element without attributes is read as ``00000000``, black (unverified;
        openpyxl's reading, kept).
        """
        stated = from_xlio(color)
        if stated is None and color is not None:
            stated = StatedColor("rgb", rgb="00000000", tint=color.tint)
        return self._color(stated)

    def _color(self, color: Any) -> str | None:
        """A :class:`StatedColor` (or an object of its shape) as ``RRGGBB`` or ``None``."""
        if color is None:
            return None
        kind = getattr(color, "type", None)
        # openpyxl returns an error text, not a value, when a field of another type is read, so
        # only the field named by ``type`` is touched.
        value = getattr(color, kind, None) if kind in ("rgb", "indexed", "theme") else None
        return resolve(kind, value, getattr(color, "tint", 0.0) or 0.0, self._theme, self._palette)

    def _read_normal_font(self) -> Font:
        """The Normal style's font; else the first font; else Calibri 11."""
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
        if font is None:
            return font_aspect.FALLBACK
        return font_aspect.to_font(font, font_aspect.FALLBACK, self._cell_color, self._theme_fonts)


def _parts(value: list[Any]) -> Iterator[tuple[str, Any]]:
    """Each run of openpyxl's ``CellRichText``: its text, escapes decoded, and its properties.

    The properties are an ``InlineFont``, or ``None`` for a run without them.
    """
    for part in value:
        properties = getattr(part, "font", None)
        text = getattr(part, "text", part) if properties is not None else part
        yield unescape(str(text)), properties


def _drawn(font: Font) -> Font:
    """A font as drawn: what the file stored for a theme font (scheme, name) does not count."""
    return dataclasses.replace(font, scheme=None, stored_name=None) if font.scheme else font


class _RunFont:
    """A run's properties (openpyxl's ``InlineFont``) in the shape of an ``xlio.Font``."""

    __slots__ = ("_properties",)

    def __init__(self, properties: Any) -> None:
        self._properties = properties

    name = property(lambda self: self._properties.rFont)
    size = property(lambda self: self._properties.sz)
    bold = property(lambda self: self._properties.b)
    italic = property(lambda self: self._properties.i)
    underline = property(lambda self: self._properties.u)
    strike = property(lambda self: self._properties.strike)
    vert_align = property(lambda self: self._properties.vertAlign)
    scheme = property(lambda self: self._properties.scheme)

    @property
    def color(self) -> _RunColor | None:
        color = self._properties.color
        return None if color is None else _RunColor(color)


class _RunColor:
    """An openpyxl ``Color`` in the shape of an ``xlio.Color``.

    openpyxl returns an error text, not a value, when a field of another type is read, so only
    the field its ``type`` names is taken.
    """

    __slots__ = ("auto", "indexed", "rgb", "theme", "tint")

    def __init__(self, color: Any) -> None:
        kind = getattr(color, "type", None)
        self.indexed = color.indexed if kind == "indexed" else None
        self.theme = color.theme if kind == "theme" else None
        self.auto = kind == "auto"
        self.rgb = color.rgb if kind == "rgb" else None
        self.tint = getattr(color, "tint", 0.0) or 0.0


def _report_lost_runs(notes: Collector, structure: Structure, ref: CellRange) -> None:
    """Runs in inline strings, which openpyxl's read-only reader drops (§4.16)."""
    ranges, _ = structure.inline_runs.ranges()
    report_cells(notes, "text.runs_lost", (r for r in map(ref.intersection, ranges) if r))


@contextmanager
def _quiet_extensions() -> Iterator[None]:
    """Silence openpyxl's warning that it drops a sheet extension, while it parses a sheet.

    sheetpeek reports those extensions itself, as notices; openpyxl's warning would tell the
    caller a second time, about openpyxl. Only that warning, from openpyxl's sheet reader, is
    silenced. ``catch_warnings`` changes process-wide state: two threads reading at once can
    let one such warning through, which is harmless.
    """
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message=r".* extension is not supported and will be removed$",
            category=UserWarning,
            module=r"openpyxl\.worksheet\._reader$",
        )
        yield


def open(  # noqa: A001 - the module-level open() of gzip, tarfile, shelve
    source: Source, *, limits: ReadLimits | None = None
) -> Book:
    """Open a workbook to read several sheets with one load.

    Args:
        source: A path, or a binary file object (read into memory).
        limits: What reading may cost. By default, ``ReadLimits()``; lower them for files from
            people you do not trust.

    Raises:
        OSError: A path that cannot be opened.
        LimitError: The file exceeds ``limits``.
        PackageError: The file is not a readable xlsx workbook, or it is refused.

    Example:
        ``with sheetpeek.open("book.xlsx") as book: sheet = book.sheet("Data")``
    """
    return Book(source, limits=limits)


def read(
    source: Source,
    sheet: str | None = None,
    *,
    ref: str | CellRange | None = None,
    limits: ReadLimits | None = None,
) -> Sheet:
    """Read one sheet of a workbook.

    Args:
        source: A path, or a binary file object (read into memory).
        sheet: The sheet's name. By default, the sheet Excel opens on.
        ref: The range to read, such as ``"A1:H40"``. By default, the used range.
        limits: What reading may cost. By default, ``ReadLimits()``; lower them for files from
            people you do not trust.

    Raises:
        OSError: A path that cannot be opened.
        KeyError: There is no sheet with that name.
        NotAWorksheetError: The sheet is a chartsheet or a dialog sheet (a :class:`ValueError`).
        ValueError: ``ref`` is not a range.
        LimitError: The file, or the range, exceeds ``limits``.
        PackageError: The file is not a readable xlsx workbook, or it is refused.
    """
    with Book(source, limits=limits) as book:
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


def _columns(structure: Structure, ref: CellRange | None) -> tuple[Column, ...]:
    if ref is None:
        return ()
    found: dict[int, Column] = {}
    for spec in structure.columns:
        for index in range(max(spec.min, ref.min_col), min(spec.max, ref.max_col) + 1):
            found[index] = Column(index, spec.width, spec.hidden)
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
