"""What a formula cell shows on opening: its cached value, or the one xlrecalc computes.

Excel's rule (``recalc_on_load="excel"``) and its two alternatives are xlrecalc's
(``ARCHITECTURE.md`` §5.7). A :class:`Recalc` belongs to one ``Book``. The book's own cell stream
passes through an ``xlrecalc.Loader`` as the reader streams each worksheet, so the engine is built
from the same reading of the file. The loader's ``opening`` answers which formulas Excel computes
on opening, from the workbook part and the formulas fed, or else from a formulas-only reading of
the worksheets not fed yet; it is asked only when a sheet read has formulas. Unless it answers
``"none"``, the engine is taken from the loader once per book and asked only for the formula cells
of the range read and the cells their arrays fill.
"""

from __future__ import annotations

import itertools
import re
from collections.abc import Iterator
from typing import Literal

import xlformat
import xlio
import xlrecalc

from sheetpeek._model import CellKind, CellValue

RecalcOnLoad = Literal["excel", "always", "never"]
RECALC_ON_LOAD: tuple[RecalcOnLoad, ...] = ("excel", "always", "never")
Opening = Literal["all", "some", "none"]

_EXACT = 2**53  # whole numbers below this are exact as floats
_FUNCTION = re.compile(r"[A-Z_][A-Z0-9_.]*")
_EXTERNAL_LABEL = "external or 3-D reference"
_EXTERNAL = re.compile(r"\[[^\]]*\]|[\w.']+:[\w.']+!")
"""A reference to another workbook (``[1]Sheet1!A1``) or across sheets (``Sheet1:Sheet3!A1``)."""


def recalc_mode(value: object) -> RecalcOnLoad:
    """``value`` if it is one of the three modes.

    Raises:
        ValueError: It is not.
    """
    for mode in RECALC_ON_LOAD:
        if value == mode and isinstance(value, str):
            return mode
    raise ValueError(f"recalc_on_load must be one of {RECALC_ON_LOAD}: {value!r}")


class Recalc:
    """Recalculation on opening for one book: the loader's answer, then the engine."""

    def __init__(self, mode: RecalcOnLoad, book: xlio.Workbook) -> None:
        self.mode = mode
        self._answer: Opening | None = None
        self._loader: xlrecalc.Loader | None = None
        self._loaded: set[str] = set()
        self._engine: xlrecalc.Workbook | None = None
        self.error: str | None = None
        """Why recalculation cannot be done for this book (xlrecalc refused it), else ``None``."""
        try:
            self._loader = xlrecalc.Loader(book, recalc_on_load=mode)
        except Exception as error:  # xlrecalc refuses what xlio read: report, keep cached values
            self.error = _reason(error)

    def cells(self, sheet: xlio.Sheet) -> Iterator[xlio.Cell]:
        """The worksheet's cells with their runs, through the loader the first time it is read.

        An error xlrecalc raises while recording ends the recording (recalculation cannot be
        done, :attr:`error`) and the rest of the sheet is read from xlio alone; an error of
        xlio's own is raised as it would be without the loader.
        """
        loader = self._loader
        if loader is None or self.error is not None or sheet.name in self._loaded:
            yield from sheet.cells(rich=True)
            return
        self._loaded.add(sheet.name)
        given = 0
        stream = loader.cells(sheet, rich=True)
        while True:
            try:
                cell = next(stream)
            except StopIteration:
                return
            except (xlio.PackageError, xlio.LimitError, MemoryError):
                raise
            except Exception as error:  # xlrecalc's recording failed: keep cached values
                self.error = _reason(error)
                yield from itertools.islice(sheet.cells(rich=True), given, None)
                return
            given += 1
            yield cell

    def answer(self) -> Opening | None:
        """Which formulas Excel computes on opening; ``None`` when it cannot be told.

        Asked once: the worksheets not fed yet have their formulas read for it, without their
        values, unless the workbook part settles it. It is asked even after a failure to
        record, so that a book whose cached values all stand is not reported. ``"none"`` lets
        the loader go, so the sheets read later are not recorded.
        """
        if self._answer is None and self._loader is not None:
            try:
                answer = self._loader.opening
            except (xlio.LimitError, MemoryError):
                raise
            except Exception as error:  # another sheet xlrecalc would read too
                answer = None
                self.error = self.error or _reason(error)
            if answer is None:  # also once the book is closed, which a reading book is not
                self.error = self.error or "xlrecalc could not read the workbook"
                self._loader = None
                return None
            self._answer = answer
            if answer == "none":
                self._loader = None  # nothing to compute: what it recorded is let go
        return self._answer

    def engine(self) -> xlrecalc.Workbook | None:
        """The engine, loaded and recalculated on first use; ``None`` when xlrecalc fails.

        Everything Excel would compute on opening is computed at once, as Excel does, and each
        value is then read as it stands. The report sheetpeek reads for each formula cell
        covers the whole workbook, so the sheets not fed are read here in any case.
        """
        if self._engine is None and self.error is None and self._loader is not None:
            loader, self._loader = self._loader, None
            try:
                engine = loader.finish()
                engine.recalculate()
            except Exception as error:
                self.error = _reason(error)
            else:
                self._engine = engine
        return self._engine


def shown(value: object) -> tuple[CellValue, CellKind] | None:
    """An engine value as a sheetpeek cell value and kind; ``None`` for a kind not shown."""
    if value is None:
        return None, "empty"
    if isinstance(value, xlrecalc.ErrorValue):
        try:
            return xlformat.ExcelError(value.code), "error"
        except ValueError:
            return None
    if isinstance(value, bool):
        return value, "bool"
    if isinstance(value, float):
        if value.is_integer() and -_EXACT < value < _EXACT:
            return int(value), "number"
        return value, "number"
    if isinstance(value, int):
        return value, "number"
    if isinstance(value, str):
        return value, "text"
    return None


def why_not(report: xlrecalc.Report, where: xlrecalc.CellRef, formula: str | None) -> str:
    """Why xlrecalc did not compute a cell, for a notice's detail."""
    error = report.exceptions.get(where)
    if error is not None:
        return f"an error inside xlrecalc ({error.splitlines()[0][:200]})"
    if formula is not None:
        text = formula.upper()
        named = sorted(
            label
            for label in report.unsupported
            if (found := _FUNCTION.match(label)) is not None and f"{found.group()}(" in text
        )
        if _EXTERNAL_LABEL in report.unsupported and _EXTERNAL.search(formula):
            named.append(_EXTERNAL_LABEL)
        if named:
            return "; ".join(named)
    return "depends on a formula xlrecalc does not compute"


def _reason(error: BaseException) -> str:
    """A refusal's first line, without where the file lives."""
    return f"xlrecalc could not read the workbook: {str(error).splitlines()[0][:300]}"
