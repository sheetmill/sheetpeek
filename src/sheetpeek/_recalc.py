"""What a formula cell shows on opening: its cached value, or the one xlrecalc computes.

Excel's rule (``recalc_on_load="excel"``) and its two alternatives are xlrecalc's
(``ARCHITECTURE.md`` §5.7). A :class:`Recalc` belongs to one ``Book``. Its
``xlrecalc.RecalcCheck`` reads the workbook part, which settles most files; otherwise it is fed
the formula cells as the sheets are streamed, and the worksheets not read yet get one pass for
their formulas only, ended as soon as the answer is ``"all"``. Unless the check has answered
``"none"``, the book's own cell stream passes through an ``xlrecalc.Loader`` as the reader
streams each worksheet, so the engine is built from the same reading of the file; it is
finished (sheets not read are read then) only when the answer is not ``"none"``, once per book,
and asked only for the formula cells of the range read and the cells their arrays fill.
"""

from __future__ import annotations

import io
import itertools
import os
import re
from collections.abc import Callable, Iterable, Iterator
from typing import Any, BinaryIO, Literal

import xlformat
import xlio
import xlrecalc

from sheetpeek._model import CellKind, CellValue
from sheetpeek._package import ReadLimits

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


class FileView(io.RawIOBase):
    """A read-only view of an open file with a position of its own (``os.pread``).

    xlrecalc reads the very file the book checked and reads, without moving the book's own
    position.
    """

    def __init__(self, fd: int) -> None:
        super().__init__()
        self._fd = fd
        self._size = os.fstat(fd).st_size
        self._position = 0

    def readable(self) -> bool:
        return True

    def seekable(self) -> bool:
        return True

    def tell(self) -> int:
        return self._position

    def seek(self, offset: int, whence: int = io.SEEK_SET) -> int:
        base = {io.SEEK_SET: 0, io.SEEK_CUR: self._position, io.SEEK_END: self._size}[whence]
        self._position = max(0, base + offset)
        return self._position

    def readinto(self, buffer: Any) -> int:
        view = memoryview(buffer).cast("B")
        data = os.pread(self._fd, len(view), self._position)
        view[: len(data)] = data
        self._position += len(data)
        return len(data)

    def close(self) -> None:  # the book's file stays open
        super().close()


class Recalc:
    """Recalculation on opening for one book: the check's answer, then the engine."""

    def __init__(
        self,
        source: Callable[[], BinaryIO],
        mode: RecalcOnLoad,
        limits: ReadLimits,
        book: xlio.Workbook,
    ) -> None:
        self.mode = mode
        self._limits = xlrecalc.ReadLimits(
            max_parts=limits.max_parts, max_part_bytes=limits.max_part_bytes
        )
        self._check: xlrecalc.RecalcCheck | None = None
        self._answer: Opening | None = None
        self._loader: xlrecalc.Loader | None = None
        self._loaded: set[str] = set()
        self._engine: xlrecalc.Workbook | None = None
        self.error: str | None = None
        """Why recalculation cannot be done for this book (xlrecalc refused it), else ``None``."""
        self.fed: set[str] = set()
        """The worksheets whose formulas the check has been given."""
        try:
            self._check = xlrecalc.RecalcCheck(source(), recalc_on_load=mode, limits=self._limits)
        except Exception as error:  # xlrecalc refuses what xlio read: report, keep cached values
            self.error = _reason(error)
            return
        if self._check.result is not None:
            self._answer = self._check.finish()
        if self._answer != "none":
            try:
                self._loader = xlrecalc.Loader(book, recalc_on_load=mode)
            except Exception as error:
                self.error = _reason(error)

    def cells(self, sheet: xlio.Sheet) -> Iterator[xlio.Cell]:
        """The worksheet's cells with their runs, through the loader the first time it is read.

        An error xlrecalc raises while recording ends the loader (recalculation cannot be done,
        :attr:`error`) and the rest of the sheet is read from xlio alone; an error of xlio's own
        is raised as it would be without the loader.
        """
        loader = self._loader
        if loader is None or sheet.name in self._loaded:
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
                self._loader = None
                yield from itertools.islice(sheet.cells(rich=True), given, None)
                return
            given += 1
            yield cell

    @property
    def feeding(self) -> bool:
        """Whether the check still takes formula cells."""
        return self._answer is None and self.error is None

    def add(self, formula: str | None, *, cached: bool, ca: bool) -> bool:
        """Give the check one formula cell; whether it takes more."""
        assert self._check is not None
        self._check.add(formula, cached=cached, ca=ca)
        if self._check.result is not None:
            self._settle(self._check.finish())
            return False
        return True

    def _settle(self, answer: Opening) -> None:
        self._answer = answer
        if answer == "none":
            self._loader = None  # nothing to compute: what it recorded is let go

    def answer(self, rest: Callable[[], Iterable[tuple[str, Iterable[Any]]]]) -> Opening | None:
        """The check's answer, after feeding it ``rest`` (each worksheet not fed, its cells).

        ``None`` when recalculation cannot be done (:attr:`error`).
        """
        if self._answer is None and self.error is None:
            assert self._check is not None
            try:
                more = True
                for name, cells in rest():
                    self.fed.add(name)
                    for cell in cells:
                        formula = cell.formula
                        if formula is None:
                            continue
                        ca = formula.attributes.get("ca") in ("1", "true")
                        more = self.add(formula.text, cached=cell.type != "none", ca=ca)
                        if not more:
                            break
                    if not more:
                        if self._loader is not None:  # the engine records the rest in this pass
                            for _ in cells:
                                pass
                        break
                if more:
                    self._settle(self._check.finish())
            except Exception as error:  # another sheet xlrecalc would read too
                self.error = _reason(error)
        return self._answer

    def engine(self) -> xlrecalc.Workbook | None:
        """The engine, loaded and recalculated on first use; ``None`` when xlrecalc fails.

        Everything Excel would compute on opening is computed at once, as Excel does, and each
        value is then read as it stands: asked for one by one, each stale cell would be a
        recalculation of its own, and in xlrecalc 0.4 each such recalculation goes over every
        cell that rests on a value it keeps (a workbook with a function it does not compute
        took hours).
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
