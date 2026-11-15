"""Backends turn a :class:`~sheetpeek.RenderedSheet` into text, and :func:`render` drives them.

A backend is any object with a ``name``, the width ``mode`` it needs, and a
``render(sheet, out)`` method that writes to a text stream. It may also declare ``support``, a
mapping from feature code to :class:`~sheetpeek.Support`; a backend without one has every
feature it is given reported as not shown. The built-in ones are registered by
name; any other object of that shape can be passed to :func:`render` directly, and
:func:`register_backend` gives it a name.
"""

from __future__ import annotations

import io
import os
from collections.abc import Callable
from typing import IO, Literal, Protocol, TextIO, runtime_checkable

from sheetpeek._model import CellRange, Sheet
from sheetpeek._notices import Notice, Reason, deliver
from sheetpeek._package import ReadLimits
from sheetpeek._rendered import Mode, RenderedSheet
from sheetpeek.backends.delimited import DelimitedBackend
from sheetpeek.backends.html import HTMLBackend
from sheetpeek.backends.markdown import MarkdownBackend
from sheetpeek.backends.rich import RichBackend
from sheetpeek.backends.text import TextBackend


@runtime_checkable
class Backend(Protocol):
    """What :func:`render` needs from a backend."""

    name: str
    mode: Mode

    def render(self, sheet: RenderedSheet, out: TextIO, /) -> None:
        """Write the sheet to ``out``."""


_REGISTRY: dict[str, Callable[[], Backend]] = {
    "text": TextBackend,
    "csv": lambda: DelimitedBackend(","),
    "tsv": lambda: DelimitedBackend("\t"),
    "markdown": MarkdownBackend,
    "rich": RichBackend,
    "html": HTMLBackend,
}


def register_backend(name: str, factory: Callable[[], Backend]) -> None:
    """Make a backend available to :func:`render` by name.

    Args:
        name: The name passed as ``to``. An existing name is replaced.
        factory: Called with no arguments for each render, returning the backend.
    """
    _REGISTRY[name] = factory


def backend_names() -> list[str]:
    """The names :func:`render` accepts, sorted."""
    return sorted(_REGISTRY)


def get_backend(to: str | Backend) -> Backend:
    """The backend for a name, or ``to`` itself when it is already a backend.

    Raises:
        ValueError: no backend has that name.
    """
    if isinstance(to, str):
        try:
            return _REGISTRY[to]()
        except KeyError:
            raise ValueError(
                f"no backend named {to!r}; the backends are {backend_names()}"
            ) from None
    return to


def render(
    source: str | os.PathLike[str] | IO[bytes] | Sheet | RenderedSheet,
    to: str | Backend = "text",
    *,
    sheet: str | None = None,
    ref: str | CellRange | None = None,
    out: TextIO | None = None,
    limits: ReadLimits | None = None,
    on_notice: Callable[[Notice], object] | Literal["ignore"] | None = None,
    print_area: bool = False,
) -> str | None:
    """Render a sheet with a backend.

    Nothing the sheet uses is dropped silently: what the output leaves out or approximates is
    reported (:func:`notices`). By default one :class:`~sheetpeek.NotShownWarning` summarizes
    it; ``on_notice`` takes each :class:`~sheetpeek.Notice` instead, or ``"ignore"`` turns the
    warning off for this call.

    Args:
        source: A workbook (path or binary file object), a :class:`~sheetpeek.Sheet`, or a
            :class:`~sheetpeek.RenderedSheet`. A rendered sheet laid out in another mode than the
            backend needs is laid out again with default options.
        to: A backend name (``text``, ``html``, ``csv``, ``tsv``, ``markdown``, ``rich``) or a
            backend.
        sheet: The sheet to read, when ``source`` is a workbook.
        ref: The range to read, when ``source`` is a workbook.
        out: A text stream to write to. Without one, the output is returned as a string.
        limits: What reading may cost, when ``source`` is a workbook. By default,
            ``ReadLimits()``.
        on_notice: Called with each notice, ``"ignore"``, or ``None`` for one warning.
        print_area: Show the sheet's print area instead of its used range, as Excel prints it:
            each of its ranges as a block of its own, in the file's order, a range of whole
            columns or rows cut to the used range. Blocks are separated by a blank line (an
            empty record in CSV and TSV); HTML draws them one below another on one page. A sheet
            without a print area is shown whole. Off by default: Excel's Normal view shows the
            whole sheet. Needs a workbook as ``source``, and no ``ref``.

    Returns:
        The output, or ``None`` when it was written to ``out``.

    Raises:
        ValueError: ``print_area`` with a ``ref``, or with a sheet as ``source``.
    """
    backend = get_backend(to)
    if print_area:
        blocks, problem = _print_area(source, backend, sheet, ref, limits)
    else:
        blocks, problem = [_laid_out(source, backend, sheet, ref, limits)], None
    buffer = io.StringIO() if out is None else None
    target = out if out is not None else buffer
    assert target is not None
    found = write_blocks(backend, blocks, problem, target)
    deliver(found, on_notice, blocks[0].sheet.name, backend.name, 2)
    return buffer.getvalue() if buffer is not None else None


def write_blocks(
    backend: Backend, blocks: list[RenderedSheet], problem: str | None, out: TextIO
) -> tuple[Notice, ...]:
    """Write one or more blocks of a sheet; their notices, and the print area's problem's."""
    if len(blocks) > 1 and isinstance(backend, HTMLBackend):
        backend.render_blocks(blocks, out)
    else:
        for k, block in enumerate(blocks):
            if k:
                out.write("\n")
            backend.render(block, out)
    found = [n for block in blocks for n in _all_notices(block, backend)]
    if problem is not None:
        reason: Reason = "not_implemented" if "formula" in problem else "file"
        message = f"part of the print area is not shown: {problem}"
        found.append(
            Notice("print.area_unusable", "omitted", reason, message, blocks[0].sheet.name)
        )
    return tuple(dict.fromkeys(found))


def _print_area(
    source: str | os.PathLike[str] | IO[bytes] | Sheet | RenderedSheet,
    backend: Backend,
    sheet: str | None,
    ref: str | CellRange | None,
    limits: ReadLimits | None,
) -> tuple[list[RenderedSheet], str | None]:
    """The print area's blocks, laid out, and why a part of it was left out."""
    if ref is not None:
        raise ValueError("print_area shows the print area; it takes no ref")
    if isinstance(source, (Sheet, RenderedSheet)):
        raise ValueError("print_area reads the workbook: pass a path or a binary file object")
    from sheetpeek._read import Book

    with Book(source, limits=limits) as book:
        sheets, problem = book._print_area_sheets(sheet)
        if not sheets:  # nothing of any area is on the sheet: the sheet, empty
            sheets = (book.sheet(sheet),)
    return [s.layout(backend.mode) for s in sheets], problem


def notices(
    source: str | os.PathLike[str] | IO[bytes] | Sheet | RenderedSheet,
    to: str | Backend = "text",
    *,
    sheet: str | None = None,
    ref: str | CellRange | None = None,
    limits: ReadLimits | None = None,
) -> tuple[Notice, ...]:
    """What a render of ``source`` to ``to`` would leave out or approximate, without rendering.

    The reader's notices (what sheetpeek does not read yet), the layout's (what every output
    approximates) and the output's own (what this output cannot show), in that order.

    Args:
        source: As for :func:`render`.
        to: A backend name or a backend.
        sheet: The sheet to read, when ``source`` is a workbook.
        ref: The range to read, when ``source`` is a workbook.
        limits: What reading may cost, when ``source`` is a workbook.
    """
    backend = get_backend(to)
    return _all_notices(_laid_out(source, backend, sheet, ref, limits), backend)


def _laid_out(
    source: str | os.PathLike[str] | IO[bytes] | Sheet | RenderedSheet,
    backend: Backend,
    sheet: str | None,
    ref: str | CellRange | None,
    limits: ReadLimits | None,
) -> RenderedSheet:
    if isinstance(source, RenderedSheet):
        return source if source.mode == backend.mode else source.sheet.layout(backend.mode)
    if not isinstance(source, Sheet):
        from sheetpeek._read import read

        source = read(source, sheet, ref=ref, limits=limits)
    return source.layout(backend.mode)


def _all_notices(rendered: RenderedSheet, backend: Backend) -> tuple[Notice, ...]:
    from sheetpeek._census import declared_support, output_notices

    return rendered.notices + output_notices(rendered, backend.name, declared_support(backend))


__all__ = [
    "Backend",
    "DelimitedBackend",
    "HTMLBackend",
    "MarkdownBackend",
    "RichBackend",
    "TextBackend",
    "backend_names",
    "get_backend",
    "notices",
    "register_backend",
    "render",
]
