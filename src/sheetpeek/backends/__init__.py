"""Backends turn a :class:`~sheetpeek.RenderedSheet` into text, and :func:`render` drives them.

A backend is any object with a ``name``, the width ``mode`` it needs, and a
``render(sheet, out)`` method that writes to a text stream. The built-in ones are registered by
name; any other object of that shape can be passed to :func:`render` directly, and
:func:`register_backend` gives it a name.
"""

from __future__ import annotations

import io
import os
from collections.abc import Callable
from typing import IO, Protocol, TextIO, runtime_checkable

from sheetpeek._model import CellRange, Sheet
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
) -> str | None:
    """Render a sheet with a backend.

    Args:
        source: A workbook (path or binary file object), a :class:`~sheetpeek.Sheet`, or a
            :class:`~sheetpeek.RenderedSheet`. A rendered sheet laid out in another mode than the
            backend needs is laid out again with default options.
        to: A backend name (``text``, ``html``, ``csv``, ``tsv``, ``markdown``, ``rich``) or a
            backend.
        sheet: The sheet to read, when ``source`` is a workbook.
        ref: The range to read, when ``source`` is a workbook.
        out: A text stream to write to. Without one, the output is returned as a string.

    Returns:
        The output, or ``None`` when it was written to ``out``.
    """
    backend = get_backend(to)
    if isinstance(source, RenderedSheet):
        rendered = source if source.mode == backend.mode else source.sheet.layout(backend.mode)
    else:
        if not isinstance(source, Sheet):
            from sheetpeek._read import read

            source = read(source, sheet, ref=ref)
        rendered = source.layout(backend.mode)
    if out is not None:
        backend.render(rendered, out)
        return None
    buffer = io.StringIO()
    backend.render(rendered, buffer)
    return buffer.getvalue()


__all__ = [
    "Backend",
    "DelimitedBackend",
    "HTMLBackend",
    "MarkdownBackend",
    "RichBackend",
    "TextBackend",
    "backend_names",
    "get_backend",
    "register_backend",
    "render",
]
