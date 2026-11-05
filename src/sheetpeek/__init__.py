"""Render Excel workbooks as HTML, Markdown and terminal text, showing what Excel shows.

``read`` gives a :class:`Sheet`: the cached values, number formats, styles, column widths,
row heights and merged ranges of one worksheet. See ``ARCHITECTURE.md``.
"""

from __future__ import annotations

from sheetpeek._model import (
    Cell,
    CellKind,
    CellRange,
    CellStyle,
    CellValue,
    Column,
    Font,
    HorizontalAlignment,
    Row,
    Sheet,
    SheetState,
    VerticalAlignment,
)
from sheetpeek._package import LimitError, PackageError, ReadLimits
from sheetpeek._read import Book, open, read  # noqa: A004 - sheetpeek.open, like gzip.open
from sheetpeek._rendered import (
    Alignment,
    Gap,
    Mode,
    RenderedCell,
    RenderedColumn,
    RenderedRow,
    RenderedSheet,
    SpillEdge,
    TextRun,
)
from sheetpeek.backends import Backend, register_backend, render

__version__ = "0.2.0"

__all__ = [
    "Alignment",
    "Backend",
    "Book",
    "Cell",
    "CellKind",
    "CellRange",
    "CellStyle",
    "CellValue",
    "Column",
    "Font",
    "Gap",
    "HorizontalAlignment",
    "LimitError",
    "Mode",
    "PackageError",
    "ReadLimits",
    "RenderedCell",
    "RenderedColumn",
    "RenderedRow",
    "RenderedSheet",
    "Row",
    "Sheet",
    "SheetState",
    "SpillEdge",
    "TextRun",
    "VerticalAlignment",
    "__version__",
    "open",
    "read",
    "register_backend",
    "render",
]
