"""Render Excel workbooks as HTML, Markdown and terminal text, showing what Excel shows.

``read`` gives a :class:`Sheet`: the cached values, number formats, styles, column widths,
row heights and merged ranges of one worksheet. See ``ARCHITECTURE.md``.
"""

from __future__ import annotations

from sheetpeek._census import Presence
from sheetpeek._hyperlink import Hyperlink, LinkSource
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
from sheetpeek._notices import (
    FEATURES,
    SHOWN,
    Feature,
    Notice,
    NotShownWarning,
    Outcome,
    Reason,
    Support,
)
from sheetpeek._outline_model import Outline, OutlineMark, OutlineMarkKind, RenderedOutline
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
from sheetpeek.backends import Backend, notices, register_backend, render

__version__ = "0.2.0"

__all__ = [
    "FEATURES",
    "SHOWN",
    "Alignment",
    "Backend",
    "Book",
    "Cell",
    "CellKind",
    "CellRange",
    "CellStyle",
    "CellValue",
    "Column",
    "Feature",
    "Font",
    "Gap",
    "HorizontalAlignment",
    "Hyperlink",
    "LimitError",
    "LinkSource",
    "Mode",
    "NotShownWarning",
    "Notice",
    "Outcome",
    "Outline",
    "OutlineMark",
    "OutlineMarkKind",
    "PackageError",
    "Presence",
    "ReadLimits",
    "Reason",
    "RenderedCell",
    "RenderedColumn",
    "RenderedOutline",
    "RenderedRow",
    "RenderedSheet",
    "Row",
    "Sheet",
    "SheetState",
    "SpillEdge",
    "Support",
    "TextRun",
    "VerticalAlignment",
    "__version__",
    "notices",
    "open",
    "read",
    "register_backend",
    "render",
]
