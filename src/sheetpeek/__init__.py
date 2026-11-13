"""Render Excel workbooks as HTML, Markdown and terminal text, showing what Excel shows.

``read`` gives a :class:`Sheet`: the cached values, number formats, styles, column widths,
row heights and merged ranges of one worksheet. See ``ARCHITECTURE.md``.
"""

from __future__ import annotations

from sheetpeek._census import Presence
from sheetpeek._fill_border_model import (
    Borders,
    BorderStyle,
    Diagonal,
    Fill,
    Gradient,
    Pattern,
    PatternKind,
    RenderedBorders,
    Side,
    SlantedBorders,
)
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
from sheetpeek._runs_model import TextSpan
from sheetpeek._sheets import NotAWorksheetError, SheetInfo
from sheetpeek._text_model import TextRule
from sheetpeek._view_model import SheetView
from sheetpeek.backends import Backend, notices, register_backend, render
from sheetpeek.backends.html._workbook import render_workbook

__version__ = "0.2.0"

__all__ = [
    "FEATURES",
    "SHOWN",
    "Alignment",
    "Backend",
    "Book",
    "BorderStyle",
    "Borders",
    "Cell",
    "CellKind",
    "CellRange",
    "CellStyle",
    "CellValue",
    "Column",
    "Diagonal",
    "Feature",
    "Fill",
    "Font",
    "Gap",
    "Gradient",
    "HorizontalAlignment",
    "Hyperlink",
    "LimitError",
    "LinkSource",
    "Mode",
    "NotAWorksheetError",
    "NotShownWarning",
    "Notice",
    "Outcome",
    "Outline",
    "OutlineMark",
    "OutlineMarkKind",
    "PackageError",
    "Pattern",
    "PatternKind",
    "Presence",
    "ReadLimits",
    "Reason",
    "RenderedBorders",
    "RenderedCell",
    "RenderedColumn",
    "RenderedOutline",
    "RenderedRow",
    "RenderedSheet",
    "Row",
    "Sheet",
    "SheetInfo",
    "SheetState",
    "SheetView",
    "Side",
    "SlantedBorders",
    "SpillEdge",
    "Support",
    "TextRule",
    "TextRun",
    "TextSpan",
    "VerticalAlignment",
    "__version__",
    "notices",
    "open",
    "read",
    "register_backend",
    "render",
    "render_workbook",
]
