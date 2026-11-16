"""A cell's style, read aspect by aspect: number format, border, fill, font, alignment.

Each aspect module has ``read``, which gives the :class:`~sheetpeek.CellStyle` fields the model
carries of that aspect, and ``unread``, which gives what the style uses of it that the model does
not carry yet (reported as notices). A feature that comes to be drawn moves from one to the other
inside its own aspect module, so groups working on different aspects do not meet here.

The modules take xlio's style objects (``xlio.CellFormat`` and its parts), with the number
format's code resolved by the reader.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from typing import Any

from sheetpeek._colors import DEFAULT_PALETTE
from sheetpeek._model import CellStyle, Font
from sheetpeek._style import alignment, border, fill, font, number_format

Code = tuple[str, str, tuple[tuple[str, str], ...]]
"""What a style uses and the model does not carry: notice code, detail, context."""

Found = tuple[list[Code], list[Code]]
"""An aspect's findings: (for any cell, for cells with a value)."""


@dataclass(frozen=True, slots=True)
class Context:
    """What reading an aspect needs from the workbook."""

    color: Callable[[Any], str | None]
    """An ``xlio.Color`` to ``RRGGBB``, or ``None`` for automatic."""
    normal_font: Font
    theme_fonts: dict[str, str] = field(default_factory=dict)
    """The theme's Latin typefaces, ``{"major": ..., "minor": ...}``."""
    palette: Sequence[str] = DEFAULT_PALETTE
    """The workbook's indexed palette (``_colors.workbook_palette``)."""


ASPECTS = (number_format, border, fill, font, alignment)
"""The aspect modules, in the order their findings are reported."""


def resolve(source: Any, context: Context) -> tuple[CellStyle, tuple[Code, ...], tuple[Code, ...]]:
    """A cell's style, and what it uses that is not read: (for any cell, for cells with a value)."""
    fields: dict[str, Any] = {}
    every: list[Code] = []
    text: list[Code] = []
    for aspect in ASPECTS:
        part = getattr(source, aspect.ATTRIBUTE, None)
        fields.update(aspect.read(part, context))
        if part is not None:
            for_any, for_values = aspect.unread(part, context)
            every += for_any
            text += for_values
    return CellStyle(**fields), tuple(every), tuple(text)
