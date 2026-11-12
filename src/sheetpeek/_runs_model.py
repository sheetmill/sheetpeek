"""Rich-text runs: parts of one cell's text in fonts of their own (``ARCHITECTURE.md`` §4.16)."""

from __future__ import annotations

from dataclasses import dataclass

from sheetpeek._model import Font


@dataclass(frozen=True, slots=True)
class TextSpan:
    """A run of a cell's rich text, in one font.

    In :attr:`Cell.runs <sheetpeek.Cell.runs>` only ``text`` and ``font`` are set. In a laid-out
    :attr:`TextRun.spans <sheetpeek.TextRun.spans>` (``excel`` mode) the others are too, and the
    font's color is the color drawn: the number format's section color when it has one.
    """

    text: str
    font: Font
    """The run's own font, fully resolved. A run without properties takes the cell's font; a
    run with properties takes nothing from the cell's, only the Normal font's name and size where
    it states none."""
    width: float = 0.0
    """In a laid-out run, the span's width in points at its drawn size; 0 in the model."""
    scale: float = 1.0
    """In a laid-out run, the drawn size over ``font.size``: 2/3 for superscript and subscript,
    times the step of shrink to fit."""
    shift: float = 0.0
    """In a laid-out run, the points its baseline lies below the run's, or above it when
    negative, for superscript and subscript."""
