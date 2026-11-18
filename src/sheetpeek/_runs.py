"""Rich-text runs in the layout: a cell's text measured, broken and cut by its runs (§4.16).

Every rule here is unverified (no capture in this repository yet): widths add per run, and each
line of text takes the height of its largest run.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass

import xlformat

from sheetpeek._layout import _UNDRAWN
from sheetpeek._model import Cell, Font
from sheetpeek._rows import WRAPPED_LEADING, baseline_offsets, font_row_height
from sheetpeek._runs_model import TextSpan
from sheetpeek._text_styles import ACCOUNTING_EXTRA, accounting, scaled

Measure = Callable[[Font], xlformat.Font]
"""The font xlformat measures a model font with, at its drawn size."""


def drawn_spans(cell: Cell, text: str) -> tuple[TextSpan, ...] | None:
    """The cell's runs as they draw ``text``, what the cell shows; ``None`` to draw it in one font.

    Characters Excel does not draw are removed from each run as from the cell's text. A number
    format whose text section adds characters shows other text, and draws it in the cell's font
    (unverified).
    """
    if not cell.runs or cell.kind != "text":
        return None
    spans = [TextSpan(_UNDRAWN.sub("", s.text), s.font) for s in cell.runs]
    spans = [s for s in spans if s.text]
    if "".join(s.text for s in spans) != text:
        return None
    return tuple(spans)


@dataclass(frozen=True, slots=True)
class Part:
    """One run of a text being laid out: where it lies in the text, and its fonts."""

    start: int
    end: int
    font: Font
    measured: xlformat.Font


class Text:
    """A cell's text and the fonts it is drawn in, measured by offsets into the text.

    Text in one font is one part; a text with runs has a part per run.
    """

    def __init__(self, text: str, parts: list[Part]) -> None:
        self.text = text
        self.parts = parts

    @classmethod
    def plain(cls, text: str, font: Font, measured: xlformat.Font) -> Text:
        """Text in one font."""
        return cls(text, [Part(0, len(text), font, measured)])

    @classmethod
    def of_spans(
        cls, spans: tuple[TextSpan, ...], measure: Measure, *, breaks: bool, scale: float = 1.0
    ) -> Text:
        """Text in runs. Carriage returns are never drawn; line feeds only when ``breaks``."""
        text, parts, start = [], [], 0
        for span in spans:
            piece = span.text.replace("\r", "")
            if not breaks:
                piece = piece.replace("\n", "")
            if not piece:
                continue
            parts.append(
                Part(start, start + len(piece), span.font, scaled(measure(span.font), scale))
            )
            text.append(piece)
            start += len(piece)
        return cls("".join(text), parts)

    def pieces(self, start: int, end: int) -> Iterator[tuple[str, Part]]:
        """The text from ``start`` to ``end`` cut at the parts' edges, each piece with its part."""
        for part in self.parts:
            low, high = max(start, part.start), min(end, part.end)
            if low < high:
                yield self.text[low:high], part

    def width(self, start: int, end: int) -> float:
        """The width in points of the text from ``start`` to ``end``: its pieces' widths added."""
        if len(self.parts) == 1:
            return xlformat.text_width(self.text[start:end], font=self.parts[0].measured)
        return sum(xlformat.text_width(t, font=p.measured) for t, p in self.pieces(start, end))

    def fonts(self, start: int, end: int) -> list[Font]:
        """The fonts of the text from ``start`` to ``end``; an empty line takes the font there."""
        found = [p.font for _, p in self.pieces(start, end)]
        if found or not self.parts:
            return found
        before = [p for p in self.parts if p.start <= start]
        return [(before[-1] if before else self.parts[0]).font]

    def scaled(self, scale: float) -> Text:
        """The same text measured at ``scale`` times its size (shrink to fit)."""
        return Text(
            self.text,
            [Part(p.start, p.end, p.font, scaled(p.measured, scale)) for p in self.parts],
        )


def line_font(fonts: list[Font]) -> Font:
    """The font that sets a line's height and baseline: its largest run's (unverified)."""
    return max(fonts, key=lambda f: (font_row_height(f), f.size))


def line_metrics(fonts: list[Font], lines: int) -> tuple[float, float, float]:
    """(line pitch, baseline below its top, baseline above its bottom) of a line in ``fonts``.

    Excel's screen (``baseline_offsets``): wrapped lines are the font's row height apart, the
    point each adds to a fitted row lying below the last (§4.7); an accounting underline adds
    room under the text (§4.13). ``lines`` is the block's number of lines.
    """
    pitch, above, below = baseline_offsets(line_font(fonts))
    if any(accounting(f) for f in fonts):
        below += ACCOUNTING_EXTRA
    return pitch, above, below


def wrap(text: str, budget: float, width: Callable[[int, int], float]) -> list[tuple[int, int]]:
    """Lines of ``text`` broken at spaces to fit ``budget`` points, as offsets into ``text``.

    ``width(start, end)`` measures a part of the text. Explicit line breaks are kept; a word
    wider than the budget is broken between characters. A word joins a line only if the line
    fits with the space after the word, when one follows: Excel keeps that space on the line
    (captured, ``text_style`` groups ``wrap`` and ``lone``, ``ARCHITECTURE.md`` §4.7). A space
    at a break is not drawn, and neither are spaces that would start a line.
    """
    lines: list[tuple[int, int]] = []
    position = 0
    for paragraph in text.split("\n"):
        start = end = position  # the line so far is text[start:end]; empty when equal
        offset = position
        words = paragraph.split(" ")
        last = len(words) - 1
        for k, word in enumerate(words):
            word_start, word_end = offset, offset + len(word)
            offset = word_end + 1
            after = word_end + (k < last)  # with the space that follows, if any
            if end > start and width(start, after) > budget:
                lines.append((start, end))
                start = word_start
            elif end == start:
                start = word_start
            end = word_end
            while end - start > 1 and width(start, end) > budget:
                cut = end - start - 1
                while cut > 1 and width(start, start + cut) > budget:
                    cut -= 1
                lines.append((start, start + cut))
                start += cut
        lines.append((start, end))
        position += len(paragraph) + 1
    return lines


def fitted_height(text: Text, lines: list[tuple[int, int]]) -> float:
    """The height a fitted row gives the text's lines: each its largest run's, +1 pt if several."""
    extra = WRAPPED_LEADING if len(lines) > 1 else 0.0
    return sum(line_metrics(text.fonts(a, b), len(lines))[0] + extra for a, b in lines)


def runs_shrink_step(text: Text, budget: float, steps: int) -> int:
    """The largest ``k`` from 1 to ``steps`` at which the text at ``k / steps`` fits ``budget``."""
    for k in range(steps, 0, -1):
        if text.scaled(k / steps).width(0, len(text.text)) <= budget:
            return k
    return 1
