"""Text runs: each an absolutely positioned span, anchored as Excel anchors it, clipped."""

from __future__ import annotations

import html
from typing import TYPE_CHECKING

from sheetpeek._carlito import ADVANCES, FIRST, UNITS_PER_EM
from sheetpeek._model import Font, one_line
from sheetpeek._rendered import TextRun
from sheetpeek.backends.html._css import _ASCENT, _BASELINE, _CALIBRI, _DESCENT

if TYPE_CHECKING:
    from sheetpeek.backends.html._table import Page


def run(page: Page, run: TextRun) -> str:
    """One text run, positioned in sheet coordinates inside its clip window."""
    left, top, right, bottom = run.clip
    size = run.font.size
    spacing = _spacing(run)
    # A Calibri run spaced to Excel's width that lies inside its clip window needs no clip.
    inside = (
        run.font.name.lower() == "calibri"
        and not run.text.strip(_PRINTABLE)  # spaced from the Carlito table
        and left <= run.x
        and run.x + run.width <= right
        and top <= run.y - _ASCENT * size
        and run.y + _DESCENT * size <= bottom
    )
    # Each run is anchored where Excel anchors it: left text by its start, right text by its
    # end, centered text by its center, so another font changes only the far end.
    x0, x1, y0, y1 = (
        (0, page.width, 0, 0)
        if inside
        else (page.x(left), page.x(right), page.y(top), page.y(bottom))
    )
    if run.align == "left":
        anchor, classes = f"left:{page.x(run.x) - x0}px", "sp-run"
    elif run.align == "right":
        anchor, classes = f"right:{x1 - page.x(run.x + run.width)}px", "sp-run"
    else:
        center = (page.x(run.x) + page.x(run.x + run.width)) / 2 - x0
        anchor, classes = f"left:{center:g}px", "sp-run sp-center"
    baseline = page.y(run.y) - y0 - _BASELINE * size * 4 / 3
    style = [anchor, f"top:{_number(baseline)}px"]
    if _face(run.font) != _face(page.normal):
        style.append(f"font:{_font(run.font)}")
    decoration = _decoration(run.font)
    if decoration:
        style.append(f"text-decoration:{decoration}")
    if run.color:
        style.append(f"color:#{run.color}")
    if spacing:
        style.append(f"letter-spacing:{_number(spacing)}pt")
    span = f'<span class="{classes}" style="{";".join(style)}">{html.escape(run.text)}</span>'
    if inside:
        return span
    return (
        f'<span class="sp-clip" style="left:{x0}px;top:{y0}px;width:{x1 - x0}px;'
        f'height:{y1 - y0}px">{span}</span>'
    )


_PRINTABLE = "".join(chr(code) for code in range(32, 127))


def _number(value: float) -> str:
    """A CSS number to three decimals, without trailing zeros."""
    text = f"{value:.3f}".rstrip("0").rstrip(".")
    return "0" if text in ("-0", "") else text


def _face(font: Font) -> tuple[str, float, bool, bool]:
    return font.name.lower(), font.size, font.bold, font.italic


def _style(font: Font) -> str:
    if font.bold:
        return "bold italic" if font.italic else "bold"
    return "italic" if font.italic else "regular"


def _font(font: Font) -> str:
    family = _CALIBRI if font.name.lower() == "calibri" else f"'{_quoted(font.name)}', sans-serif"
    weight = "bold " if font.bold else ""
    italic = "italic " if font.italic else ""
    return f"{italic}{weight}{font.size:g}pt {family}"


def _quoted(name: str) -> str:
    return html.escape(one_line(name).replace("\\", "\\\\").replace("'", "\\'"))


def _decoration(font: Font) -> str:
    parts = []
    if font.underline in ("double", "doubleAccounting"):
        parts.append("underline double")
    elif font.underline:
        parts.append("underline")
    if font.strike:
        parts.append("line-through")
    return " ".join(parts)


def _spacing(run: TextRun) -> float:
    """Points of letter-spacing that make Calibri text as wide as Excel draws it.

    Excel for Mac sets glyphs on whole points; a browser draws Calibri's own advances, which
    Carlito shares. Only Calibri runs of printable ASCII are corrected.
    """
    font = run.font
    if font.name.lower() != "calibri" or not run.text:
        return 0.0
    table = ADVANCES[_style(font)]
    total = 0
    for char in run.text:
        k = ord(char) - FIRST
        if not 0 <= k < len(table):
            return 0.0
        total += table[k]
    own = total * font.size / UNITS_PER_EM
    return (run.width - own) / len(run.text)
