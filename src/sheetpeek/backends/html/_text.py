"""Text runs: each an absolutely positioned span, anchored as Excel anchors it, clipped.

A run's underlines and strikethrough are rules beside it in its clip window, one CSS pixel thick
at whole-pixel positions, as Excel draws them one screen point thick at any zoom (§4.13). An
angled run and its rules are turned with ``transform: rotate()`` about the start of its baseline
(§4.17).
"""

from __future__ import annotations

import html
from typing import TYPE_CHECKING

from sheetpeek._carlito import ADVANCES, FIRST, UNITS_PER_EM
from sheetpeek._model import Font, one_line
from sheetpeek._rendered import TextRun
from sheetpeek._runs_model import TextSpan
from sheetpeek._text_model import TextRule
from sheetpeek.backends.html._css import _ASCENT, _BASELINE, _CALIBRI, _DESCENT

if TYPE_CHECKING:
    from sheetpeek.backends.html._table import Page


def run(page: Page, run: TextRun) -> str:
    """One text run, positioned in sheet coordinates inside its clip window."""
    if run.spans:
        return _spanned(page, run)
    left, top, right, bottom = run.clip
    size = run.font.size * run.scale
    drawn = page.size(size)
    spacing = page.size(_spacing(run.font, run.text, run.width, run.scale))
    # A Calibri run spaced to Excel's width that lies inside its clip window needs no clip.
    inside = (
        not run.angle
        and run.font.name.lower() == "calibri"
        and not run.text.strip(_PRINTABLE)  # spaced from the Carlito table
        and left <= run.x
        and run.x + run.width <= right
        and top <= run.y - _ASCENT * size
        and run.y + _DESCENT * size <= bottom
        and (not run.rules or all(left <= r.x0 and r.x1 <= right and top <= r.y
             and r.y + 75 / page.zoom * (2 * r.lines - 1) <= bottom for r in run.rules))
    )  # fmt: skip
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
    baseline = page.y(run.y) - y0 - _BASELINE * drawn * 4 / 3
    style = [anchor, f"top:{_number(baseline)}px"]
    if _face(run.font, drawn) != page.normal_face:
        style.append(f"font:{_font(run.font, drawn)}")
    if run.direction == "rtl":
        style.append("direction:rtl")
    if run.color and run.color != "000000":  # the page's ink (--sheetpeek-ink) is black
        style.append(f"color:#{run.color}")
    if spacing:
        style.append(f"letter-spacing:{_number(spacing)}pt")
    if run.angle:
        style.append(_turn(run.angle, 0, _BASELINE * drawn * 4 / 3))
    span = f'<span class="{classes}" style="{";".join(style)}">{html.escape(run.text)}</span>'
    if run.rules:
        span += "".join(_rule(page, rule, run.color, x0, y0, run) for rule in run.rules)
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


def _face(font: Font, size: float | None = None) -> tuple[str, float, bool, bool]:
    return font.name.lower(), font.size if size is None else size, font.bold, font.italic


def _style(font: Font) -> str:
    if font.bold:
        return "bold italic" if font.italic else "bold"
    return "italic" if font.italic else "regular"


def _font(font: Font, size: float | None = None) -> str:
    family = _CALIBRI if font.name.lower() == "calibri" else f"'{_quoted(font.name)}', sans-serif"
    weight = "bold " if font.bold else ""
    italic = "italic " if font.italic else ""
    return f"{italic}{weight}{_number(font.size if size is None else size)}pt {family}"


def _quoted(name: str) -> str:
    return html.escape(one_line(name).replace("\\", "\\\\").replace("'", "\\'"))


def _rule(
    page: Page, rule: TextRule, color: str | None, x0: int, y0: int, run: TextRun | None = None
) -> str:
    """An underline or strikethrough: 1 px thick, a double one 1 px lines 2 px apart.

    A rule thicker than a point is its thickness at the page's zoom, a double one's lines two
    thicknesses apart (unverified past zoom 100). The rule of an angled run turns with it, about
    the start of the run's baseline.
    """
    left = page.x(rule.x0)
    width = page.x(rule.x1) - left
    if width <= 0:
        return ""
    classes = "sp-rule sp-double" if rule.lines == 2 else "sp-rule"
    tint = f";color:#{color}" if color else ""
    if rule.thickness > 1:
        thick = max(1, round(page.size(rule.thickness)))
        tint += (
            f";height:{3 * thick}px;border-width:{thick}px"
            if rule.lines == 2
            else f";height:{thick}px"
        )
    top = page.y(rule.y)
    turn = ""
    if run is not None and run.angle:
        turn = ";" + _turn(run.angle, page.x(run.x) - left, page.y(run.y) - top)
    return (
        f'<i class="{classes}" style="left:{left - x0}px;top:{top - y0}px;'
        f'width:{width}px{tint}{turn}"></i>'
    )


def _turn(angle: float, x: float, y: float) -> str:
    """CSS turning an element ``angle`` degrees counterclockwise about ``(x, y)`` in its box."""
    return f"transform:rotate({_number(-angle)}deg);transform-origin:{_number(x)}px {_number(y)}px"


def _spacing(font: Font, text: str, width: float, scale: float) -> float:
    """Points of letter-spacing that make Calibri text as wide as Excel draws it.

    Excel for Mac sets glyphs on whole points; a browser draws Calibri's own advances, which
    Carlito shares. Only Calibri text of printable ASCII is corrected. ``width`` is Excel's
    width of ``text`` drawn at ``scale`` times the font's size.
    """
    if font.name.lower() != "calibri" or not text:
        return 0.0
    table = ADVANCES[_style(font)]
    total = 0
    for char in text:
        k = ord(char) - FIRST
        if not 0 <= k < len(table):
            return 0.0
        total += table[k]
    own = total * font.size * scale / UNITS_PER_EM
    return (width - own) / len(text)


def _spanned(page: Page, run: TextRun) -> str:
    """A run of rich text: a span per run inside the positioned run, all on one baseline (§4.16).

    The positioned run takes the font of its largest span, which sets the line box; each span
    has ``line-height: 0``, so that none grows the line box and moves the baseline, as a larger
    run nested in the default way would. Superscript and subscript spans are moved by a relative
    offset, never ``vertical-align``, which also moves the baseline.
    """
    left, top, right, bottom = run.clip
    x0, x1, y0, y1 = page.x(left), page.x(right), page.y(top), page.y(bottom)
    if run.align == "left":
        anchor, classes = f"left:{page.x(run.x) - x0}px", "sp-run"
    elif run.align == "right":
        anchor, classes = f"right:{x1 - page.x(run.x + run.width)}px", "sp-run"
    else:
        center = (page.x(run.x) + page.x(run.x + run.width)) / 2 - x0
        anchor, classes = f"left:{center:g}px", "sp-run sp-center"
    largest = max(run.spans, key=lambda s: s.font.size * s.scale)
    face = (largest.font, page.size(largest.font.size * largest.scale))
    baseline = page.y(run.y) - y0 - _BASELINE * face[1] * 4 / 3
    style = [anchor, f"top:{_number(baseline)}px", f"font:{_font(*face)}"]
    if run.direction == "rtl":
        style.append("direction:rtl")
    if run.angle:
        style.append(_turn(run.angle, 0, _BASELINE * face[1] * 4 / 3))
    inner = "".join(_span(page, span, face) for span in run.spans)
    rules = "".join(_rule(page, r, _rule_color(run, r), x0, y0, run) for r in run.rules)
    return (
        f'<span class="sp-clip" style="left:{x0}px;top:{y0}px;width:{x1 - x0}px;'
        f'height:{y1 - y0}px"><span class="{classes}" style="{";".join(style)}">{inner}</span>'
        f"{rules}</span>"
    )


def _span(page: Page, span: TextSpan, face: tuple[Font, float]) -> str:
    """One run of rich text inside its line."""
    size = page.size(span.font.size * span.scale)
    style = ["line-height:0"]
    if _face(span.font, size) != _face(*face):
        style.append(f"font:{_font(span.font, size)}")
    if span.font.color:
        style.append(f"color:#{span.font.color}")
    spacing = page.size(_spacing(span.font, span.text, span.width, span.scale))
    if spacing:
        style.append(f"letter-spacing:{_number(spacing)}pt")
    if span.shift:
        style.append(f"position:relative;top:{round(page.size(span.shift) * 4 / 3)}px")
    return f'<span style="{";".join(style)}">{html.escape(span.text)}</span>'


def _rule_color(run: TextRun, rule: TextRule) -> str | None:
    """The color of the span a rule belongs to.

    That is the last span with the rule's kind that starts at or before the rule, else the first
    with that kind.
    """
    x, found = run.x, None
    for span in run.spans:
        drawn = span.font.strike if rule.kind == "strike" else span.font.underline is not None
        if drawn and (found is None or x <= rule.x0 + 0.01):
            found = span
        x += span.width
    return found.font.color if found is not None else run.color
