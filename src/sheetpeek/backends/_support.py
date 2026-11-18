"""What each built-in output shows of the features the census finds (``Backend.support``).

One table per output, side by side, so that the choices can be read and reviewed together. A
code an output leaves out of its table is reported as not declared; ``tests/test_notices.py``
checks that every census code is declared by every built-in output.
"""

from __future__ import annotations

from sheetpeek._notices import SHOWN, Support

_approx = Support.approximated
_omit = Support.omitted


def _no_borders() -> Support:
    return _omit("not_implemented", "borders are drawn only in HTML so far")


def _local() -> Support:
    return _omit("option", "file and relative targets are linked only with local_links=True")


def _headings(drawn: bool) -> Support:
    if drawn:
        return _approx("drawn, because a headings option is set; the file hides them", "option")
    return SHOWN


def _outline_html(buttons: str, level_buttons: bool) -> Support:
    if buttons == "none":
        return _omit("option", "buttons='none' leaves the outline gutter out")
    if not level_buttons:
        return _approx("the level buttons are drawn only beside both headings", "option")
    return SHOWN


def _filter_buttons(buttons: str) -> Support:
    if buttons == "none":
        return _omit("option", "buttons='none' leaves the filter buttons out")
    return SHOWN


def _outline_characters(buttons: str) -> Support:
    if buttons == "none":
        return _omit("option", "buttons='none' leaves the outline gutter out")
    return _approx("a gutter of characters: bars, - and + buttons, dots; no level buttons")


def html(
    *,
    headings: bool,
    buttons: str = "static",
    level_buttons: bool = False,
    local_links: bool = False,
) -> dict[str, Support]:
    """The HTML output (``excel`` mode)."""
    return {
        "font.name": SHOWN,
        "font.size": SHOWN,
        "font.bold": SHOWN,
        "font.italic": SHOWN,
        "text.runs": SHOWN,
        "font.underline": SHOWN,
        "font.underline_accounting": SHOWN,
        "font.strike": SHOWN,
        "font.vert_align": SHOWN,
        "text.shrink": SHOWN,
        "text.reading_order": SHOWN,
        "text.rotation": SHOWN,
        "text.stacked": SHOWN,
        "font.color": SHOWN,
        "grid.color": SHOWN,  # 0.5 CSS px: one device pixel on a 2x display, as Excel draws it
        "format.color": SHOWN,
        "fill": SHOWN,
        "fill.pattern": SHOWN,
        "fill.gradient": SHOWN,
        "border": SHOWN,
        "border.diagonal": SHOWN,
        "border.slant_dash_dot": _approx("as mediumDashDot, gaps square", "not_implemented"),
        "border.slanted": _approx(
            "a merge's borders are drawn upright and its text placed as without borders; "
            "Excel slants an angled cell's borders and places its text between them",
            "not_implemented",
        ),
        "align.horizontal": SHOWN,
        "align.vertical": SHOWN,
        "align.wrap": SHOWN,
        "align.indent": SHOWN,
        "link": SHOWN,
        "link.internal": SHOWN,
        "link.elsewhere": _approx("a link that goes nowhere: its target is not on this page"),
        "link.local": SHOWN if local_links else _local(),
        "merge": SHOWN,
        "column.width": SHOWN,
        "row.height": SHOWN,
        "gridlines.hidden": SHOWN,
        "headings.hidden": _headings(headings),
        "outline": _outline_html(buttons, level_buttons),
        "filter.buttons": _filter_buttons(buttons),
    }


def _characters(*, headings: bool) -> dict[str, Support]:
    """What ``rich`` and plain text share (``characters`` mode)."""
    return {
        "align.horizontal": _approx("placed in whole characters"),
        "align.vertical": _approx("placed in whole lines"),
        "align.wrap": _approx("wrapped by characters"),
        "align.indent": _approx("one character a level"),
        "merge": _approx("the anchor's text across the merged area"),
        "column.width": _approx("in whole characters"),
        "row.height": _approx("as many lines as the text needs"),
        "gridlines.hidden": _approx("column separators drawn"),
        "headings.hidden": _headings(headings),
    }


def rich(
    *, headings: bool, buttons: str = "static", local_links: bool = False
) -> dict[str, Support]:
    """The ``rich`` terminal output."""
    terminal = _approx("in the terminal's colors")
    return {
        "font.name": _omit(),
        "font.size": _omit(),
        "font.bold": SHOWN,
        "font.italic": SHOWN,
        "text.runs": _approx("drawn in the cell's font"),
        "font.underline": _approx("as the terminal's underline"),
        "font.underline_accounting": _approx("as the terminal's underline"),
        "font.strike": SHOWN,
        "font.vert_align": _approx("drawn at full size on the baseline"),
        "text.shrink": _approx("clipped at full size; a number too wide shows #"),
        "text.reading_order": _approx("the terminal's own text direction"),
        "text.rotation": _approx("drawn horizontally"),
        "text.stacked": _approx("drawn horizontally"),
        "font.color": terminal,
        "grid.color": _approx("the column separators in the grid color"),
        "format.color": terminal,
        "fill": terminal,
        "fill.pattern": _approx("one color, the blend of its two by the pattern's share"),
        "fill.gradient": _approx("one color, the gradient's at its midpoint"),
        "border": _no_borders(),
        "border.diagonal": _no_borders(),
        "border.slant_dash_dot": _no_borders(),
        "border.slanted": _no_borders(),
        "link": SHOWN,  # a terminal hyperlink (OSC 8), plain text where a terminal has none
        "link.internal": _omit(why="a terminal has no place in the sheet to go to"),
        "link.elsewhere": _omit(why="a terminal has no place in the workbook to go to"),
        "link.local": SHOWN if local_links else _local(),
        **_characters(headings=headings),
        "outline": _outline_characters(buttons),
        "filter.buttons": _omit("not_implemented", "filter buttons are not drawn yet"),
    }


def text(*, headings: bool, buttons: str = "static") -> dict[str, Support]:
    """Plain fixed-width text."""
    plain = _omit()
    return {
        "font.name": plain,
        "font.size": plain,
        "font.bold": plain,
        "font.italic": plain,
        "text.runs": plain,
        "font.underline": plain,
        "font.underline_accounting": plain,
        "font.strike": plain,
        "font.vert_align": plain,
        "text.shrink": _approx("clipped at full size; a number too wide shows #"),
        "text.reading_order": _approx("the terminal's own text direction"),
        "text.rotation": _approx("drawn horizontally"),
        "text.stacked": _approx("drawn horizontally"),
        "font.color": plain,
        "grid.color": plain,
        "format.color": plain,
        "fill": plain,
        "fill.pattern": plain,
        "fill.gradient": plain,
        "border": _no_borders(),
        "border.diagonal": _no_borders(),
        "border.slant_dash_dot": _no_borders(),
        "border.slanted": _no_borders(),
        "link": plain,
        "link.internal": plain,
        "link.elsewhere": plain,
        "link.local": plain,
        **_characters(headings=headings),
        "outline": _outline_characters(buttons),
        "filter.buttons": _omit("not_implemented", "filter buttons are not drawn yet"),
    }


def markdown(*, emphasis: bool, letters: bool, local_links: bool = False) -> dict[str, Support]:
    """Markdown (``none`` mode: full text, no widths)."""
    plain = _omit()
    emphasized = SHOWN if emphasis else _omit("option", "emphasis is off")
    return {
        "font.name": plain,
        "font.size": plain,
        "font.bold": emphasized,
        "font.italic": emphasized,
        "text.runs": SHOWN if emphasis else _omit("option", "emphasis is off"),
        "font.underline": plain,
        "font.underline_accounting": plain,
        "font.strike": _omit("not_implemented"),
        "font.vert_align": _approx("drawn as plain text"),
        "text.shrink": SHOWN,  # the full text: nothing to shrink
        "text.reading_order": plain,
        "text.rotation": plain,
        "text.stacked": plain,
        "font.color": plain,
        "grid.color": plain,
        "format.color": plain,
        "fill": plain,
        "fill.pattern": plain,
        "fill.gradient": plain,
        "border": plain,
        "border.diagonal": plain,
        "border.slant_dash_dot": plain,
        "border.slanted": plain,
        "align.horizontal": _approx("one alignment per column, from the first body row"),
        "align.vertical": plain,
        "align.wrap": plain,
        "align.indent": plain,
        "link": SHOWN,
        "link.internal": _omit(why="a Markdown table has no anchors to link to"),
        "link.elsewhere": _omit(why="a Markdown table has no anchors to link to"),
        "link.local": SHOWN if local_links else _local(),
        "merge": _approx("the anchor's text in its cell, the covered cells empty"),
        "column.width": plain,
        "row.height": plain,
        "gridlines.hidden": plain,
        "headings.hidden": _headings(letters),
        "outline": plain,
        "filter.buttons": plain,
    }


def delimited() -> dict[str, Support]:
    """CSV and TSV: values only."""
    plain = _omit()
    return {
        "font.name": plain,
        "font.size": plain,
        "font.bold": plain,
        "font.italic": plain,
        "text.runs": plain,
        "font.underline": plain,
        "font.underline_accounting": plain,
        "font.strike": plain,
        "font.vert_align": plain,
        "text.shrink": SHOWN,  # the full text: nothing to shrink
        "text.reading_order": plain,
        "text.rotation": plain,
        "text.stacked": plain,
        "font.color": plain,
        "grid.color": plain,
        "format.color": plain,
        "fill": plain,
        "fill.pattern": plain,
        "fill.gradient": plain,
        "border": plain,
        "border.diagonal": plain,
        "border.slant_dash_dot": plain,
        "border.slanted": plain,
        "align.horizontal": plain,
        "align.vertical": plain,
        "align.wrap": plain,
        "align.indent": plain,
        "link": plain,
        "link.internal": plain,
        "link.elsewhere": plain,
        "link.local": plain,
        "merge": _approx("the anchor's text in its cell, the covered cells empty"),
        "column.width": plain,
        "row.height": plain,
        "gridlines.hidden": plain,
        "headings.hidden": SHOWN,
        "outline": plain,
        "filter.buttons": plain,
    }
