"""Capture how Excel for Mac places and decorates text, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.screen run text_style
    uv run --group capture python -m capture.screen read text_style     # the saved captures again
    uv run --group capture python -m capture.text_style run wrap lone   # some groups only
    uv run --group capture python -m capture.text_style build PATH      # the workbooks only

Five groups, each its own workbook and window (their layouts share cells), with the default
markers B2 and J20. Rows are fitted by Excel on opening unless a group sets a height.

- ``align``, zoom 100. Rows 4, 6, 8, 10 and 12 are 60 pt high with vertical alignment top,
  center, justify, distributed and bottom: column C (24 characters) holds one line, "Hxg"; E
  three lines, "aaa", "bbb" and "ccc" with wrap on; G "aaaa bbbb cccc dddd" with wrap off (does
  vertical justify or distributed wrap by itself?). Horizontal distributed: C14 "aa bb cc", E14
  "abc" (one word), G14 the long text with wrap off (does distributed wrap?); C16 "aa bb cc" with
  indent 2; G16 the long text wrapped with ``justifyLastLine`` and G18 without it; their rows are
  fitted to the lines drawn. Indents: right alignment with indent 0, 1 and 3 on "HHH" (C22, E22,
  G22), the number 123 right aligned with indent 2 (I22) and with general alignment and indent 2
  (G24), "HHH" left with indent 3 (C24). Reading order on "abc!" (C26 context, E26 right to left)
  and on three Hebrew letters and "!" (G26 context, I26 left to right). Measured: the text's
  ink box, its lines (ink bands of rows) and its words (ink clusters of columns).
- ``shrink``, zoom 100 and 200. Rows 4 to 22: C (default width) holds "H" repeated 4 to 22
  times with shrink to fit, left aligned; G (60 characters) the same strings without shrink.
  Measured: the ink width and height of C over G's, so the scale and whether it is stepped. And
  in column E: wrap together with shrink (E4: which wins, and how many lines?), the number
  123456789012 with shrink (E6: shrunk, or ``#``?), 12 "H" with shrink and left indent 2 (E8),
  right aligned (E10), at 20 pt (E12: is the row fitted to the shrunk size?), and a long text
  with shrink in the merge E14:F14. Measured: the ink box, the lines, the fitted row height.
- ``fonts``, zoom 100 and 200. Rows 4 to 12: underline none, single, double, single accounting
  and double accounting on "HHH" (C, 14 characters, left), the number 123 (E, right) and "HH" in
  a wide cell (G, 30 characters: does an accounting underline span the cell?). Rows 14 and 16:
  "HHH" plain (C), superscript (E) and subscript (G) at 11 and 24 pt. C18 superscript alone, C20
  subscript alone (their fitted rows), G18 strikethrough. C22 a font that is not installed, G22
  the same text in Calibri. Row 24 at 24 pt: outline (C), shadow (E), plain (G), condense and
  extend (I). Measured: the glyphs' baseline and the horizontal rules (underlines, strike):
  each rule's offset below the baseline, thickness, and where it starts and ends; the ink box.
- ``wrap``, zoom 100 and 200. The width a wrapped line may take: two words measuring 44 to 50
  points with the space (``NARROW``) in default columns, wrap alone (C) and with shrink to fit
  (E), rows 4 to 16; 111 to 117 points (``WIDE``) in columns of 20 characters (G and I) the same
  way; single words of 45 to 50 points (``WORDS``, broken between characters when too wide),
  rows 18 to 28 in C and E; the 47-point pair distributed (C30) and justified (E30), and E4's
  text of ``shrink`` with wrap alone (G30). Measured: the lines and the ink box.
- ``lone``, zoom 100. Where distributed text puts a line of one word: "spilled text that is
  much wider than one column" in default columns, its rows fitted (row 4): distributed (C),
  with ``justifyLastLine`` (E), justified (G), distributed with indent 1 (I); in a 15 pt row
  (row 6): distributed without and with wrap (C, E), wrapped (G), distributed at the top (I);
  row 8 distributed: "spilled" and "abc" alone (C, E), a lone word in the middle line (G), at
  the end (I). Measured: the lines, the ink box and each line's ink (``line_inks``).

A cell outside the window at zoom 200 is skipped at that zoom. ``read`` writes
``tests/captures/text_style.tsv``: one row per case and zoom, lengths in points: x from the
cell's left edge (``ink_left``, ``rule_start``, ``words``) or from its right edge, or the
merge's (``ink_right``, ``rule_end``, negative inside), y from the cell's top edge (``ink_top``,
``ink_bottom``, ``line_tops``, ``baseline``), rules' offsets down from the baseline; lists are
space separated, ``start+length`` each (``line_inks`` from the cell's left edge, one a line). A
group not captured yet is left out. ``row_height`` and the ``display`` fields are Excel's
reads.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "text_style"
OUTPUT = ROOT / "tests" / "captures" / "text_style.tsv"
GROUPS = ("align", "shrink", "fonts", "wrap", "lone")
ZOOMS = {
    "align": (100,), "shrink": (100, 200), "fonts": (100, 200), "wrap": (100, 200), "lone": (100,),
}  # fmt: skip
COLUMNS = (
    "group", "name", "zoom", "cell", "row_height", "cell_width", "lines", "line_tops", "ink_left",
    "ink_right", "ink_top", "ink_bottom", "words", "baseline", "glyph_top", "rule_offset",
    "rule_thickness", "rule_start", "rule_end", "width_ratio", "height_ratio", "ink_width",
    "ref_width", "font_name", "font_size", "underline", "h_align", "v_align", "text", "line_inks",
)  # fmt: skip
READS = {
    "row_height": "geometry.height",
    "cell_width": "geometry.width",
    "font_name": "display.font_name",
    "font_size": "display.font_size",
    "underline": "display.underline",
    "h_align": "display.h_align",
    "v_align": "display.v_align",
    "text": "geometry.text",
}

LONG = "aaaa bbbb cccc dddd"
HEBREW = "אבג!"
VERTICAL = {4: "top", 6: "center", 8: "justify", 10: "distributed", 12: "bottom"}
UNDERLINES = {4: None, 6: "single", 8: "double", 10: "singleAccounting", 12: "doubleAccounting"}
SAMPLE = "Hamburgefonstiv 0123"
MISSING = "Sheetpeek Missing Sans"
"""A font name no system has, for Excel's fallback."""


@dataclass(frozen=True)
class Case:
    """One measured cell: what it shows and what is measured (``ink``, ``lines``, ``words``,
    ``rules``, ``scale``)."""

    group: str
    name: str
    cell: str
    measure: tuple[str, ...]
    last: str | None = None
    """The last cell of the measured box (a merge, or text spilling right), if not ``cell``."""
    reference: str | None = None
    """For ``scale``: the cell with the same text unshrunk."""


def _align(ws: Any) -> list[Case]:
    ws.column_dimensions["C"].width = 24
    found = []
    for row, vertical in VERTICAL.items():
        ws.row_dimensions[row].height = 60
        ws[f"C{row}"] = "Hxg"
        ws[f"E{row}"] = "aaa\nbbb\nccc"
        ws[f"G{row}"] = LONG
        ws[f"C{row}"].alignment = Alignment(vertical=vertical)
        ws[f"E{row}"].alignment = Alignment(vertical=vertical, wrap_text=True)
        ws[f"G{row}"].alignment = Alignment(vertical=vertical)
        found.extend(
            Case("align", f"vertical {vertical} {label}", f"{column}{row}", ("lines", "ink"))
            for column, label in (("C", "one line"), ("E", "three lines"), ("G", "long unwrapped"))
        )
    ws["C14"], ws["E14"], ws["G14"] = "aa bb cc", "abc", LONG
    for column in "CEG":
        ws[f"{column}14"].alignment = Alignment(horizontal="distributed")
    ws["C16"] = "aa bb cc"
    ws["C16"].alignment = Alignment(horizontal="distributed", indent=2)
    ws["G16"] = ws["G18"] = LONG
    ws["G16"].alignment = Alignment(horizontal="distributed", wrap_text=True, justifyLastLine=True)
    ws["G18"].alignment = Alignment(horizontal="distributed", wrap_text=True)
    found.extend(
        Case("align", f"distributed {label}", cell, ("lines", "words", "ink"))
        for cell, label in (
            ("C14", "words"),
            ("E14", "one word"),
            ("G14", "long unwrapped"),
            ("C16", "indent 2"),
            ("G16", "wrapped justifyLastLine"),
            ("G18", "wrapped"),
        )
    )
    for column, indent in (("C", 0), ("E", 1), ("G", 3)):
        ws[f"{column}22"] = "HHH"
        ws[f"{column}22"].alignment = Alignment(horizontal="right", indent=indent)
    ws["I22"] = 123
    ws["I22"].alignment = Alignment(horizontal="right", indent=2)
    ws["G24"] = 123
    ws["G24"].alignment = Alignment(indent=2)
    ws["C24"] = "HHH"
    ws["C24"].alignment = Alignment(horizontal="left", indent=3)
    found.extend(
        Case("align", f"indent {label}", cell, ("ink",))
        for cell, label in (
            ("C22", "right 0"),
            ("E22", "right 1"),
            ("G22", "right 3"),
            ("I22", "number right 2"),
            ("C24", "left 3"),
            ("G24", "number general 2"),
        )
    )
    for column, text, order, label in (
        ("C", "abc!", 0, "context latin"),
        ("E", "abc!", 2, "right-to-left latin"),
        ("G", HEBREW, 0, "context hebrew"),
        ("I", HEBREW, 1, "left-to-right hebrew"),
    ):
        ws[f"{column}26"] = text
        ws[f"{column}26"].alignment = Alignment(readingOrder=order)
        found.append(Case("align", f"reading order {label}", f"{column}26", ("words", "ink")))
    return found


def _shrink(ws: Any) -> list[Case]:
    ws.column_dimensions["G"].width = 60
    found = []
    for row in range(4, 23):
        ws[f"C{row}"] = ws[f"G{row}"] = "H" * row
        ws[f"C{row}"].alignment = Alignment(horizontal="left", shrink_to_fit=True)
        ws[f"G{row}"].alignment = Alignment(horizontal="left")
        found.append(Case("shrink", f"scale {row}", f"C{row}", ("scale", "ink"), None, f"G{row}"))
    ws["E4"] = "shrink and wrap both set here"
    ws["E4"].alignment = Alignment(wrap_text=True, shrink_to_fit=True)
    ws["E6"] = 123456789012
    ws["E6"].number_format = "0"
    ws["E6"].alignment = Alignment(shrink_to_fit=True)
    for cell, options in (
        ("E8", {"horizontal": "left", "indent": 2}),
        ("E10", {"horizontal": "right"}),
        ("E12", {"horizontal": "left"}),
    ):
        ws[cell] = "H" * 12
        ws[cell].alignment = Alignment(shrink_to_fit=True, **options)
    ws["E12"].font = Font(name="Calibri", size=20)
    ws["E14"] = "a long text merged and shrunk"
    ws.merge_cells("E14:F14")
    ws["E14"].alignment = Alignment(shrink_to_fit=True)
    found.extend(
        Case("shrink", label, cell, ("lines", "ink"), last)
        for cell, label, last in (
            ("E4", "wrapped", None),
            ("E6", "number", None),
            ("E8", "left indent 2", None),
            ("E10", "right", None),
            ("E12", "20 pt", None),
            ("E14", "merged", "F14"),
        )
    )
    return found


def _fonts(ws: Any) -> list[Case]:
    for column, width in (("C", 14), ("E", 14), ("G", 30), ("I", 14)):
        ws.column_dimensions[column].width = width
    found = []
    for row, underline in UNDERLINES.items():
        ws[f"C{row}"] = "HHH"
        ws[f"E{row}"] = 123
        ws[f"G{row}"] = "HH"
        for column, label in (("C", "text"), ("E", "number"), ("G", "wide cell")):
            ws[f"{column}{row}"].font = Font(name="Calibri", size=11, underline=underline)
            name = f"underline {underline or 'none'} {label}"
            found.append(Case("fonts", name, f"{column}{row}", ("rules", "ink")))
    for row, size in ((14, 11), (16, 24)):
        for column, position in (("C", None), ("E", "superscript"), ("G", "subscript")):
            ws[f"{column}{row}"] = "HHH"
            ws[f"{column}{row}"].font = Font(name="Calibri", size=size, vertAlign=position)
            name = f"{position or 'baseline'} {size}"
            found.append(Case("fonts", name, f"{column}{row}", ("rules", "ink")))
    for row, position in ((18, "superscript"), (20, "subscript")):
        ws[f"C{row}"] = "Hxg"
        ws[f"C{row}"].font = Font(name="Calibri", size=11, vertAlign=position)
        found.append(Case("fonts", f"{position} alone", f"C{row}", ("ink",)))
    ws["G18"] = "HHH"
    ws["G18"].font = Font(name="Calibri", size=11, strike=True)
    found.append(Case("fonts", "strike", "G18", ("rules", "ink")))
    ws["C22"] = ws["G22"] = SAMPLE
    ws["C22"].font = Font(name=MISSING, size=11)
    ws["G22"].font = Font(name="Calibri", size=11)
    found.append(Case("fonts", "missing font", "C22", ("ink",), "F22"))
    found.append(Case("fonts", "missing font reference", "G22", ("ink",)))
    for column, options, label in (
        ("C", {"outline": True}, "outline"),
        ("E", {"shadow": True}, "shadow"),
        ("G", {}, "plain 24"),
        ("I", {"condense": True, "extend": True}, "condense extend"),
    ):
        ws[f"{column}24"] = "HHH"
        ws[f"{column}24"].font = Font(name="Calibri", size=24, **options)
        found.append(Case("fonts", label, f"{column}24", ("ink",)))
    return found


# Two words whose widths with the space between them are the key, in points at Calibri 11 by
# xlformat's tables (whole points at that size): where wrapping breaks them tells the width a line
# may take. Around 48 for a default column (53 pt less the insets), around 115 for 20 characters.
NARROW = {
    44: "aaaa baaii", 45: "aaaa aaaai", 46: "aaaa baaai", 47: "aaaa aaaaa",
    48: "aaaa aaaaii", 49: "aaaa baaaii", 50: "aaaa aaaaai",
}  # fmt: skip
WIDE = {
    111: "aaaaaaaaaa bbbbbbaaaai", 112: "aaaaaaaaaa bbbbbbbaaai", 113: "aaaaaaaaaa bbbbbaaaaaii",
    114: "aaaaaaaaaa bbbbbbaaaaii", 115: "aaaaaaaaaa bbbbbbbaaaii", 116: "aaaaaaaaaa bbbbbbaaaaai",
    117: "aaaaaaaaaa bbbbbbbaaaai",
}  # fmt: skip
WORDS = {
    45: "bbbbaaaii", 46: "bbbaaaaai", 47: "bbbbaaaai", 48: "bbbbbaaai", 49: "bbbaaaaaii",
    50: "bbbbaaaaii",
}  # fmt: skip
"""Single words of these widths: a word wider than a line is broken between characters."""
SPILLED = "spilled text that is much wider than one column"


def _wrap(ws: Any) -> list[Case]:
    for column in "GI":
        ws.column_dimensions[column].width = 20
    found = []
    shrink = Alignment(wrap_text=True, shrink_to_fit=True)
    for k, (narrow, wide) in enumerate(zip(NARROW.items(), WIDE.items(), strict=True)):
        row = 4 + 2 * k
        for column, (width, text), options, label in (
            ("C", narrow, Alignment(wrap_text=True), "wrap"),
            ("E", narrow, shrink, "shrink wrap"),
            ("G", wide, Alignment(wrap_text=True), "wrap"),
            ("I", wide, shrink, "shrink wrap"),
        ):
            ws[f"{column}{row}"] = text
            ws[f"{column}{row}"].alignment = options
            found.append(Case("wrap", f"{label} {width}", f"{column}{row}", ("lines", "ink")))
    for k, (width, word) in enumerate(WORDS.items()):
        row = 18 + 2 * k
        for column, options, label in (
            ("C", Alignment(wrap_text=True), "word"),
            ("E", shrink, "shrink word"),
        ):
            ws[f"{column}{row}"] = word
            ws[f"{column}{row}"].alignment = options
            found.append(Case("wrap", f"{label} {width}", f"{column}{row}", ("lines", "ink")))
    for column, horizontal in (("C", "distributed"), ("E", "justify")):
        ws[f"{column}30"] = NARROW[47]
        ws[f"{column}30"].alignment = Alignment(horizontal=horizontal, wrap_text=True)
        found.append(Case("wrap", f"{horizontal} 47", f"{column}30", ("lines", "ink")))
    ws["G30"] = "shrink and wrap both set here"
    ws["G30"].alignment = Alignment(wrap_text=True)
    found.append(Case("wrap", "wrap only, the shrink text", "G30", ("lines", "ink")))
    return found


def _lone(ws: Any) -> list[Case]:
    found = []
    for column, text, options, label in (
        ("C", SPILLED, {"horizontal": "distributed"}, "distributed"),
        ("E", SPILLED, {"horizontal": "distributed", "justifyLastLine": True}, "last line"),
        ("G", SPILLED, {"horizontal": "justify"}, "justify"),
        ("I", SPILLED, {"horizontal": "distributed", "indent": 1}, "distributed indent 1"),
    ):
        ws[f"{column}4"] = text
        ws[f"{column}4"].alignment = Alignment(wrap_text=True, **options)
        found.append(Case("lone", f"{label} fitted", f"{column}4", ("lines", "ink", "line_inks")))
    ws.row_dimensions[6].height = 15
    for column, settings, label in (
        ("C", {"horizontal": "distributed"}, "distributed 15"),
        ("E", {"horizontal": "distributed", "wrap_text": True}, "distributed wrapped 15"),
        ("G", {"wrap_text": True}, "wrapped 15"),
        ("I", {"horizontal": "distributed", "vertical": "top"}, "distributed top 15"),
    ):
        ws[f"{column}6"] = SPILLED
        ws[f"{column}6"].alignment = Alignment(**settings)
        found.append(Case("lone", label, f"{column}6", ("lines", "ink", "line_inks")))
    for column, text in (
        ("C", "spilled"),
        ("E", "abc"),
        ("G", "aa bb cccccccc dd ee"),
        ("I", "aaaa bbbb cccc"),
    ):
        ws[f"{column}8"] = text
        ws[f"{column}8"].alignment = Alignment(horizontal="distributed", wrap_text=True)
        found.append(
            Case("lone", f"distributed {text}", f"{column}8", ("lines", "ink", "line_inks"))
        )
    return found


FILLS = {
    "align": _align, "shrink": _shrink, "fonts": _fonts, "wrap": _wrap, "lone": _lone,
}  # fmt: skip


def fill(ws: Any, group: str) -> list[Case]:
    """Lay one group out on a worksheet; its cases."""
    return FILLS[group](ws)


def cases(group: str | None = None) -> list[Case]:
    """Every case of ``group``, or of every group, in order."""
    chosen = GROUPS if group is None else (group,)
    return [case for g in chosen for case in fill(openpyxl.Workbook().active, g)]


def _cells(found: list[Case]) -> list[str]:
    """The cells whose geometry a group's measuring needs."""
    cells = [c for case in found for c in (case.cell, case.last, case.reference) if c]
    return list(dict.fromkeys(cells))


def build(path: Path) -> list[Case]:
    """Write each group to its own worksheet of one workbook at ``path``; every case."""
    wb = openpyxl.Workbook()
    found = []
    for k, group in enumerate(GROUPS):
        ws = wb.active if k == 0 else wb.create_sheet()
        ws.title = group.capitalize()
        found.extend(fill(ws, group))
    wb.save(path)
    return found


def run(groups: tuple[str, ...] = GROUPS) -> None:
    """Capture each group's window at its zooms, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    for group in groups:
        capture(
            lambda ws, group=group: fill(ws, group),
            cells=_cells(cases(group)),
            out=BUILD_DIR / group,
            zooms=ZOOMS[group],
            name=f"text_style_{group}",
        )
    write()


def _span(cap: Any, start: float, length: float, origin: float) -> str:
    from capture.screen.ink import pt

    return f"{pt(cap, start - origin)}+{pt(cap, length)}"


def _words(cap: Any, box: tuple[int, int, int, int], left: float) -> str:
    """Ink clusters of columns (gaps under 2 px joined), from the cell's left edge."""
    from capture.screen.ink import bands

    merged: list[tuple[int, int]] = []
    for x, n in bands(cap, box, axis="x"):
        if merged and x - (merged[-1][0] + merged[-1][1]) < 2:
            merged[-1] = (merged[-1][0], x + n - merged[-1][0])
        else:
            merged.append((x, n))
    return " ".join(_span(cap, x, n, left) for x, n in merged)


def _rules(
    cap: Any, box: tuple[int, int, int, int], glyph: tuple[int, int, int, int], edges: Any
) -> dict[str, str]:
    """The baseline and the horizontal rules (underlines, strike) under or through the glyphs."""
    from capture.screen.ink import bands, pt, rule_rows

    left, right, top = edges
    rules = rule_rows(cap, box, max(4, int(0.6 * (glyph[2] - glyph[0]))))
    ruled = {y for y0, h, _, _ in rules for y in range(y0, y0 + h)}
    glyph_rows = [y for y0, n in bands(cap, box) for y in range(y0, y0 + n) if y not in ruled]
    base = max(glyph_rows) + 1 if glyph_rows else glyph[3]
    found = {"baseline": pt(cap, base - top)}
    if glyph_rows:
        found["glyph_top"] = pt(cap, base - min(glyph_rows))
    found["rule_offset"] = " ".join(pt(cap, y - base) for y, _, _, _ in rules)
    found["rule_thickness"] = " ".join(pt(cap, h) for _, h, _, _ in rules)
    found["rule_start"] = " ".join(pt(cap, xs - left) for _, _, xs, _ in rules)
    found["rule_end"] = " ".join(pt(cap, xe - right) for _, _, _, xe in rules)
    return found


def measure(cap: Any, case: Case) -> dict[str, str] | None:
    """One case's measurements at one zoom, or ``None`` when it is outside the window."""
    from capture.screen.ink import bands, cells_box, ink, pt, visible

    end = case.last or case.cell
    shown = [case.cell, end, *([case.reference] if case.reference else [])]
    if not all(visible(cap, cell) for cell in shown):
        return None
    left, top = cap.edge(case.cell, "left"), cap.edge(case.cell, "top")
    right = cap.edge(end, "right")
    box = cells_box(cap, case.cell, case.last, inset=1)
    glyph = ink(cap, box)
    found: dict[str, str] = {}
    if glyph is not None and "ink" in case.measure:
        found["ink_left"] = pt(cap, glyph[0] - left)
        found["ink_right"] = pt(cap, glyph[2] - right)
        found["ink_top"] = pt(cap, glyph[1] - top)
        found["ink_bottom"] = pt(cap, glyph[3] - top)
    if "lines" in case.measure:
        # bands under a screen point apart are one line: an "i"'s dot stands half a point
        # above its stem (text_style_more's indent cases)
        lines: list[tuple[int, int]] = []
        for y, n in bands(cap, box):
            if lines and y - (lines[-1][0] + lines[-1][1]) < cap.scale:
                lines[-1] = (lines[-1][0], y + n - lines[-1][0])
            else:
                lines.append((y, n))
        found["lines"] = str(len(lines))
        found["line_tops"] = " ".join(_span(cap, y, n, top) for y, n in lines)
    if "words" in case.measure:
        found["words"] = _words(cap, box, left)
    if "line_inks" in case.measure:
        spans = []
        for y, n in bands(cap, box):
            line = ink(cap, (box[0], y, box[2], y + n))
            if line is not None:
                spans.append(_span(cap, line[0], line[2] - line[0], left))
        found["line_inks"] = " ".join(spans)
    if glyph is not None and "rules" in case.measure:
        found.update(_rules(cap, box, glyph, (left, right, top)))
    if glyph is not None and case.reference:
        other = ink(cap, cells_box(cap, case.reference, inset=1))
        if other is not None:
            found["width_ratio"] = f"{round((glyph[2] - glyph[0]) / (other[2] - other[0]), 4):g}"
            found["height_ratio"] = f"{round((glyph[3] - glyph[1]) / (other[3] - other[1]), 4):g}"
            found["ink_width"] = pt(cap, glyph[2] - glyph[0])
            found["ref_width"] = pt(cap, other[2] - other[0])
    return found


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``text_style.tsv``."""
    from capture.screen.run import header, load

    records = []
    first: dict[str, str] = {}
    for group in GROUPS:
        out = BUILD_DIR / group
        if not (out / "formats.tsv").exists():
            print(f"{group}: not captured yet, left out", file=sys.stderr)
            continue
        first = first or header(out)
        for cap in load(out):
            for case in cases(group):
                found = measure(cap, case)
                if found is None:
                    continue
                reads = {key: cap.value(case.cell, read) or "" for key, read in READS.items()}
                records.append(
                    {
                        "group": case.group,
                        "name": case.name,
                        "zoom": str(cap.zoom),
                        "cell": case.cell,
                        **reads,
                        **found,
                    }
                )
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after", "zooms"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: text_style\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cases in {argv[1]}")
        return 0
    if argv[:1] == ["run"] and all(group in GROUPS for group in argv[1:]):
        run(tuple(argv[1:]) or GROUPS)
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
