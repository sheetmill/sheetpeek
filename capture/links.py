"""Capture how Excel for Mac draws and follows hyperlinks, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.links run
    uv run --group capture python -m capture.links read           # the saved captures again
    uv run --group capture python -m capture.links build PATH     # the workbook only

ARCHITECTURE.md §4.10, the rules it settles: a link has no look of its own
(the Normal style draws a linked cell as plain text; the Hyperlink and Followed Hyperlink cell
styles draw the theme's ``hlink`` and ``folHlink`` with a single underline, with or without a
link); ``display`` is not drawn and an empty linked cell shows nothing; a ``HYPERLINK()`` cell
from a file is drawn as the value it shows on opening, in its own format; which target Excel
takes for a cell with both a ``<hyperlink>`` element and a ``HYPERLINK()`` formula; and that a
``location`` without a sheet is the link's own sheet.

One workbook of three sheets, ``Links`` (first and active), ``Second`` and ``Third``. Its theme's
``hlink`` is green ``008000`` and its ``folHlink`` orange ``C05000`` (``postprocess`` rewrites
openpyxl's theme), so a drawn link color is the file's theme and not a fixed blue. The cell
styles Hyperlink (builtin 8) and Followed Hyperlink (9) are written as Excel writes them:
Calibri 11, theme color 10 or 11, single underline. Every link points into the workbook, so
following one opens nothing else. On ``Links``, each case is one cell of 18 characters, with
2-character columns between and an empty row under each case row (:data:`CASES`); the merge
case spans C12:E12. Every formula is written without a cached value, so Excel computes it on
opening (openpyxl writes ``fullCalcOnLoad``), as sheetpeek does (§5.7).

Two steps, each with the workbook open on its own:

- ``look``: one window of ``Links`` at zoom 100 (markers B2 and I14). Measured per case, over
  the cell (or the merge), 4 image pixels in from its sides and 1 from its top and bottom: the
  ink, pixels more than 100 from the cell's most common color, so the gridline's gray is not
  ink (``ink_pixels``; ``ink_color``, the mean of the tenth of them farthest from it), its box
  in points from the cell's left and top (``ink_left`` to ``ink_bottom``), and the horizontal
  rules at least 60 % of the ink's width (``rules``, the count; ``rule_top`` and
  ``rule_thickness`` of the lowest, in points from the cell's top), which are underlines; and
  Excel's own reads: ``display format``'s font color and underline, the cell's own, and its
  string value (``text``).
- ``follow``: no window. For each case of :data:`FOLLOWS`, a sheet is made active, the cell's
  hyperlinks as Excel holds them are read (their count, the first one's address and sub
  address, the cell's formula), the first one is followed (the hyperlink's own ``follow``;
  the workbook's ``follow hyperlink`` takes an address, not a link), and the active sheet and
  cell after it are read. This settles where a location without a sheet goes (from
  its own sheet and from another), whether Excel keeps a ``HYPERLINK()`` cell's link as a
  hyperlink object, and, for a cell with both, which of the two Excel keeps as its link.

``read`` writes ``tests/captures/links.tsv``: one row per look case (``kind`` ``look``) and per
follow case (``kind`` ``follow``).
"""

from __future__ import annotations

import os
import re
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Color, Font, NamedStyle
from openpyxl.worksheet.hyperlink import Hyperlink

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "links"
OUTPUT = ROOT / "tests" / "captures" / "links.tsv"
ZOOM = 100
MARKERS = ("B2", "I14")
SHEETS = ("Links", "Second", "Third")
HLINK, FOLHLINK = "008000", "C05000"
"""The theme's link colors, unlike any default theme's and far from the marker's magenta."""
SLOTS = ("C", "E", "G")
INSET = 4
"""Image pixels kept off a cell's left and right edges, where its gridlines are."""
INSET_Y = 1
"""Image pixels kept off its top and bottom: an underline in a 15 pt row ends about 1 pt above
the cell's bottom (the text_style suite), inside a 4 pixel inset."""
INK = 100
"""Distance from the cell's color past which a pixel is ink: the gridline's gray (``D4D4D4``,
74 from white) is not, so a gridline at the top or bottom is neither ink nor a rule."""
RULE_SHARE = 0.6
"""A horizontal rule is an underline when it spans this share of the ink's width at least."""
TEXT = "Hyperlink text"
COLUMNS = (
    "kind", "name", "cell", "zoom", "text", "ink_pixels", "ink_color", "ink_left", "ink_right",
    "ink_top", "ink_bottom", "rules", "rule_top", "rule_thickness", "display_font_color",
    "display_underline", "own_font_color", "own_underline", "active", "sheet", "links_count",
    "formula", "link_address", "link_sub_address", "followed", "after_sheet", "after_cell",
)  # fmt: skip
READS = {
    "text": "geometry.text",
    "display_font_color": "display.font_color",
    "display_underline": "display.underline",
    "own_font_color": "own.font_color",
    "own_underline": "own.underline",
}
FOLLOW_READS = {
    "active.before": "active",
    "links.count": "links_count",
    "formula": "formula",
    "link.address": "link_address",
    "link.sub_address": "link_sub_address",
    "followed": "followed",
    "after.sheet": "after_sheet",
    "after.cell": "after_cell",
}


@dataclass(frozen=True)
class Case:
    """One cell of ``Links`` measured on the screen; ``last`` ends a merge."""

    name: str
    cell: str
    last: str | None = None


CASES = [
    Case("Normal, linked", "C4"),
    Case("Hyperlink style, linked", "E4"),
    Case("Followed Hyperlink style, linked", "G4"),
    Case("Hyperlink style, no link", "C6"),
    Case("Followed Hyperlink style, no link", "E6"),
    Case("Hyperlink style, own plain font, linked", "G6"),
    Case("Normal, display unlike the value", "C8"),
    Case("Hyperlink style, display unlike the value", "E8"),
    Case("empty, linked with a display", "G8"),
    Case("HYPERLINK() with a friendly name", "C10"),
    Case("HYPERLINK() without a friendly name", "E10"),
    Case("element and HYPERLINK()", "G10"),
    Case("empty merge, linked with a display", "C12", "E12"),
]

EXPECTED = {
    "Normal, linked": (TEXT, "000000", "none"),
    "Hyperlink style, linked": (TEXT, HLINK, "single"),
    "Followed Hyperlink style, linked": (TEXT, FOLHLINK, "single"),
    "Hyperlink style, no link": (TEXT, HLINK, "single"),
    "Followed Hyperlink style, no link": (TEXT, FOLHLINK, "single"),
    "Hyperlink style, own plain font, linked": (TEXT, "000000", "none"),
    "Normal, display unlike the value": ("v", "000000", "none"),
    "Hyperlink style, display unlike the value": ("v", HLINK, "single"),
    "empty, linked with a display": ("", "000000", "none"),
    "HYPERLINK() with a friendly name": ("Friendly name", "000000", "none"),
    "HYPERLINK() without a friendly name": ("#Third!B8", "000000", "none"),
    "element and HYPERLINK()": ("Formula text", "000000", "none"),
    "empty merge, linked with a display": ("", "000000", "none"),
}
"""What sheetpeek draws for each case, by name: the text, the font color and the underline."""


@dataclass(frozen=True)
class Follow:
    """A link followed with sheet ``active`` active: the link on ``sheet``'s ``cell``."""

    name: str
    active: str
    sheet: str
    cell: str


FOLLOWS = [
    Follow("location with a sheet (control)", "Links", "Links", "C4"),
    Follow("location without a sheet, its own sheet active", "Second", "Second", "B2"),
    Follow("location without a sheet, the first sheet active", "Links", "Second", "B2"),
    Follow("HYPERLINK() alone", "Links", "Links", "C10"),
    Follow("element and HYPERLINK()", "Links", "Links", "G10"),
]

FOLLOW_EXPECTED: dict[str, str | None] = {
    "location with a sheet (control)": "Third!E6",
    "location without a sheet, its own sheet active": "Second!B5",
    "location without a sheet, the first sheet active": "Second!B5",
    "HYPERLINK() alone": None,
    "element and HYPERLINK()": "Second!C3",
}
"""The place sheetpeek takes each link to (``Sheet!A1``), or ``None`` for no place in the
workbook: sheetpeek reads a ``HYPERLINK()`` target ``#Third!B7`` as a relative target."""


# -- the workbook --------------------------------------------------------------------------------


def _styles() -> tuple[NamedStyle, NamedStyle]:
    """The Hyperlink and Followed Hyperlink cell styles, as Excel writes them."""

    def font(theme: int) -> Font:
        return Font(
            name="Calibri", size=11, family=2, scheme="minor", underline="single",
            color=Color(theme=theme),
        )  # fmt: skip

    return (
        NamedStyle("Hyperlink", font=font(10), builtinId=8),
        NamedStyle("Followed Hyperlink", font=font(11), builtinId=9),
    )


def _link(ws: Any, cell: str, location: str, display: str | None = None) -> None:
    """A ``<hyperlink>`` into the workbook on ``cell``; the cell's value left as it was."""
    value = ws[cell].value
    ws[cell].hyperlink = Hyperlink(ref=cell, location=location, display=display)
    ws[cell].value = value  # openpyxl fills an empty linked cell with the location


def fill(ws: Any) -> list[Case]:
    """Lay the cases out on the first worksheet, and the other sheets beside it; the cases."""
    wb = ws.parent
    ws.title = "Links"
    second, third = wb.create_sheet("Second"), wb.create_sheet("Third")
    hyperlink, followed = _styles()
    for column in SLOTS:
        ws.column_dimensions[column].width = 18
    for column in ("D", "F"):
        ws.column_dimensions[column].width = 2
    for case in CASES[:6] + CASES[7:8]:
        ws[case.cell] = TEXT
    ws["C8"] = ws["E8"] = "v"
    for cell in ("E4", "C6", "E8"):
        ws[cell].style = hyperlink
    for cell in ("G4", "E6"):
        ws[cell].style = followed
    ws["G6"].style = hyperlink
    ws["G6"].font = Font(name="Calibri", size=11, family=2, scheme="minor", color=Color(theme=1))
    _link(ws, "C4", "Third!E6")
    for cell in ("E4", "G4", "G6"):
        _link(ws, cell, "Links!A1")
    _link(ws, "C8", "Links!A1", "display text much longer than the value")
    _link(ws, "E8", "Links!A1", "display text much longer than the value")
    ws["C10"] = '=HYPERLINK("#Third!B7","Friendly name")'
    ws["E10"] = '=HYPERLINK("#Third!B8")'
    ws["G10"] = '=HYPERLINK("#Third!D4","Formula text")'
    _link(ws, "G10", "Second!C3")
    ws.merge_cells("C12:E12")
    second["B2"] = "to B5"
    _link(second, "B2", "B5")
    third["A1"] = "third"
    return list(CASES)


def theme(text: str) -> str:
    """The theme with :data:`HLINK` and :data:`FOLHLINK` as its link colors."""
    for name, color in (("hlink", HLINK), ("folHlink", FOLHLINK)):
        text, count = re.subn(
            rf"(<a:{name}>\s*<a:srgbClr val=\")[0-9A-Fa-f]{{6}}", rf"\g<1>{color}", text
        )
        if count != 1:
            raise ValueError(f"the theme's {name} was not found")
    return text


EMPTY_LINKS = (
    '<hyperlink ref="G8" location="Links!A1" display="ghost"/>'
    '<hyperlink ref="C12:E12" location="Links!A1" display="ghost"/>'
)
"""The links on empty cells, which openpyxl does not write (it writes no empty cell)."""


def sheet(text: str) -> str:
    """The first sheet with the links on empty cells added."""
    if text.count("</hyperlinks>") != 1:
        raise ValueError("the first sheet has no <hyperlinks> element")
    return text.replace("</hyperlinks>", EMPTY_LINKS + "</hyperlinks>")


def postprocess(path: Path) -> None:
    """Set the theme's link colors and add the links on empty cells in the saved workbook."""
    from capture.screen.run import patch_part

    patch_part(path, {"xl/theme/theme1.xml": theme, "xl/worksheets/sheet1.xml": sheet})


def build(path: Path) -> list[Case]:
    """Write the workbook to ``path``; its cases."""
    wb = openpyxl.Workbook()
    placed = fill(wb.active)
    wb.save(path)
    postprocess(path)
    return placed


def _cells() -> list[str]:
    return [cell for case in CASES for cell in (case.cell, case.last) if cell]


def run() -> None:
    """Capture the look, then follow the links, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    capture(
        fill,
        cells=_cells(),
        out=BUILD_DIR / "look",
        markers=MARKERS,
        zooms=(ZOOM,),
        name="links_look",
        postprocess=postprocess,
    )
    follow()
    write()


def follow() -> None:
    """Open the workbook alone and follow each of :data:`FOLLOWS` (needs Excel, no window)."""
    from capture.screen.run import write_tsv
    from capture.screen.session import CONTAINER, ExcelSession, environment

    edition = os.environ.get("EXCEL_EDITION")
    if not edition:
        raise SystemExit("state the Excel edition (Excel > About Excel) in EXCEL_EDITION")
    out = BUILD_DIR / "follow"
    out.mkdir(parents=True, exist_ok=True)
    (CONTAINER / "sheetpeek-capture").mkdir(parents=True, exist_ok=True)
    path = CONTAINER / "sheetpeek-capture" / f"links_follow_{os.getpid()}.xlsx"
    found = environment(edition)
    session = ExcelSession(path, evidence=out / "excel_timeouts")
    rows: list[tuple[str, str, str]] = []
    try:
        build(path)
        with session:
            for case in FOLLOWS:
                rows.extend(
                    (case.name, key, value)
                    for key, value in session.follow(case.active, case.sheet, case.cell)
                )
    finally:
        path.unlink(missing_ok=True)
    found["suite"] = "links_follow"
    found["workbooks_open_after"] = session.open_after or "unknown"
    write_tsv(out / "formats.tsv", found, rows)


# -- measuring -----------------------------------------------------------------------------------


def _distance(a: tuple[int, int, int], b: tuple[int, int, int]) -> float:
    return float(sum((x - y) ** 2 for x, y in zip(a, b, strict=True)) ** 0.5)


def measure(cap: Any, case: Case) -> dict[str, str]:
    """One case's ink and underlines at one zoom."""
    from capture.screen.ink import cells_box, rule_rows

    x0, y0, x1, y1 = cells_box(cap, case.cell, case.last)
    x0, y0, x1, y1 = x0 + INSET, y0 + INSET_Y, x1 - INSET, y1 - INSET_Y
    pixels = [cap.image.pixel(x, y) for y in range(y0, y1) for x in range(x0, x1)]
    paper = Counter(pixels).most_common(1)[0][0]

    def inked(rgb: tuple[int, int, int]) -> bool:
        return _distance(rgb, paper) > INK

    ink = [p for p in pixels if inked(p)]
    if not ink:
        return {"ink_pixels": "0", "ink_color": "none", "rules": "0"}
    core = sorted(ink, key=lambda p: -_distance(p, paper))[: max(1, len(ink) // 10)]
    mean = [round(sum(p[i] for p in core) / len(core)) for i in range(3)]
    xs = [x for y in range(y0, y1) for x in range(x0, x1) if inked(cap.image.pixel(x, y))]
    ys = [y for y in range(y0, y1) for x in range(x0, x1) if inked(cap.image.pixel(x, y))]
    left, top = cap.edge(case.cell, "left"), cap.edge(case.cell, "top")
    per = cap.scale

    def spt(pixels: float) -> str:
        return f"{round(pixels / per, 2):g}"

    found = {
        "ink_pixels": str(len(ink)),
        "ink_color": "".join(f"{c:02X}" for c in mean),
        "ink_left": spt(min(xs) - left),
        "ink_right": spt(max(xs) + 1 - left),
        "ink_top": spt(min(ys) - top),
        "ink_bottom": spt(max(ys) + 1 - top),
    }
    width = max(xs) + 1 - min(xs)
    rules = rule_rows(cap, (x0, y0, x1, y1), max(2, round(RULE_SHARE * width)), inked)
    found["rules"] = str(len(rules))
    if rules:
        y, height, _, _ = rules[-1]
        found |= {"rule_top": spt(y - top), "rule_thickness": spt(height)}
    return found


def _header_lines(first: dict[str, str]) -> str:
    return "".join(
        f"# {key}: {value}\n"
        for key, value in first.items()
        if key not in ("markers", "suite", "workbooks_open_after")
    )


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``links.tsv``."""
    from capture.screen.run import header, load

    records: list[dict[str, str]] = []
    first: dict[str, str] = {}
    look = BUILD_DIR / "look"
    if (look / "formats.tsv").exists():
        first = header(look)
        for cap in load(look):
            for case in CASES:
                reads = {key: cap.value(case.cell, read) or "" for key, read in READS.items()}
                records.append(
                    {"kind": "look", "name": case.name, "cell": case.cell, "zoom": str(cap.zoom)}
                    | measure(cap, case)
                    | reads
                )
    else:
        print("look: not captured yet, left out", file=sys.stderr)
    out = BUILD_DIR / "follow"
    if (out / "formats.tsv").exists():
        first = first or header(out)
        found: dict[str, dict[str, str]] = {}
        for line in (out / "formats.tsv").read_text(encoding="utf-8").splitlines():
            if line.startswith("# ") or line.startswith("cell\t"):
                continue
            name, key, value = [*line.split("\t"), "", ""][:3]
            if key in FOLLOW_READS:
                found.setdefault(name, {})[FOLLOW_READS[key]] = value
        records.extend(
            {"kind": "follow", "name": case.name, "cell": case.cell, "sheet": case.sheet}
            | found.get(case.name, {})
            for case in FOLLOWS
        )
    else:
        print("follow: not captured yet, left out", file=sys.stderr)
    with open(OUTPUT, "w", encoding="utf-8") as file:
        file.write(_header_lines(first))
        file.write("# suite: links\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            file.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cases in {argv[1]}")
        return 0
    if argv == ["run"]:
        run()
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
