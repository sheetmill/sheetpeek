"""Capture where Excel stops honoring a workbook's fonts (route C, the screen).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.screen run font_cut
    uv run --group capture python -m capture.screen read font_cut     # the saved captures again
    uv run --group capture python -m capture.font_cut build DIR       # the workbooks only

Excel keeps one font table per workbook: ``styles.xml``'s font list in order, then each rich-text
run's inline font. Past an entry of it, a cell is drawn and measured in the default font while
Excel's object model reports its own. Two groups, each its own workbook and window at zoom 100,
under openpyxl's Calibri 11 Normal font (font 0):

- ``list``: 989 filler fonts (Calibri 11, each a different near-black color) on a second sheet
  take entries 1 to 989; then "Hxg" in Arial 24, each probe a slightly different red, in C4 to
  C23, takes entries 990 to 1,009, in row order (``LIST_FIRST``).
- ``rich``: 994 filler fonts take entries 1 to 994; then C4 to C13 hold "Hxg" as one rich-text
  run each, in Arial 24 and a red of its own, whose inline fonts follow the list in cell order:
  entries 995 to 1,004 (``RICH_FIRST``). All on one sheet: the order of runs spread over several
  sheets is not asked.

Rows are fitted by Excel on opening (no height given). Measured per probe: the row's height and
the font Excel's object model reports for the cell (its reads), and in the window the ink of the
cell: its top and bottom in points from the cell's top, and how many of its pixels are red (the
probe's own font) and how many dark gray or black (``red`` and ``dark``). ``read`` writes
``tests/captures/font_cut.tsv``: one row per probe.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font
from openpyxl.styles.colors import Color

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "font_cut"
OUTPUT = ROOT / "tests" / "captures" / "font_cut.tsv"
GROUPS = ("list", "rich")
LIST_FIRST = 990
"""``list``: the font table entry of the first probe (C4); C23 is entry 1,009."""
LIST_PROBES = 20
RICH_FIRST = 995
"""``rich``: the table entry of the first rich-text run (C4); C13's is 1,004."""
RICH_PROBES = 10
FIRST_ROW = 4
MARKERS = ("B2", "E3")
ZOOMS = (100,)
TEXT = "Hxg"
FIELDS = (
    "group", "entry", "cell", "row_height", "font_name", "font_size", "font_color", "ink_top",
    "ink_bottom", "red", "dark",
)  # fmt: skip


def _filler(ws: Any, count: int) -> None:
    """``count`` distinct fonts, entries 1 to ``count``, on a sheet of their own."""
    sheet = ws.parent.create_sheet("filler")
    for k in range(1, count + 1):
        cell = sheet.cell(k, 1, "f")
        cell.font = Font(name="Calibri", size=11, color=f"FF{k // 256:02X}{k % 256:02X}00")


def _red(k: int) -> str:
    """A red of its own for probe ``k``: each probe font must be a font of its own."""
    return f"FF{255 - k:02X}0000"


def probes(group: str) -> list[tuple[str, int]]:
    """Each probe's cell and font table entry."""
    first, count = (LIST_FIRST, LIST_PROBES) if group == "list" else (RICH_FIRST, RICH_PROBES)
    return [(f"C{FIRST_ROW + k}", first + k) for k in range(count)]


def fill(ws: Any, group: str) -> list[tuple[str, int]]:
    """Lay one group out: its filler fonts first, then its probes; each probe's cell and entry."""
    ws.column_dimensions["C"].width = 14
    found = probes(group)
    if group == "list":
        _filler(ws, LIST_FIRST - 1)
        for k, (cell, _) in enumerate(found):
            ws[cell] = TEXT
            ws[cell].font = Font(name="Arial", size=24, color=_red(k))
    else:
        _filler(ws, RICH_FIRST - 1)
        for k, (cell, _) in enumerate(found):
            run = InlineFont(rFont="Arial", sz=24, color=Color(rgb=_red(k)))
            ws[cell] = CellRichText([TextBlock(run, TEXT)])
    return found


def build(directory: Path) -> dict[str, Path]:
    """Each group's workbook in ``directory``."""
    directory.mkdir(parents=True, exist_ok=True)
    paths = {}
    for group in GROUPS:
        wb = openpyxl.Workbook()
        fill(wb.active, group)
        paths[group] = directory / f"font_cut_{group}.xlsx"
        wb.save(paths[group])
    return paths


def references(group: str) -> list[str]:
    return [cell for cell, _ in probes(group)] + [
        f"A{FIRST_ROW + k}" for k in range(len(probes(group)))
    ]


def run(groups: tuple[str, ...] = GROUPS) -> None:
    """Capture each group's window, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    for group in groups:
        capture(
            lambda ws, group=group: fill(ws, group),
            cells=references(group),
            out=BUILD_DIR / group,
            markers=MARKERS,
            zooms=ZOOMS,
            name=f"font_cut_{group}",
        )
    write()


def _ink(cap: Any, cell: str) -> dict[str, str]:
    """The ink in ``cell``: its top and bottom, and its red and dark pixels."""
    from capture.screen.ink import content_box, pt

    x0, y0, x1, y1 = content_box(cap, cell)
    top = cap.edge(cell, "top")
    red = dark = 0
    rows = []
    for y in range(y0, y1):
        inked = False
        for x in range(x0, x1):
            r, g, b = cap.image.pixel(x, y)
            if r > 150 and g < 110 and b < 110:
                red += 1
                inked = True
            elif max(r, g, b) < 120:
                dark += 1
                inked = True
        if inked:
            rows.append(y)
    found = {"red": str(red), "dark": str(dark), "ink_top": "", "ink_bottom": ""}
    if rows:
        found["ink_top"] = pt(cap, rows[0] - top)
        found["ink_bottom"] = pt(cap, rows[-1] + 1 - top)
    return found


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``font_cut.tsv``."""
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
            for cell, entry in probes(group):
                row = cell[1:]
                records.append(
                    {
                        "group": group,
                        "entry": str(entry),
                        "cell": cell,
                        "row_height": cap.value(f"A{row}", "geometry.height") or "",
                        "font_name": cap.value(cell, "display.font_name") or "",
                        "font_size": cap.value(cell, "display.font_size") or "",
                        "font_color": cap.value(cell, "display.font_color") or "",
                        **_ink(cap, cell),
                    }
                )
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after", "zooms"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: font_cut\n" + "\t".join(FIELDS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in FIELDS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        print(build(Path(argv[1])))
        return 0
    if argv[:1] == ["run"] and all(g in GROUPS for g in argv[1:]):
        run(tuple(argv[1:]) or GROUPS)
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
