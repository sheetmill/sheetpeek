"""Capture where Excel stops honoring a workbook's fonts (route C, the screen).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.screen run font_cut
    uv run --group capture python -m capture.screen read font_cut     # the saved captures again
    uv run --group capture python -m capture.font_cut build DIR       # the workbooks only

Excel keeps one font table per workbook: ``styles.xml``'s font list in order, then each rich-text
run's inline font. Past an entry of it, a cell is drawn and measured in the default font while
Excel's object model reports its own. Seven groups (``GROUPS``), each its own workbook and
window at zoom 100, ``list`` and ``rich`` under openpyxl's Calibri 11 Normal font (font 0):

- ``list``: 989 filler fonts (Calibri 11, each a different near-black color) on a second sheet
  take entries 1 to 989; then "Hxg" in Arial 24, each probe a slightly different red, in C4 to
  C23, takes entries 990 to 1,009, in row order (``LIST_FIRST``).
- ``rich``: 994 filler fonts take entries 1 to 994; then C4 to C13 hold "Hxg" as one rich-text
  run each, in Arial 24 and a red of its own, whose inline fonts follow the list in cell order:
  entries 995 to 1,004 (``RICH_FIRST``). All on one sheet: the order of runs spread over several
  sheets is not asked.
- ``faces_calibri`` and ``faces_times``: openpyxl's theme, a Normal font of Calibri 11 or Times
  New Roman 10; "Hxg" in black in Arial, Aptos Narrow, Segoe UI, Tahoma and Calibri 11 (E4 to E8,
  list entries 1 to 5, ``REFERENCES``), then filler, then probes at entries 997 to 1,002 (C4 to
  C9): which face a cut cell takes on Excel for Windows, by its glyphs against the references',
  and whether the cut moves with the Normal font alone. ``faces_aptos`` the same under a theme of
  Aptos Narrow body and Aptos Display heading fonts with an Aptos Narrow 11 Normal font (a theme
  font): whether the cut moves with a Normal font other than the theme's body font or other than
  Calibri.
- ``theme_10`` and ``theme_16``: the theme's body font Arial and heading font Aptos Narrow
  (``THEME_FONTS``, for openpyxl's Calibri and Cambria, rewritten before Excel opens the file),
  a Normal font of Times New Roman 10 or 16 (``NORMALS``); 994 filler fonts, then the probes at
  list entries 995 to 1,004 in C4 to C13: which font a cut cell takes, the theme's body or heading
  font, the Normal font or Calibri, at the Normal font's size or Excel's standard size.

Rows are fitted by Excel on opening (no height given). Measured per probe: the row's height and
the font Excel's object model reports for the cell (its reads), and in the window the ink of the
cell: its top and bottom in points from the cell's top, and how many of its pixels are red (the
probe's own font) and how many dark gray or black (``red`` and ``dark``), and where each glyph's
ink starts (``glyphs``, points from the cell's inside: the face, by its advances). The run also
records Excel's standard font and size (``standard_font.applescript``). ``read`` writes
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
from openpyxl.utils.indexed_list import IndexedList

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "font_cut"
OUTPUT = ROOT / "tests" / "captures" / "font_cut.tsv"
GROUPS = ("list", "rich", "theme_10", "theme_16", "faces_calibri", "faces_times", "faces_aptos")
LIST_FIRST = 990
"""``list``: the font table entry of the first probe (C4); C23 is entry 1,009."""
LIST_PROBES = 20
RICH_FIRST = 995
"""``rich``: the table entry of the first rich-text run (C4); C13's is 1,004."""
RICH_PROBES = 10
THEME_FIRST = 995
"""``theme_<size>``: the font list entry of the first probe (C4); C13's is 1,004."""
THEME_PROBES = 10
THEME_FONTS = {"minor": "Arial", "major": "Aptos Narrow"}
"""``theme_<size>``: the theme's body (minor) and heading (major) fonts, for openpyxl's Calibri
and Cambria."""
NORMALS = {"theme_10": ("Times New Roman", 10), "theme_16": ("Times New Roman", 16)}
"""``theme_<size>``: the Normal font, neither theme font nor Calibri."""
FACES = {
    "faces_calibri": None,
    "faces_times": ("Times New Roman", 10),
    "faces_aptos": ("Aptos Narrow", 11),
}
"""``faces_<normal>``: the Normal font (``None``: openpyxl's Calibri 11), under openpyxl's theme;
``faces_aptos``' a theme font, its theme's body font (``APTOS_THEME``)."""
APTOS_THEME = {"minor": "Aptos Narrow", "major": "Aptos Display"}
"""``faces_aptos``' theme fonts, as Excel for Windows' default workbook has them."""
REFERENCES = (
    ("Arial", 11),
    ("Aptos Narrow", 11),
    ("Segoe UI", 11),
    ("Tahoma", 11),
    ("Calibri", 11),
)
"""``faces_<normal>``: "Hxg" in black in each, E4 to E8, list entries 1 to 5: the faces a cut cell
may be drawn in on Excel for Windows."""
FACES_FIRST = 997
"""``faces_<normal>``: the list entry of the first probe (C4); C9's is 1,002."""
FACES_PROBES = 6
STANDARD = BUILD_DIR / "standard_font.txt"
"""Excel's standard font and size, as the run found them (Excel for Mac's settings)."""
FIRST_ROW = 4
MARKERS = ("B2", "E3")
ZOOMS = (100,)
TEXT = "Hxg"
FIELDS = (
    "group", "entry", "cell", "row_height", "font_name", "font_size", "font_color", "ink_top",
    "ink_bottom", "red", "dark", "glyphs",
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
    if group in FACES:
        found = [(f"C{FIRST_ROW + k}", FACES_FIRST + k) for k in range(FACES_PROBES)]
        return found + [(f"E{FIRST_ROW + k}", 1 + k) for k in range(len(REFERENCES))]
    first, count = {
        "list": (LIST_FIRST, LIST_PROBES),
        "rich": (RICH_FIRST, RICH_PROBES),
    }.get(group, (THEME_FIRST, THEME_PROBES))
    return [(f"C{FIRST_ROW + k}", first + k) for k in range(count)]


def fill(ws: Any, group: str) -> list[tuple[str, int]]:
    """Lay one group out: its filler fonts first, then its probes; each probe's cell and entry."""
    ws.column_dimensions["C"].width = 14
    found = probes(group)
    if group in FACES:
        if FACES[group] is not None:
            name, size = FACES[group]
            normal = (
                Font(name=name, sz=size, family=2, scheme="minor")
                if group == "faces_aptos"
                else Font(name=name, sz=size, family=1)
            )
            ws.parent._fonts = IndexedList([normal])
            for style in ws.parent._named_styles:
                if style.name == "Normal":
                    style.font = normal
        ws.column_dimensions["E"].width = 14
        for k, (face, size) in enumerate(REFERENCES):  # entries 1 to 5, before the filler
            ws[f"E{FIRST_ROW + k}"] = TEXT
            ws[f"E{FIRST_ROW + k}"].font = Font(name=face, size=size, color="FF000000")
        sheet = ws.parent.create_sheet("filler")
        for k in range(len(REFERENCES) + 1, FACES_FIRST):
            cell = sheet.cell(k, 1, "f")
            cell.font = Font(name="Calibri", size=11, color=f"FF{k // 256:02X}{k % 256:02X}00")
        for k in range(FACES_PROBES):
            cell = f"C{FIRST_ROW + k}"
            ws[cell] = TEXT
            ws[cell].font = Font(name="Arial", size=24, color=_red(k))
    elif group in NORMALS:
        name, size = NORMALS[group]
        normal = Font(name=name, sz=size, family=1)
        ws.parent._fonts = IndexedList([normal])
        for style in ws.parent._named_styles:
            if style.name == "Normal":
                style.font = normal
        _filler(ws, THEME_FIRST - 1)
        for k, (cell, _) in enumerate(found):
            ws[cell] = TEXT
            ws[cell].font = Font(name="Arial", size=24, color=_red(k))
    elif group == "list":
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


def themed(text: str, fonts: dict[str, str] = THEME_FONTS) -> str:
    """openpyxl's theme with ``fonts`` (``THEME_FONTS``) for its body and heading fonts."""
    text = text.replace(
        '<a:minorFont>\n        <a:latin typeface="Calibri"/>',
        f'<a:minorFont>\n        <a:latin typeface="{fonts["minor"]}"/>',
    )
    return text.replace(
        '<a:majorFont>\n        <a:latin typeface="Cambria"/>',
        f'<a:majorFont>\n        <a:latin typeface="{fonts["major"]}"/>',
    )


def postprocess(group: str) -> Any:
    """For ``theme_<size>`` and ``faces_aptos``, the theme rewritten before Excel opens the file;
    else ``None``."""
    if group not in NORMALS and group != "faces_aptos":
        return None
    fonts = APTOS_THEME if group == "faces_aptos" else THEME_FONTS

    def rewrite(path: Path) -> None:
        from capture.screen.run import patch_part

        patch_part(path, {"xl/theme/theme1.xml": lambda text: themed(text, fonts)})

    return rewrite


def build(directory: Path) -> dict[str, Path]:
    """Each group's workbook in ``directory``."""
    directory.mkdir(parents=True, exist_ok=True)
    paths = {}
    for group in GROUPS:
        wb = openpyxl.Workbook()
        fill(wb.active, group)
        paths[group] = directory / f"font_cut_{group}.xlsx"
        wb.save(paths[group])
        rewrite = postprocess(group)
        if rewrite is not None:
            rewrite(paths[group])
    return paths


def standard_font() -> str:
    """Excel's standard font and size, tab separated (needs Excel for Mac)."""
    import subprocess

    script = Path(__file__).with_name("standard_font.applescript")
    found = subprocess.run(
        ["osascript", str(script)], capture_output=True, text=True, timeout=60, check=True
    )
    return found.stdout.strip()


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
            postprocess=postprocess(group),
        )
    STANDARD.parent.mkdir(parents=True, exist_ok=True)
    STANDARD.write_text(standard_font() + "\n", encoding="utf-8")
    write()


def _inked(rgb: tuple[int, int, int]) -> bool:
    r, g, b = rgb
    return (r > 150 and g < 110 and b < 110) or max(r, g, b) < 120


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
    found = {"red": str(red), "dark": str(dark), "ink_top": "", "ink_bottom": "", "glyphs": ""}
    columns = [
        x for x in range(x0, x1) if any(_inked(cap.image.pixel(x, y)) for y in range(y0, y1))
    ]
    starts = [x for k, x in enumerate(columns) if k == 0 or columns[k - 1] != x - 1]
    found["glyphs"] = " ".join(pt(cap, x - x0) for x in starts)
    if rows:
        found["ink_top"] = pt(cap, rows[0] - top)
        found["ink_bottom"] = pt(cap, rows[-1] + 1 - top)
    return found


def group_records(cap: Any, group: str) -> list[dict[str, str]]:
    """Every probe of ``group`` measured in one capture of its window, in ``FIELDS``' keys."""
    records = []
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
    return records


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
            records += group_records(cap, group)
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after", "zooms"):
                out.write(f"# {key}: {value}\n")
        if STANDARD.exists():
            name, _, size = STANDARD.read_text(encoding="utf-8").strip().partition("\t")
            out.write(f"# standard_font: {name} {size}\n")
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
