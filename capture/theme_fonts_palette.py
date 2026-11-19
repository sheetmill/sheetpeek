"""Capture theme fonts and the workbook palette in Excel for Mac, from the screen and its reads.

    export EXCEL_EDITION="Microsoft 365"
    uv run --group capture python -m capture.theme_fonts_palette run no_theme
    uv run --group capture python -m capture.theme_fonts_palette run empty_typeface
    uv run --group capture python -m capture.theme_fonts_palette run palette_bad
    uv run --group capture python -m capture.theme_fonts_palette run palette_long
    uv run --group capture python -m capture.theme_fonts_palette run          # the other groups
    uv run --group capture python -m capture.theme_fonts_palette read         # captures again
    uv run --group capture python -m capture.theme_fonts_palette build PATH   # the workbooks

Each group is its own workbook, opened on its own (ARCHITECTURE §4.14 and §3.3). The groups
whose file is crafted so that Excel may repair or refuse it (:data:`REFUSALS`) run one at a
time, by name, before the rest; a refusal is kept (``refused.txt``) and written as the group's
one row, its ``case`` ``refused:`` and Excel's message. ``run`` with no group takes the others
(as does ``python -m capture.screen run theme_fonts_palette`` once the suite is listed there).

In a window (zoom 100, the dark appearance; the second marker below and right of the cases):

- ``fonts``: a theme whose minor font is Courier New (Japanese: Hiragino Mincho ProN) and whose
  major font is Georgia (Japanese: Hiragino Maru Gothic ProN). Text in Calibri-minor,
  Times-New-Roman-minor, Calibri-major and Calibri without a scheme; text in the Normal style
  (Calibri, minor, as openpyxl writes it); controls in Courier New, Georgia and Times New Roman
  without a scheme. Japanese text in Calibri-minor, Calibri-major and plain Calibri, with
  controls in the two Hiragino faces. D4 lies in a default-width column, E4 in a column 20
  characters wide: their widths follow the Normal font.
- ``palette``: ``indexedColors`` with the default's 64 entries but 3, 10, 12 and 17 recolored
  (:data:`RECOLORED`). Fills in indexed 2, 3, 10, 11 and 17; bold 18 pt text in font colors
  indexed 2, 3, 10 and 17; bold 18 pt numbers in ``[Color10]``, ``[Color3]``, ``[Red]`` and
  ``[Blue]``.
- ``no_theme``: the theme part removed (its relationship and content type too): text in
  Arial-minor and in the Normal style, controls in Arial and Calibri. Excel may repair it.
- ``empty_typeface``: the same cells under a theme whose minor ``<a:latin typeface=""/>``.

Read only (no window; Excel's ``display format`` and geometry settle them):

- ``palette_short``: ``indexedColors`` of 20 entries, 10 teal: fills in 10, 17, 30 and 2.
- ``palette_bad``: 64 entries, 10 teal and 17 ``rgb="zz"``: fills in 10 and 17.
- ``palette_long``: 70 entries, 10 teal, 63 and 64 to 69 recolored: fills in 10, 63 and 66.

Measured for every case: Excel's display format (font name, font color, fill) and the cell's
width and height in points, exact, from AppleScript. In a window, besides: a text cell's ink
(pixels more than 64 from the cell's background in a channel, inside the cell less 2 device
pixels on each side) as a bitmap with an id (equal ids, identical glyphs), its width and height
in screen points, and its color (the mean of the ink pixels at least 90 % as far from the
background as the farthest: the strokes' cores, not their anti-aliased edges); a fill cell's
mean color over 5 by 5 device pixels at its center. ``read`` writes
``tests/captures/theme_fonts_palette.tsv``, one row per group and cell.

Settles (§4.14): theme fonts (scheme fonts draw the theme's Latin typeface; the Normal font and
the default column follow the theme; without a theme or with an empty typeface the stored name
is drawn; Japanese text in a minor or major font draws the theme's ``Jpan`` font); the workbook
palette (entry by entry, 0 to 7 not mirroring 8 to 15; fills and fonts follow it; a short list,
a non-color entry and a list past 64); number-format colors (``[ColorN]`` draws palette N + 7,
``[Red]`` and ``[Blue]`` keep their default colors).
"""

from __future__ import annotations

import hashlib
import re
import shutil
import sys
import zipfile
from collections import Counter
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color

GROUPS = (
    "fonts",
    "palette",
    "palette_short",
    "no_theme",
    "empty_typeface",
    "palette_bad",
    "palette_long",
)
WINDOWED = ("fonts", "palette", "no_theme", "empty_typeface")
"""Groups captured in a window; the others are read through Excel's own properties only."""
REFUSALS = ("no_theme", "empty_typeface", "palette_bad", "palette_long")
"""Groups whose file Excel may repair or refuse: run on their own, a refusal is the answer."""
MARKERS = {"fonts": "G19", "palette": "H10", "no_theme": "F9", "empty_typeface": "F9"}
"""Each window's second marker (the first is B2)."""
ZOOM = 100
TEXT = "Hamburgefonstiv 0123456789"
JAPANESE = "日本語の文字表示"
NUMBER = 8888
TEAL, BROWN, OLIVE, ORANGE = "008080", "804000", "996600", "FF8000"
RECOLORED = {3: OLIVE, 10: TEAL, 12: ORANGE, 17: BROWN}
"""The ``palette`` group's entries: [Red] is slot 10 ([Color3]), [Blue] slot 12, [Color10] 17."""
MINCHO, MARU = "Hiragino Mincho ProN", "Hiragino Maru Gothic ProN"
INK = 64
"""A channel this far from the background makes a pixel ink."""
CORE = 0.9
"""Ink pixels this close to the farthest from the background give the ink's color."""
INSET = 2
"""Device pixels left out on each side of a cell: its gridlines."""
ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "theme_fonts_palette"
OUTPUT = ROOT / "tests" / "captures" / "theme_fonts_palette.tsv"
COLUMNS = (
    "group", "cell", "case", "kind", "font_name", "font_color", "fill", "width", "height",
    "ink", "ink_w", "ink_h", "ink_color", "pixel",
)  # fmt: skip
READS = {
    "font_name": "display.font_name",
    "font_color": "display.font_color",
    "fill": "display.fill",
    "width": "geometry.width",
    "height": "geometry.height",
}


@dataclass(frozen=True)
class Case:
    """One cell to measure: its group, address, what it is there for and how it is measured.

    ``kind`` is ``text`` (ink compared with a control's), ``font`` (font color), ``format``
    (number-format color), ``fill`` (fill color) or ``width`` (the column's width).
    """

    group: str
    cell: str
    case: str
    kind: str


def _font(name: str, scheme: str | None = None, **more: Any) -> Font:
    return Font(name=name, sz=more.pop("sz", 11), scheme=scheme, **more)


TEXT_FONTS = {
    "fonts": {
        "B4": ("Calibri, minor", _font("Calibri", "minor"), TEXT),
        "B5": ("Times New Roman, minor", _font("Times New Roman", "minor"), TEXT),
        "B6": ("Calibri, major", _font("Calibri", "major"), TEXT),
        "B7": ("Calibri, no scheme", _font("Calibri"), TEXT),
        "B8": ("Normal style", None, TEXT),
        "B9": ("control Courier New", _font("Courier New"), TEXT),
        "B10": ("control Georgia", _font("Georgia"), TEXT),
        "B11": ("control Times New Roman", _font("Times New Roman"), TEXT),
        "B13": ("Japanese, Calibri, minor", _font("Calibri", "minor"), JAPANESE),
        "B14": ("Japanese, Calibri, major", _font("Calibri", "major"), JAPANESE),
        "B15": ("Japanese, Calibri, no scheme", _font("Calibri"), JAPANESE),
        "B16": (f"control Japanese {MINCHO}", _font(MINCHO), JAPANESE),
        "B17": (f"control Japanese {MARU}", _font(MARU), JAPANESE),
    },
    "no_theme": {
        "B4": ("Arial, minor", _font("Arial", "minor"), TEXT),
        "B5": ("Normal style", None, TEXT),
        "B6": ("control Arial", _font("Arial"), TEXT),
        "B7": ("control Calibri", _font("Calibri"), TEXT),
    },
}
TEXT_FONTS["empty_typeface"] = TEXT_FONTS["no_theme"]
CONTROLS = {
    "fonts": {"Courier New": "B9", "Georgia": "B10", "Times New Roman": "B11", "Calibri": "B7"},
    "no_theme": {"Arial": "B6", "Calibri": "B7"},
    "empty_typeface": {"Arial": "B6", "Calibri": "B7"},
}
"""Each group's control cell for a typeface: a case drawn in that face draws the control's ink.
Calibri without a scheme is its own control in ``fonts``: its ink must differ from Courier New's."""
JAPANESE_CONTROLS = {"B13": "B16", "B14": "B17"}
"""The theme's Jpan fonts: minor Mincho, major Maru Gothic (ARCHITECTURE §4.14, other scripts)."""
WIDTHS = {
    "fonts": {"D4": None, "E4": 20},
    "no_theme": {"D4": None},
    "empty_typeface": {"D4": None},
}
"""Width cases: a default-width column (``None``) and one of a given width in characters."""
PALETTE_FILLS = {
    "palette": {"B4": 2, "B5": 3, "B6": 10, "B7": 11, "B8": 17},
    "palette_short": {"B4": 10, "B5": 17, "B6": 30, "B7": 2},
    "palette_bad": {"B4": 10, "B5": 17},
    "palette_long": {"B4": 10, "B5": 63, "B6": 66},
}
PALETTE_FONTS = {"D4": 2, "D5": 3, "D6": 10, "D7": 17}
FORMATS = {"F4": "[Color10]0", "F5": "[Color3]0", "F6": "[Red]0", "F7": "[Blue]0"}

EXPECTED: dict[str, dict[str, str | float | None]] = {
    "fonts": {
        "B4": "Courier New", "B5": "Courier New", "B6": "Georgia", "B7": "Calibri",
        "B8": "Courier New", "B9": "Courier New", "B10": "Georgia", "B11": "Times New Roman",
        "B13": "Courier New", "B14": "Georgia", "B15": "Calibri", "B16": MINCHO, "B17": MARU,
        "D4": 53.0, "E4": 120.0,
    },
    "palette": {
        "B4": "FF0000", "B5": OLIVE, "B6": TEAL, "B7": "00FF00", "B8": BROWN,
        "D4": "FF0000", "D5": "00FF00", "D6": TEAL, "D7": BROWN,
        "F4": BROWN, "F5": TEAL, "F6": "FF0000", "F7": "0000FF",
    },
    "palette_short": {"B4": TEAL, "B5": "008000", "B6": "0066CC", "B7": "FF0000"},
    "palette_bad": {"B4": TEAL, "B5": "000000"},
    "palette_long": {"B4": TEAL, "B5": "123456", "B6": "000000"},
    "no_theme": {
        "B4": "Aptos Narrow", "B5": "Aptos Narrow", "B6": "Arial", "B7": "Calibri", "D4": 53.0,
    },
    "empty_typeface": {
        "B4": "Arial", "B5": "Calibri", "B6": "Arial", "B7": "Calibri", "D4": 53.0,
    },
}  # fmt: skip
"""What sheetpeek draws today, from :func:`sheetpeek_view` on each group's built workbook: the
typeface of a text case, the color (``RRGGBB``) of a fill, font or format case, ``None`` for no
fill, and a width case's column width in points."""


def cases(group: str) -> list[Case]:
    """A group's cases, in the order they are measured."""
    found = [
        Case(group, cell, name, "text") for cell, (name, _, _) in TEXT_FONTS.get(group, {}).items()
    ]
    for cell, width in WIDTHS.get(group, {}).items():
        label = "default column" if width is None else f"column of {width} characters"
        found.append(Case(group, cell, label, "width"))
    for cell, index in PALETTE_FILLS.get(group, {}).items():
        found.append(Case(group, cell, f"fill indexed {index}", "fill"))
    if group == "palette":
        for cell, index in PALETTE_FONTS.items():
            found.append(Case(group, cell, f"font indexed {index}", "font"))
        for cell, code in FORMATS.items():
            found.append(Case(group, cell, f"format {code}", "format"))
    return found


def _palette(size: int, entries: dict[int, str]) -> list[str]:
    from sheetpeek._colors import DEFAULT_PALETTE

    colors = [f"00{DEFAULT_PALETTE[i % 64]}" for i in range(size)]
    for index, rgb in entries.items():
        colors[index] = f"00{rgb}"
    return colors


PALETTES = {
    "palette": _palette(64, RECOLORED),
    "palette_short": _palette(20, {10: TEAL}),
    "palette_bad": _palette(64, {10: TEAL}),
    "palette_long": _palette(70, {10: TEAL, 63: "123456", **dict.fromkeys(range(64, 70), BROWN)}),
}


def fill(ws: Any, group: str) -> list[Case]:
    """Lay one group out on a worksheet; its cases."""
    wb = ws.parent
    ws.column_dimensions["B"].width = 40
    for cell, (_, font, text) in TEXT_FONTS.get(group, {}).items():
        ws[cell] = text
        if font is not None:
            ws[cell].font = font
    for cell, width in WIDTHS.get(group, {}).items():
        ws[cell] = 1
        if width is not None:
            ws.column_dimensions[cell[0]].width = width
    if group in PALETTES:
        wb._colors = PALETTES[group]
        ws.column_dimensions["B"].width = 20
        for cell, index in PALETTE_FILLS[group].items():
            ws[cell].fill = PatternFill("solid", fgColor=Color(indexed=index))
    if group == "palette":
        big = {"sz": 18, "bold": True}
        for column in "DF":
            ws.column_dimensions[column].width = 20
        for row in range(4, 9):
            ws.row_dimensions[row].height = 30
        for cell, index in PALETTE_FONTS.items():
            ws[cell] = "MMMM"
            ws[cell].font = Font(color=Color(indexed=index), **big)
        for cell, code in FORMATS.items():
            ws[cell] = NUMBER
            ws[cell].number_format = code
            ws[cell].font = Font(**big)
    return cases(group)


def theme_fonts(
    minor: str | None = None,
    major: str | None = None,
    minor_jpan: str | None = None,
    major_jpan: str | None = None,
) -> Callable[[Path], None]:
    """A postprocess that sets the theme's Latin and Japanese typefaces (``None`` keeps one)."""

    def edit(text: str) -> str:
        for block, latin, jpan in (
            ("minorFont", minor, minor_jpan),
            ("majorFont", major, major_jpan),
        ):
            start = text.index(f"<a:{block}>")
            end = text.index(f"</a:{block}>", start)
            part = text[start:end]
            if latin is not None:
                part, count = re.subn(
                    r'<a:latin typeface="[^"]*"', f'<a:latin typeface="{latin}"', part
                )
                assert count == 1, block
            if jpan is not None:
                part, count = re.subn(
                    r'<a:font script="Jpan" typeface="[^"]*"',
                    f'<a:font script="Jpan" typeface="{jpan}"',
                    part,
                )
                assert count == 1, block
            text = text[:start] + part + text[end:]
        return text

    def post(path: Path) -> None:
        from capture.screen.run import patch_part

        patch_part(path, {"xl/theme/theme1.xml": edit})

    return post


def drop_theme(path: Path) -> None:
    """Remove the theme part, its relationship and its content type from a saved workbook."""
    tmp = path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for name in src.namelist():
            if name == "xl/theme/theme1.xml":
                continue
            text = src.read(name)
            if name == "xl/_rels/workbook.xml.rels":
                found = re.subn(rb"<Relationship [^>]*theme1\.xml[^>]*/>", b"", text)
                assert found[1] == 1, name
                text = found[0]
            if name == "[Content_Types].xml":
                found = re.subn(rb'<Override PartName="/xl/theme/theme1\.xml"[^>]*/>', b"", text)
                assert found[1] == 1, name
                text = found[0]
            dst.writestr(name, text)
    shutil.move(tmp, path)


def bad_entry(path: Path) -> None:
    """Make the palette's entry 17 ``rgb="zz"``."""
    from capture.screen.run import patch_part

    def edit(text: str) -> str:
        entries = list(re.finditer(r'<rgbColor rgb="[^"]*"', text))
        assert len(entries) == 64, len(entries)
        hit = entries[17]
        return text[: hit.start()] + '<rgbColor rgb="zz"' + text[hit.end() :]

    patch_part(path, {"xl/styles.xml": edit})


def postprocess(group: str) -> Callable[[Path], None] | None:
    """What a group's saved workbook needs that openpyxl does not write, or ``None``."""
    return {
        "fonts": theme_fonts("Courier New", "Georgia", MINCHO, MARU),
        "no_theme": drop_theme,
        "empty_typeface": theme_fonts(minor=""),
        "palette_bad": bad_entry,
    }.get(group)


def build(path: Path) -> list[Case]:
    """Write every group to a workbook of its own in the folder ``path``; every case."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        found.extend(fill(wb.active, group))
        target = path / f"theme_fonts_palette_{group}.xlsx"
        wb.save(target)
        post = postprocess(group)
        if post is not None:
            post(target)
    return found


def sheetpeek_view(path: Path, group: str) -> dict[str, str | float | None]:
    """What sheetpeek draws for each of a group's cases in the workbook at ``path``."""
    from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

    import sheetpeek

    sheet = sheetpeek.read(path)
    laid = sheet.layout("excel")
    found: dict[str, str | float | None] = {}
    for case in cases(group):
        letters, row = coordinate_from_string(case.cell)
        column = column_index_from_string(letters)
        cell = sheet.cell(row, column)
        if case.kind == "text":
            found[case.cell] = cell.font.name
        elif case.kind == "fill":
            found[case.cell] = cell.style.fill
        elif case.kind == "font":
            found[case.cell] = cell.font.color
        elif case.kind == "format":
            found[case.cell] = laid.cell(row, column).format_color
        else:
            found[case.cell] = next(c.width_pt for c in laid.columns if c.index == column)
    return found


def run(groups: tuple[str, ...] | None = None) -> None:
    """Capture or read each group, then write the TSV (needs Excel).

    Without groups, every group but :data:`REFUSALS`, which run on their own first.
    """
    from capture.screen.run import capture, read_only
    from capture.screen.session import ExcelError

    if groups is None:
        groups = tuple(g for g in GROUPS if g not in REFUSALS)
    for group in groups:
        cells = [c.cell for c in cases(group)]
        out = BUILD_DIR / group
        (out / "refused.txt").unlink(missing_ok=True)
        try:
            if group in WINDOWED:
                capture(
                    lambda ws, group=group: fill(ws, group),
                    cells=cells,
                    out=out,
                    markers=("B2", MARKERS[group]),
                    zooms=(ZOOM,),
                    name=f"theme_fonts_palette_{group}",
                    postprocess=postprocess(group),
                )
            else:
                read_only(
                    lambda ws, group=group: fill(ws, group),
                    cells=cells,
                    out=out,
                    name=f"theme_fonts_palette_{group}",
                    postprocess=postprocess(group),
                )
        except ExcelError as error:
            if group not in REFUSALS:
                raise
            record_refusal(group, str(error))
            print(f"{group}: Excel refused the file", file=sys.stderr)
    write()


def record_refusal(group: str, message: str) -> None:
    """Keep Excel's refusal of a group's file in ``refused.txt``, without the script's path."""
    out = BUILD_DIR / group
    out.mkdir(parents=True, exist_ok=True)
    reason = message.rpartition("execution error: ")[2].strip() or message
    (out / "refused.txt").write_text(reason + "\n", encoding="utf-8")


def _hex(pixels: list[tuple[int, int, int]]) -> str:
    return "".join(f"{round(sum(p[i] for p in pixels) / len(pixels)):02X}" for i in range(3))


def _rgb(read: str | None) -> str:
    """Excel's ``r,g,b`` as ``RRGGBB``; anything else as it stands."""
    if read and re.fullmatch(r"\d+,\d+,\d+", read):
        return "".join(f"{int(part):02X}" for part in read.split(","))
    return read or ""


def _ink(cap: Any, cell: str) -> dict[str, str]:
    """A text cell's ink: bitmap id, size in screen points, core color."""
    per = cap.scale * 100 / cap.zoom  # image pixels a screen point
    x0, x1 = round(cap.edge(cell, "left")) + INSET, round(cap.edge(cell, "right")) - INSET
    y0, y1 = round(cap.edge(cell, "top")) + INSET, round(cap.edge(cell, "bottom")) - INSET
    pixels = {(x, y): cap.image.pixel(x, y) for y in range(y0, y1) for x in range(x0, x1)}
    background = Counter(pixels.values()).most_common(1)[0][0]

    def far(p: tuple[int, int, int]) -> int:
        return max(abs(a - b) for a, b in zip(p, background, strict=True))

    dark = [xy for xy, p in pixels.items() if far(p) > INK]
    if not dark:
        return {}
    gx0, gy0 = min(x for x, _ in dark), min(y for _, y in dark)
    gx1, gy1 = max(x for x, _ in dark), max(y for _, y in dark)
    marked = set(dark)
    bitmap = "/".join(
        "".join("#" if (x, y) in marked else "." for x in range(gx0, gx1 + 1))
        for y in range(gy0, gy1 + 1)
    )
    farthest = max(far(pixels[xy]) for xy in dark)
    core = [pixels[xy] for xy in dark if far(pixels[xy]) >= CORE * farthest]

    def spt(count: int) -> str:
        return f"{round(count / per, 2):g}"

    return {
        "ink": hashlib.sha1(bitmap.encode()).hexdigest()[:8],
        "ink_w": spt(gx1 - gx0 + 1),
        "ink_h": spt(gy1 - gy0 + 1),
        "ink_color": _hex(core),
    }


def _center(cap: Any, cell: str) -> str:
    x, y = round(cap.edge(cell, "cx")), round(cap.edge(cell, "cy"))
    return _hex([cap.image.pixel(x + dx, y + dy) for dy in range(-2, 3) for dx in range(-2, 3)])


def record(group: str, case: Case, value: Callable[[str, str], str | None]) -> dict[str, str]:
    """A case's row from Excel's reads (``value(cell, property)``), without pixels."""
    found = {"group": group, "cell": case.cell, "case": case.case, "kind": case.kind}
    for key, read in READS.items():
        got = value(case.cell, read) or ""
        found[key] = _rgb(got) if key in ("font_color", "fill") else got
    return found


def measure(cap: Any, group: str) -> list[dict[str, str]]:
    """Every case's record in one group's window."""
    found = []
    for case in cases(group):
        row = record(group, case, cap.value)
        if case.kind in ("text", "font", "format"):
            row |= _ink(cap, case.cell)
        elif case.kind == "fill":
            row["pixel"] = _center(cap, case.cell)
        found.append(row)
    return found


def _reads(out: Path) -> Callable[[str, str], str | None]:
    """Cell reads of a saved ``formats.tsv``, as ``Capture.value`` gives them."""
    found: dict[tuple[str, str], str] = {}
    for line in (out / "formats.tsv").read_text(encoding="utf-8").splitlines():
        if line.startswith("# ") or line.startswith("cell\t"):
            continue
        cell, key, value = [*line.split("\t"), "", ""][:3]
        found[(cell, key)] = value
    return lambda cell, key: found.get((cell, key))


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``theme_fonts_palette.tsv``."""
    from capture.screen.run import header, load

    records: list[dict[str, str]] = []
    first: dict[str, str] = {}
    for group in GROUPS:
        out = BUILD_DIR / group
        refused = out / "refused.txt"
        if refused.exists():
            reason = refused.read_text(encoding="utf-8").strip()
            records.append({"group": group, "case": f"refused: {reason}"})
            continue
        if not (out / "formats.tsv").exists():
            print(f"{group}: not captured yet, left out", file=sys.stderr)
            continue
        first = first or header(out)
        if group in WINDOWED:
            for cap in load(out):
                records.extend(measure(cap, group))
        else:
            value = _reads(out)
            records.extend(record(group, case, value) for case in cases(group))
    with open(OUTPUT, "w", encoding="utf-8") as out_file:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out_file.write(f"# {key}: {value}\n")
        out_file.write("# suite: theme_fonts_palette\n" + "\t".join(COLUMNS) + "\n")
        for row in records:
            out_file.write("\t".join(row.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cases in {argv[1]}")
        return 0
    if argv[:1] == ["run"] and all(group in GROUPS for group in argv[1:]):
        run(tuple(argv[1:]) or None)
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
