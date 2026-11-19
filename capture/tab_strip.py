"""Capture how Excel for Mac draws the sheet tab strip below the grid, from the screen (route C).

    export EXCEL_EDITION="Microsoft 365"
    uv run --group capture python -m capture.tab_strip run precedence     # alone, first
    uv run --group capture python -m capture.tab_strip run                # the other groups
    uv run --group capture python -m capture.tab_strip run light precedence
    uv run --group capture python -m capture.tab_strip run light          # the others, light
    uv run --group capture python -m capture.tab_strip read               # the captures again
    uv run --group capture python -m capture.tab_strip build PATH         # the workbooks

The tab strip is window chrome: it lies below the grid, at the window's bottom left, above the
status bar, and a full-screen window (2560 points wide) shows every tab of these workbooks. Each
group is its own workbook and window, at zoom 100; the active sheet is the first one (the one
the capture's markers, B2 and D6, are on). Captures so far are in macOS's dark appearance;
``run light`` takes the groups again in the light appearance (switch macOS's appearance once,
for all suites), into a folder of its own: a run refuses an appearance other than the one it
names. Groups:

- ``order``: seven sheets in file order One (visible, active, uncolored), Two (hidden), Three,
  Four (very hidden), Five, Six (hidden), Seven, their ``sheetId`` in another order (7, 3, 5, 1,
  6, 2, 4). Every sheet but One has its own tab color (:data:`ORDER`), hidden ones too, so each
  tab on screen names its sheet by its color, and a hidden sheet's tab would show.
- ``colors``: Active (active, rgb ``0070C0``), Plain (uncolored), Rgb (``C00000``), Theme
  (theme 5, accent 2), Tint (theme 5 at tint 0.4), Indexed (17), Auto (``auto="1"``) and Bare
  (uncolored): every kind of tab color inactive, the active tab colored, auto beside uncolored.
- ``precedence``: First (active, uncolored) and four tabs whose ``tabColor`` carries two kinds
  at once (written over openpyxl's XML): indexed 17 with rgb, theme 5 with rgb, auto with rgb,
  theme 5 with indexed 17. Excel may repair or refuse the file, so the group runs on its own; a
  refusal is kept (``refused.txt``) and written as the group's one row, its ``note``
  ``refused:`` and Excel's message.

Measured in each window: the strip's top (the first device pixel, down column A from the first
marker, that is neither the sheet's white nor a gridline, for at least 4 points), its background
(the commonest color of the row 3 points below the top), and in that row the runs of one color
(within 10 per channel), each at least 2 points wide: a run unlike the background is a ``fill``
item (a filled tab); a run like it is split where its ink (below) leaves a gap of 8 points, each
part an ``ink`` item (an uncolored tab, the scroll arrows, the new-sheet button). For each item:
its left edge and width in screen points; the color down a column 6 points inside its left edge,
one sample a point for 24 points (``profile``); a fill's height (how far that column stays
unlike the background, by steps of at most 10) and its colors 3 points below its top and 2
above its bottom; the ink inside it (pixels more than 48 from the row's color in that column,
in a channel): its width, height and core color (the pixels at least 90 % as far from the
background as the farthest). ``read`` writes ``tests/captures/tab_strip.tsv``, one row per
appearance, group and item, left to right.

Settles (ARCHITECTURE §4.15, "The tab strip" and "The sheet list"): which sheets have a tab
(visible only, file order, not ``sheetId`` order); an inactive colored tab's fill and text; the
active tab's gradient and text, colored and uncolored; the tab's size and spacing; ``auto``
drawn as an uncolored tab; which of indexed, theme, auto and rgb Excel takes when several are
given.
"""

from __future__ import annotations

import hashlib
import itertools
import re
import subprocess
import sys
from collections import Counter
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles.colors import Color


@dataclass(frozen=True)
class Tab:
    """One sheet of a group: its name, state and tab color, and the crafted ``tabColor`` XML."""

    name: str
    state: str = "visible"
    color: Color | None = None
    xml: str | None = None
    """A ``tabColor`` element written over openpyxl's, for colors of two kinds at once."""


def _rgb(value: str) -> Color:
    return Color(rgb="FF" + value)


RED, GREEN, BLUE, AMBER, PURPLE, CYAN = "C00000", "00B050", "0070C0", "FFC000", "7030A0", "00FFFF"
ORDER = (
    Tab("One"),
    Tab("Two", "hidden", _rgb(RED)),
    Tab("Three", color=_rgb(GREEN)),
    Tab("Four", "veryHidden", _rgb(BLUE)),
    Tab("Five", color=_rgb(AMBER)),
    Tab("Six", "hidden", _rgb(PURPLE)),
    Tab("Seven", color=_rgb(CYAN)),
)
SHEET_IDS = (7, 3, 5, 1, 6, 2, 4)
"""The ``order`` group's ``sheetId`` attributes, in file order: unlike the file order."""
COLORS = (
    Tab("Active", color=_rgb(BLUE)),
    Tab("Plain"),
    Tab("Rgb", color=_rgb(RED)),
    Tab("Theme", color=Color(theme=5)),
    Tab("Tint", color=Color(theme=5, tint=0.4)),
    Tab("Indexed", color=Color(indexed=17)),
    Tab("Auto", color=Color(auto=True)),
    Tab("Bare"),
)
PRECEDENCE = (
    Tab("First"),
    Tab("IdxRgb", xml=f'<tabColor indexed="17" rgb="FF{RED}"/>'),
    Tab("ThmRgb", xml=f'<tabColor theme="5" rgb="FF{BLUE}"/>'),
    Tab("AutoRgb", xml=f'<tabColor auto="1" rgb="FF{BLUE}"/>'),
    Tab("ThmIdx", xml='<tabColor theme="5" indexed="17"/>'),
)
GROUPS: dict[str, tuple[Tab, ...]] = {"order": ORDER, "colors": COLORS, "precedence": PRECEDENCE}
REFUSALS = ("precedence",)
"""Groups whose file Excel may repair or refuse: run on their own, a refusal is the answer."""
APPEARANCES = ("dark", "light")

ACTIVE_INK = "278452"
EXPECTED: dict[str, dict[str, list[tuple[str, str, str, str]]]] = {
    "dark": {
        "order": [
            ("One", "FFFFFF", "FFFFFF", "278452"),
            ("Three", "00B050", "00B050", "FFFFFF"),
            ("Five", "FFC000", "FFC000", "FFFFFF"),
            ("Seven", "00FFFF", "00FFFF", "FFFFFF"),
        ],
        "colors": [
            ("Active", "E6F1F9", "99C6E6", "278452"),
            ("Plain", "3A3A3A", "3A3A3A", "BBBBBB"),
            ("Rgb", "C00000", "C00000", "FFFFFF"),
            ("Theme", "C0504D", "C0504D", "FFFFFF"),
            ("Tint", "DA9694", "DA9694", "FFFFFF"),
            ("Indexed", "008000", "008000", "FFFFFF"),
            ("Auto", "3A3A3A", "3A3A3A", "BBBBBB"),
            ("Bare", "3A3A3A", "3A3A3A", "BBBBBB"),
        ],
        "precedence": [
            ("First", "FFFFFF", "FFFFFF", "278452"),
            ("IdxRgb", "008000", "008000", "FFFFFF"),
            ("ThmRgb", "C0504D", "C0504D", "FFFFFF"),
            ("AutoRgb", "3A3A3A", "3A3A3A", "BBBBBB"),
            ("ThmIdx", "008000", "008000", "FFFFFF"),
        ],
    },
    "light": {
        "order": [
            ("One", "FFFFFF", "FFFFFF", "278452"),
            ("Three", "00B050", "00B050", "FFFFFF"),
            ("Five", "FFC000", "FFC000", "FFFFFF"),
            ("Seven", "00FFFF", "00FFFF", "FFFFFF"),
        ],
        "colors": [
            ("Active", "E6F1F9", "99C6E6", "278452"),
            ("Plain", "E6E6E6", "E6E6E6", "444444"),
            ("Rgb", "C00000", "C00000", "FFFFFF"),
            ("Theme", "C0504D", "C0504D", "FFFFFF"),
            ("Tint", "DA9694", "DA9694", "FFFFFF"),
            ("Indexed", "008000", "008000", "FFFFFF"),
            ("Auto", "E6E6E6", "E6E6E6", "444444"),
            ("Bare", "E6E6E6", "E6E6E6", "444444"),
        ],
        "precedence": [
            ("First", "FFFFFF", "FFFFFF", "278452"),
            ("IdxRgb", "008000", "008000", "FFFFFF"),
            ("ThmRgb", "C0504D", "C0504D", "FFFFFF"),
            ("AutoRgb", "E6E6E6", "E6E6E6", "444444"),
            ("ThmIdx", "008000", "008000", "FFFFFF"),
        ],
    },
}
"""What sheetpeek's workbook page draws for each group in each appearance, from
:func:`sheetpeek_view`: each tab left to right, as (name, fill at its top, fill at its bottom,
text color). An uncolored tab takes the page's scheme; the active tab is the first."""
TAB_SIZE = (72.0, 20.0)
TAB_GAP = 7.0
"""sheetpeek's tab, in points: 72 wide (at least), 20 high, 7 apart (§4.15)."""

MARKERS = ("B2", "D6")
ZOOM = 100
TOP_RUN = 4
"""Points of non-sheet pixels, down column A, that make the strip's top."""
ROW = 3
"""Points below the strip's top of the row that is split into items."""
SAME = 10
"""Per channel: pixels this close are one run of a row."""
MIN_RUN = 2
"""Points: a narrower run (a separator, an anti-aliased edge) is no item."""
GAP = 8
"""Points of ink-free columns that split a background run into items."""
PROFILE_AT = 6
"""Points inside an item's left edge of the column its colors are sampled down."""
BAND = 24
"""Points below the strip's top that an item is measured over."""
INK = 48
"""A channel this far from the row's color in the profile column makes a pixel ink."""
CORE = 0.9
ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "tab_strip"
OUTPUT = ROOT / "tests" / "captures" / "tab_strip.tsv"
COLUMNS = (
    "appearance", "group", "item", "kind", "left", "width", "height", "top_color",
    "bottom_color", "ink", "ink_w", "ink_h", "ink_color", "profile", "note",
)  # fmt: skip


def fill(ws: Any, group: str) -> list[Tab]:
    """Lay one group out: ``ws`` is the first, active sheet; the others follow it. Its tabs."""
    wb = ws.parent
    tabs = GROUPS[group]
    for k, tab in enumerate(tabs):
        sheet = ws if k == 0 else wb.create_sheet(tab.name)
        sheet.title = tab.name
        sheet.sheet_state = tab.state
        if tab.color is not None:
            sheet.sheet_properties.tabColor = tab.color
        elif tab.xml is not None:
            sheet.sheet_properties.tabColor = _rgb("000000")  # replaced by postprocess
        sheet["C4"] = tab.name
    wb.active = 0
    return list(tabs)


def _sheet_ids(text: str) -> str:
    found = list(re.finditer(r'sheetId="\d+"', text))
    assert len(found) == len(SHEET_IDS), len(found)
    for hit, number in reversed(list(zip(found, SHEET_IDS, strict=True))):
        text = text[: hit.start()] + f'sheetId="{number}"' + text[hit.end() :]
    return text


def _tab_color(xml: str) -> Callable[[str], str]:
    def edit(text: str) -> str:
        text, count = re.subn(r"<tabColor [^>]*/>", xml, text)
        assert count == 1, xml
        return text

    return edit


def postprocess(group: str) -> Callable[[Path], None] | None:
    """What a group's saved workbook needs that openpyxl does not write, or ``None``."""
    from capture.screen.run import patch_part

    if group == "order":
        return lambda path: patch_part(path, {"xl/workbook.xml": _sheet_ids})
    if group == "precedence":
        edits: dict[str, Callable[[str], str] | str] = {
            f"xl/worksheets/sheet{k}.xml": _tab_color(tab.xml)
            for k, tab in enumerate(PRECEDENCE, 1)
            if tab.xml is not None
        }
        return lambda path: patch_part(path, edits)
    return None


def build(path: Path) -> list[Tab]:
    """Write every group to a workbook of its own in the folder ``path``; every tab."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        found.extend(fill(wb.active, group))
        target = path / f"tab_strip_{group}.xlsx"
        wb.save(target)
        post = postprocess(group)
        if post is not None:
            post(target)
    return found


def sheetpeek_view(path: Path, appearance: str) -> list[tuple[str, str, str, str]]:
    """The tabs sheetpeek's workbook page draws: (name, top fill, bottom fill, text color)."""
    import sheetpeek

    page = sheetpeek.render_workbook(path, color_scheme=appearance, on_notice="ignore")
    assert page is not None
    nav = page[page.index("<nav>") : page.index("</nav>")]
    plain = re.search(r"\.sp-book > nav a \{ background: #(\w{6}); color: #(\w{6}); \}", page)
    assert plain is not None
    paper, ink = plain.groups()
    colored = dict(
        re.findall(r"nav a\.sp-t-(\w+) \{ background: #(\w{6}); color: #ffffff; \}", page)
    )
    active = re.search(
        r"@supports not .*?a\.sp-first \{ background: ([^;]+); color: #(\w{6})", page
    )
    assert active is not None
    stops = re.findall(r"#(\w{6})", active.group(1))
    found = []
    for classes, label in re.findall(
        r'<a href="#sp-\d+"(?: class="([^"]*)")?[^>]*>([^<]+)</a>', nav
    ):
        names = classes.split()
        if "sp-first" in names:
            look = (stops[0], stops[-1], active.group(2))
        else:
            tag = next((c[5:] for c in names if c.startswith("sp-t-")), None)
            look = (colored[tag],) * 2 + ("ffffff",) if tag else (paper, paper, ink)
        found.append((label, *(c.upper() for c in look)))
    return found


def _appearance() -> str:
    found = subprocess.run(
        ["defaults", "read", "-g", "AppleInterfaceStyle"], capture_output=True, text=True
    ).stdout.strip()
    return "dark" if found == "Dark" else "light"


def run(groups: tuple[str, ...] | None = None, appearance: str = "dark") -> None:
    """Capture each group's window in ``appearance``, then write the TSV (needs Excel).

    Without groups, every group but :data:`REFUSALS`, which run on their own first.
    """
    from capture.screen.run import capture
    from capture.screen.session import ExcelError

    if _appearance() != appearance:
        raise SystemExit(f"macOS is not in the {appearance} appearance: switch it first")
    if groups is None:
        groups = tuple(g for g in GROUPS if g not in REFUSALS)
    for group in groups:
        out = BUILD_DIR / appearance / group
        (out / "refused.txt").unlink(missing_ok=True)
        try:
            capture(
                lambda ws, group=group: fill(ws, group),
                cells=[],
                out=out,
                markers=MARKERS,
                zooms=(ZOOM,),
                name=f"tab_strip_{group}",
                postprocess=postprocess(group),
            )
        except ExcelError as error:
            if group not in REFUSALS:
                raise
            record_refusal(appearance, group, str(error))
            print(f"{group}: Excel refused the file", file=sys.stderr)
    write()


def record_refusal(appearance: str, group: str, message: str) -> None:
    """Keep Excel's refusal of a group's file in ``refused.txt``, without the script's path."""
    out = BUILD_DIR / appearance / group
    out.mkdir(parents=True, exist_ok=True)
    reason = message.rpartition("execution error: ")[2].strip() or message
    (out / "refused.txt").write_text(reason + "\n", encoding="utf-8")


# -- measuring ------------------------------------------------------------------------------------

Pixel = tuple[int, int, int]


def _far(a: Pixel, b: Pixel) -> int:
    return max(abs(x - y) for x, y in zip(a, b, strict=True))


def _hex(pixels: list[Pixel]) -> str:
    return "".join(f"{round(sum(p[i] for p in pixels) / len(pixels)):02X}" for i in range(3))


def strip_top(cap: Any) -> int | None:
    """The strip's first device row, down column A from the first marker; ``None`` if none."""
    per = cap.scale * 100 / cap.zoom
    x = round(cap.grid.x(0) + 2 * per)
    sheet = [(255, 255, 255), *([cap.gridline] if cap.gridline else [])]
    need = round(TOP_RUN * per)
    run = 0
    for y in range(round(cap.edge(MARKERS[0], "bottom")) + 2, cap.image.height):
        if all(_far(cap.image.pixel(x, y), s) > 6 for s in sheet):
            run += 1
            if run == need:
                return y - need + 1
        else:
            run = 0
    return None


def _runs(row: list[Pixel], least: int) -> list[tuple[int, int, Pixel]]:
    """Runs of one color in a row, at least ``least`` pixels: (start, end exclusive, color)."""
    found = []
    start = 0
    for x in range(1, len(row) + 1):
        if x == len(row) or _far(row[x], row[start]) > SAME:
            if x - start >= least:
                found.append((start, x, row[start + (x - start) // 2]))
            start = x
    return found


def _item(cap: Any, top: int, x0: int, x1: int, background: Pixel, kind: str) -> dict[str, str]:
    per = cap.scale * 100 / cap.zoom
    image = cap.image
    px = min(x0 + round(PROFILE_AT * per), (x0 + x1) // 2)
    bottom = min(top + round(BAND * per), image.height)

    def spt(pixels: float) -> str:
        return f"{round(pixels / per, 2):g}"

    found = {"kind": kind, "left": spt(x0), "width": spt(x1 - x0)}
    found["profile"] = " ".join(
        _hex([image.pixel(px, min(round(top + (k + 0.5) * per), image.height - 1))])
        for k in range(BAND)
    )
    y0, y1 = top, bottom
    if kind == "fill":

        def inside(y: int, toward: int) -> bool:
            here = image.pixel(px, y)
            return _far(here, background) > SAME and _far(here, image.pixel(px, toward)) <= SAME

        y = top + round(ROW * per)
        while y < bottom and inside(y, y - 1):
            y += 1
        z = top + round(ROW * per)
        while z > top and inside(z - 1, z):
            z -= 1
        y0, y1 = z, y
        found["height"] = spt(y1 - y0)
        found["top_color"] = _hex([image.pixel(px, y0 + round(ROW * per))])
        found["bottom_color"] = _hex([image.pixel(px, max(y0, y1 - 1 - round(2 * per)))])
    reference = {y: image.pixel(px, y) if kind == "fill" else background for y in range(y0, y1)}
    pixels = {
        (x, y): image.pixel(x, y) for y in range(y0 + 2, y1 - 2) for x in range(x0 + 2, x1 - 2)
    }
    ink = [xy for xy, p in pixels.items() if _far(p, reference[xy[1]]) > INK]
    if ink:
        gx0, gy0 = min(x for x, _ in ink), min(y for _, y in ink)
        gx1, gy1 = max(x for x, _ in ink), max(y for _, y in ink)
        marked = set(ink)
        bitmap = "/".join(
            "".join("#" if (x, y) in marked else "." for x in range(gx0, gx1 + 1))
            for y in range(gy0, gy1 + 1)
        )
        farthest = max(_far(pixels[xy], reference[xy[1]]) for xy in ink)
        core = [pixels[xy] for xy in ink if _far(pixels[xy], reference[xy[1]]) >= CORE * farthest]
        found |= {
            "ink": hashlib.sha1(bitmap.encode()).hexdigest()[:8],
            "ink_w": spt(gx1 - gx0 + 1),
            "ink_h": spt(gy1 - gy0 + 1),
            "ink_color": _hex(core),
        }
    return found


def _split(cap: Any, top: int, x0: int, x1: int, background: Pixel) -> list[tuple[int, int]]:
    """A background run split at gaps of :data:`GAP` points between its ink: [(x0, x1)]."""
    per = cap.scale * 100 / cap.zoom
    bottom = min(top + round(BAND * per), cap.image.height)
    inked = [
        x
        for x in range(x0 + 2, x1 - 2)
        if any(_far(cap.image.pixel(x, y), background) > INK for y in range(top + 2, bottom - 2))
    ]
    if not inked:
        return [(x0, x1)]
    cuts = [(a + b) // 2 for a, b in itertools.pairwise(inked) if b - a > GAP * per]
    edges = [x0, *cuts, x1]
    return list(itertools.pairwise(edges))


def measure(cap: Any, group: str, appearance: str) -> list[dict[str, str]]:
    """The strip's items, left to right, in one group's window."""
    per = cap.scale * 100 / cap.zoom
    top = strip_top(cap)
    if top is None:
        return [{"appearance": appearance, "group": group, "note": "no strip found"}]
    y = top + round(ROW * per)
    row = [cap.image.pixel(x, y) for x in range(cap.image.width)]
    background = Counter(row).most_common(1)[0][0]
    found = []
    for x0, x1, color in _runs(row, round(MIN_RUN * per)):
        if _far(color, background) > SAME:
            found.append(_item(cap, top, x0, x1, background, "fill"))
        else:
            for a, b in _split(cap, top, x0, x1, background):
                found.append(_item(cap, top, a, b, background, "ink"))
    for k, item in enumerate(found, 1):
        item |= {"appearance": appearance, "group": group, "item": str(k)}
    return found


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``tab_strip.tsv``."""
    from capture.screen.run import header, load

    records: list[dict[str, str]] = []
    first: dict[str, str] = {}
    for appearance in APPEARANCES:
        for group in GROUPS:
            out = BUILD_DIR / appearance / group
            refused = out / "refused.txt"
            if refused.exists():
                reason = refused.read_text(encoding="utf-8").strip()
                records.append(
                    {"appearance": appearance, "group": group, "note": f"refused: {reason}"}
                )
                continue
            if not (out / "formats.tsv").exists():
                print(f"{appearance} {group}: not captured yet, left out", file=sys.stderr)
                continue
            found = header(out)
            first = first or found
            for cap in load(out):
                records.extend(measure(cap, group, found.get("appearance", appearance)))
    with open(OUTPUT, "w", encoding="utf-8") as out_file:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after", "appearance"):
                out_file.write(f"# {key}: {value}\n")
        out_file.write("# suite: tab_strip\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out_file.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} tabs in {argv[1]}")
        return 0
    if argv[:1] == ["run"]:
        rest = argv[1:]
        appearance = "dark"
        if rest[:1] == ["light"]:
            appearance, rest = "light", rest[1:]
        if all(group in GROUPS for group in rest):
            run(tuple(rest) or None, appearance)
            return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
