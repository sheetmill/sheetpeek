"""Capture how Excel for Mac draws a frozen pane, page breaks and a saved zoom (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.freeze_view run zoom
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.freeze_view run
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.freeze_view run light
    uv run --group capture python -m capture.freeze_view read       # the saved captures again
    uv run --group capture python -m capture.freeze_view build PATH # every workbook

The sheet (``fill``) is frozen at C3 (two rows and two columns), every row and column of the
default size, so the freeze lines fall on the boundary between rows 2 and 3 and between columns
B and C. Along the horizontal line: E2 and E3 empty, G2 and G3 yellow, I3 teal below it only,
K2 with a medium black bottom, M3 with a thick blue top. Along the vertical line: B5 and C5
empty, B7 and C7 yellow, B9 with a medium black right side, B11 underlined text long enough to
spill across the line into D11, C13 right-aligned text spilling left across it into A13. The
sheet has a print area ``$B$2:$Q$24`` and manual breaks below row 16 and after column H. The
markers are D4 and P23, both inside the print area and away from the middle of its pages; S26
holds a word, so that the used range covers the print area and the breaks.

Groups, each its own workbook:

- ``normal`` (zoom 100 and 200, dark appearance): the sheet as built, in the Normal view.
- ``page_break`` (zoom 100): the same sheet saved in page break preview (``view=
  "pageBreakPreview"`` written into its ``<sheetView>``, with ``zoomScaleSheetLayoutView="100"``),
  so Excel opens it in that view; the window's zoom is then set as in every capture.
- ``zoom_*`` (no screen): a sheet whose ``zoomScale`` is ``5``, ``401``, ``85.5``, ``abc``, empty,
  or ``85`` (the control), written into the XML. Each is opened and the zoom Excel reports for
  its window read, with nothing set first. Excel may refuse a crafted file (-50 with alerts
  off): ``run zoom`` takes these groups on their own, a refusal is kept (``refused.txt``) and
  written as the group's one row, its ``kind`` ``refused`` and Excel's message in ``item``.

``run`` takes ``normal`` and ``page_break`` in the dark appearance; ``run light`` captures
``normal`` again at zoom 100 once macOS is switched to the light appearance (a person's step:
System Settings, Appearance), into ``capture/build/freeze_view/light/``. Each run refuses to start
in the other appearance. After each view the window's state is read (zoom, view, freeze panes,
split row and column: ``ExcelSession.window``) and written as a ``window`` row.

Only the markers are read; every other cell's place follows from the first marker's size.
Measured (``kind`` ``edge``) across each boundary in ``EDGES``, at its middle: the colors of the
device pixels from 3 screen points before it to 3 after it, averaged over 5 pixels along it, as
runs (``FFFFFFx5 9B9B9Bx2 FFFFFFx5``), and the line in it: the run of pixels unlike both ends of
the strip nearest the boundary, its offset from the boundary and its length in device pixels and
screen points, and its mean color. For the spilled text (``kind`` ``spill``): the ink (luma
under 150) in the text's row on either side of the vertical line, leaving out a screen point on
each side of it, as the leftmost and rightmost ink in sheet points from the line; and how many
pixels on the open line's own columns in that row are dark (luma under 100), which is text drawn
over the line. ``kind`` ``paper``: the color in the middle of A19 (outside the print area) and
F19 (inside). ``read`` writes ``tests/captures/freeze_view.tsv``.

Settles in ARCHITECTURE §4.18: what Excel's Normal view draws (neither the print area nor page
breaks; the freeze line one screen point wide, gray 155, over borders and fills, and over or
under spilled text), the freeze line's color and width at zoom 100 and 200 and in the light
appearance (``backends/html/_panes.py``, ``FREEZE_LINE``), the look of manual breaks in page
break preview (3 px of ``#0000D0``, ``BREAK_LINE``), and a ``zoomScale`` outside 10 to 400 or not
a whole number (``_read_view.py``, ``_view_model.py``).
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.worksheet.pagebreak import Break

FREEZE = "C3"
PRINT_AREA = "B2:Q24"
ROW_BREAK = 16
"""The manual row break's ``id``: below row 16."""
COLUMN_BREAK = 8
"""The manual column break's ``id``: after column H."""
MARKERS = ("D4", "P23")
END = "S26"
"""A value past the print area and the breaks, so that sheetpeek's drawn range holds them."""
YELLOW, TEAL, BLUE, BLACK = "FFFF00", "008080", "0000FF", "000000"
SPILL_RIGHT = "underlined text spills across the freeze line"
SPILL_LEFT = "right-aligned text spills left"
SPILLS = (("spill right", "B11", "C11"), ("spill left", "C13", "C13"))
"""(item, the cell holding the text, the cell whose left side is the freeze line in its row)."""

EDGES = (
    ("freeze h open", "E3", "top"),
    ("freeze h fill", "G3", "top"),
    ("freeze h fill below", "I3", "top"),
    ("freeze h medium border", "K3", "top"),
    ("freeze h thick blue", "M3", "top"),
    ("freeze v open", "C5", "left"),
    ("freeze v fill", "C7", "left"),
    ("freeze v medium border", "C9", "left"),
    ("freeze v spill right", "C11", "left"),
    ("freeze v spill left", "C13", "left"),
    ("row break", "E17", "top"),
    ("column break", "I19", "left"),
    ("print area left", "B19", "left"),
    ("print area top", "E2", "top"),
    ("print area right", "R19", "left"),
    ("print area bottom", "E25", "top"),
)
"""The boundaries measured: (item, cell, side)."""
PAPER = (("outside print area", "A19"), ("inside print area", "F19"))


@dataclass(frozen=True)
class Group:
    """One workbook: its sheet view, its zooms, and whether a window is captured."""

    name: str
    zooms: tuple[int, ...] = (100,)
    page_break: bool = False
    zoom_scale: str | None = None
    """``zoomScale`` written into the XML as it stands; ``None`` leaves it out."""

    @property
    def screen(self) -> bool:
        return self.zoom_scale is None


ZOOM_CASES = {"5": "5", "401": "401", "85.5": "85.5", "abc": "abc", "empty": "", "85": "85"}
"""Each crafted group's ``zoomScale``, by the name after ``zoom_``; ``85`` is the control."""
GROUPS = (
    Group("normal", zooms=(100, 200)),
    Group("page_break", page_break=True),
    *(Group(f"zoom_{name}", zooms=(), zoom_scale=value) for name, value in ZOOM_CASES.items()),
)
GROUP = {g.name: g for g in GROUPS}
SCREEN = ("normal", "page_break")
REFUSALS = tuple(g.name for g in GROUPS if not g.screen)
"""Groups whose file Excel may refuse: run on their own (``run zoom``), a refusal is the answer."""
LIGHT = ("normal",)
"""Groups captured again in the light appearance, at zoom 100."""

EXPECTED = {
    "freeze_color": "9B9B9B",
    "freeze_width": 1.0,
    "break_color": "0000D0",
    "break_width": 3.0,
}
"""What sheetpeek draws: the freeze line and a manual break's mark, colors and widths in screen
points (CSS px), over fills, borders and text; in the Normal view nothing at the breaks or the
print area's edges (``backends/html/_panes.py``)."""
EXPECTED_ZOOM = {"5": 100, "401": 100, "85.5": 100, "abc": 100, "empty": 100, "85": 85}
"""sheetpeek's ``Sheet.view.zoom`` for each crafted group."""
EXPECTED_SPILL = {"spill right": (-50.0, 139.0), "spill left": (-69.0, 51.0)}
"""Where sheetpeek draws each spilled text, the left and right of its glyph advances in sheet
points from the freeze line (``TextRun.x`` and ``x + width``, rounded to whole points)."""

STRIP = 3
"""Screen points measured on each side of a boundary."""
ALONG = 5
"""Device pixels averaged along a boundary."""
UNLIKE = 12
"""How far (per channel) a pixel must be from both ends of a strip to count as the line."""
DARK_TEXT = 100
"""Luma under which a pixel on the freeze line counts as text drawn over it."""
ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "freeze_view"
OUTPUT = ROOT / "tests" / "captures" / "freeze_view.tsv"
COLUMNS = (
    "group", "appearance", "zoom", "kind", "item", "cell", "side", "strip", "line_offset",
    "line_px", "line_spt", "line_color", "ink_left", "ink_right", "dark_on_line", "color",
    "state",
)  # fmt: skip

Pixel = tuple[int, int, int]


def fill(ws: Any) -> list[str]:
    """Lay the frozen sheet out on a worksheet; the cells whose boundaries are measured."""
    yellow = PatternFill("solid", fgColor="FF" + YELLOW)
    for cell in ("G2", "G3", "B7", "C7"):
        ws[cell].fill = yellow
    ws["I3"].fill = PatternFill("solid", fgColor="FF" + TEAL)
    ws["K2"].border = Border(bottom=Side("medium", color="FF" + BLACK))
    ws["M3"].border = Border(top=Side("thick", color="FF" + BLUE))
    ws["B9"].border = Border(right=Side("medium", color="FF" + BLACK))
    ws["B11"] = SPILL_RIGHT
    ws["B11"].font = Font(underline="single")
    ws["C13"] = SPILL_LEFT
    ws["C13"].alignment = Alignment(horizontal="right")
    ws[END] = "end"
    ws.freeze_panes = FREEZE
    ws.print_area = PRINT_AREA
    ws.row_breaks.append(Break(id=ROW_BREAK))
    ws.col_breaks.append(Break(id=COLUMN_BREAK))
    cells = [cell for _, cell, _ in EDGES] + [cell for _, cell in PAPER]
    cells += [cell for _, text, line in SPILLS for cell in (text, line)]
    return list(dict.fromkeys(cells))


def _crafted(name: str, attributes: str) -> Any:
    """A postprocess writing ``attributes`` into the sheet's one ``<sheetView>``."""
    from capture.screen.run import patch_part

    def edit(text: str) -> str:
        if text.count("<sheetView ") != 1:
            raise SystemExit(f"{name}: the sheet has not one <sheetView>")
        return text.replace("<sheetView ", f"<sheetView {attributes} ")

    return lambda path: patch_part(path, {"xl/worksheets/sheet1.xml": edit})


def postprocess(name: str) -> Any:
    """What a group's saved workbook needs that openpyxl does not write, or ``None``."""
    group = GROUP[name]
    if group.page_break:
        return _crafted(name, 'view="pageBreakPreview" zoomScaleSheetLayoutView="100"')
    if group.zoom_scale is not None:
        return _crafted(name, f'zoomScale="{group.zoom_scale}"')
    return None


def fill_group(ws: Any, name: str) -> list[str]:
    """Lay one group out: the frozen sheet, or a crafted zoom's one value."""
    if GROUP[name].screen:
        return fill(ws)
    ws["B2"] = "zoom"
    return []


def build(path: Path) -> list[str]:
    """Write every group to a workbook of its own in the folder ``path``; the files' names."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        fill_group(wb.active, group.name)
        target = path / f"freeze_view_{group.name}.xlsx"
        wb.save(target)
        post = postprocess(group.name)
        if post is not None:
            post(target)
        found.append(target.name)
    return found


def _appearance() -> str:
    from capture.screen.session import environment

    return environment("")["appearance"]


def _out(name: str, light: bool = False) -> Path:
    return BUILD_DIR / "light" / name if light else BUILD_DIR / name


def run(groups: tuple[str, ...] | None = None, *, light: bool = False) -> None:
    """Capture each group, then write the TSV (needs Excel).

    Without groups, :data:`SCREEN`; ``light`` captures :data:`LIGHT` at zoom 100 into
    ``light/``. The appearance must be the one a run with a window is for.
    """
    from capture.screen.run import capture
    from capture.screen.session import ExcelError

    if groups is None:
        groups = LIGHT if light else SCREEN
    wanted = "light" if light else "dark"
    if any(GROUP[name].screen for name in groups) and _appearance() != wanted:
        raise SystemExit(f"switch macOS to the {wanted} appearance first")
    for name in groups:
        group = GROUP[name]
        if not group.screen:
            try:
                read_zoom(name)
            except ExcelError as error:
                record_refusal(name, str(error))
                print(f"{name}: Excel refused the file", file=sys.stderr)
            continue
        capture(
            fill,
            cells=[],
            out=_out(name, light),
            markers=MARKERS,
            zooms=(100,) if light else group.zooms,
            name=f"freeze_view_{name}",
            postprocess=postprocess(name),
            state=True,
        )
    write()


def read_zoom(name: str) -> str:
    """Open a crafted group's workbook and keep its window's state as Excel reads it."""
    import os

    from capture.screen.run import write_tsv
    from capture.screen.session import CONTAINER, ExcelSession, environment

    edition = os.environ.get("EXCEL_EDITION")
    if not edition:
        raise SystemExit("state the Excel edition (Excel > About Excel) in EXCEL_EDITION")
    out = _out(name)
    out.mkdir(parents=True, exist_ok=True)
    (out / "refused.txt").unlink(missing_ok=True)
    folder = CONTAINER / "sheetpeek-capture"
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / f"freeze_view_{name}_{os.getpid()}.xlsx"
    found = environment(edition)
    session = ExcelSession(path, evidence=out / "excel_timeouts")
    try:
        wb = openpyxl.Workbook()
        fill_group(wb.active, name)
        wb.save(path)
        postprocess(name)(path)
        with session:
            state = session.window()
    finally:
        path.unlink(missing_ok=True)
    found["suite"] = f"freeze_view_{name}"
    found["workbooks_open_after"] = session.open_after or "unknown"
    write_tsv(out / "formats.tsv", found, [("window", "state", state)])
    return state


def record_refusal(name: str, message: str) -> None:
    """Keep Excel's refusal of a group's file in ``refused.txt``, without the script's path."""
    out = _out(name)
    out.mkdir(parents=True, exist_ok=True)
    reason = message.rpartition("execution error: ")[2].strip() or message
    (out / "refused.txt").write_text(reason + "\n", encoding="utf-8")


# -- measuring ---------------------------------------------------------------------------------


def _hex(rgb: Pixel) -> str:
    return "".join(f"{c:02X}" for c in rgb)


def _rgb(text: str) -> Pixel:
    return (int(text[0:2], 16), int(text[2:4], 16), int(text[4:6], 16))


def _luma(rgb: Pixel) -> float:
    return 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]


def runs(colors: list[str]) -> str:
    """Colors as runs of one color, ``RRGGBBxN`` each: ``FFFFFFx5 9B9B9Bx2 FFFFFFx5``."""
    found: list[list[Any]] = []
    for color in colors:
        if found and found[-1][0] == color:
            found[-1][1] += 1
        else:
            found.append([color, 1])
    return " ".join(f"{color}x{count}" for color, count in found)


def strip(text: str) -> list[str]:
    """The colors of a strip written by :func:`runs`, one per device pixel."""
    found = []
    for run in text.split():
        color, _, count = run.partition("x")
        found.extend([color] * int(count))
    return found


def line(colors: list[str]) -> tuple[int, int, str] | None:
    """The line in a strip whose boundary lies after its first half.

    The run of pixels unlike both ends of the strip (by more than :data:`UNLIKE` on a channel)
    that is nearest the boundary: its start's offset from the boundary and its length in device
    pixels, and its mean color; ``None`` when every pixel is like an end.
    """
    if not colors:
        return None
    ends = (_rgb(colors[0]), _rgb(colors[-1]))
    flags = [
        all(max(abs(a - b) for a, b in zip(_rgb(c), end, strict=True)) > UNLIKE for end in ends)
        for c in colors
    ]
    middle = len(colors) // 2
    found: list[tuple[int, int]] = []
    start = None
    for k, flag in enumerate([*flags, False]):
        if flag and start is None:
            start = k
        elif not flag and start is not None:
            found.append((start, k - start))
            start = None
    if not found:
        return None

    def distance(run: tuple[int, int]) -> int:
        first, length = run
        if first <= middle < first + length or first <= middle - 1 < first + length:
            return 0
        return min(abs(first - middle), abs(first + length - middle))

    first, length = min(found, key=distance)
    pixels = [_rgb(c) for c in colors[first : first + length]]
    mean = tuple(round(sum(p[i] for p in pixels) / len(pixels)) for i in range(3))
    return first - middle, length, _hex((mean[0], mean[1], mean[2]))


def place(cap: Any, cell: str) -> None:
    """Give ``cell`` its geometry from the first marker's: every row and column is default."""
    if cell in cap.geometry:
        return
    first = cap.geometry[MARKERS[0]]
    letters, row = coordinate_from_string(MARKERS[0])
    column = column_index_from_string(letters)
    their_letters, their_row = coordinate_from_string(cell)
    their_column = column_index_from_string(their_letters)
    cap.geometry[cell] = {
        "left": first["left"] + (their_column - column) * first["width"],
        "top": first["top"] + (their_row - row) * first["height"],
        "width": first["width"],
        "height": first["height"],
    }


def _per(cap: Any) -> float:
    """Image pixels a screen point."""
    return float(cap.scale * 100 / cap.zoom)


def _across(cap: Any, x: int, y: int, vertical: bool) -> list[str]:
    """The colors across a boundary at (x, y): a vertical boundary is crossed along x."""
    reach = round(STRIP * _per(cap))
    found = []
    for d in range(-reach, reach):
        if vertical:
            pixels = [cap.image.pixel(x + d, y + k) for k in range(-(ALONG // 2), ALONG // 2 + 1)]
        else:
            pixels = [cap.image.pixel(x + k, y + d) for k in range(-(ALONG // 2), ALONG // 2 + 1)]
        found.append(
            "".join(f"{round(sum(p[i] for p in pixels) / len(pixels)):02X}" for i in range(3))
        )
    return found


def _edge(cap: Any, cell: str, side: str) -> dict[str, str]:
    left, right = cap.edge(cell, "left"), cap.edge(cell, "right")
    top, bottom = cap.edge(cell, "top"), cap.edge(cell, "bottom")
    if side == "left":
        colors = _across(cap, round(left), round((top + bottom) / 2), True)
    else:
        colors = _across(cap, round((left + right) / 2), round(top), False)
    found = {"strip": runs(colors)}
    drawn = line(colors)
    if drawn is not None:
        offset, length, color = drawn
        found |= {
            "line_offset": str(offset),
            "line_px": str(length),
            "line_spt": f"{round(length / _per(cap), 2):g}",
            "line_color": color,
        }
    return found


def _spill(cap: Any, text_cell: str, line_cell: str, open_line: tuple[int, int] | None) -> dict:
    """The ink of a spilled text on either side of the vertical freeze line, in its row."""
    per = _per(cap)
    x_line = round(cap.edge(line_cell, "left"))
    _, row = coordinate_from_string(text_cell)
    top, bottom = round(cap.edge(f"A{row}", "top")) + 2, round(cap.edge(f"A{row}", "bottom")) - 2
    band = round(per)
    x0, x1 = round(cap.edge(f"A{row}", "left")) + 2, round(cap.edge(f"F{row}", "right"))
    xs = [
        x
        for x in range(x0, x1)
        if not x_line - band <= x < x_line + band
        and any(_luma(cap.image.pixel(x, y)) < 150 for y in range(top, bottom))
    ]
    found: dict[str, str] = {}
    if xs:
        found["ink_left"] = f"{round((min(xs) - x_line) / cap.scale, 2):g}"
        found["ink_right"] = f"{round((max(xs) + 1 - x_line) / cap.scale, 2):g}"
    start, length = open_line if open_line is not None else (-1, 2)
    columns = range(x_line + start, x_line + start + length)
    dark = sum(
        1 for x in columns for y in range(top, bottom) if _luma(cap.image.pixel(x, y)) < DARK_TEXT
    )
    found["dark_on_line"] = str(dark)
    return found


def measure(cap: Any, name: str, appearance: str = "dark") -> list[dict[str, str]]:
    """Every record of one group's capture at one zoom."""
    base = {"group": name, "appearance": appearance, "zoom": str(cap.zoom)}
    for cell in fill(openpyxl.Workbook().active) + [f"A{r}" for r in (11, 13)] + ["F11", "F13"]:
        place(cap, cell)
    records = [base | {"kind": "window", "state": cap.value("window", f"state.{cap.zoom}") or ""}]
    open_line = None
    for item, cell, side in EDGES:
        found = _edge(cap, cell, side)
        if item == "freeze v open" and "line_px" in found:
            open_line = (int(found["line_offset"]), int(found["line_px"]))
        records.append(base | {"kind": "edge", "item": item, "cell": cell, "side": side} | found)
    for item, text_cell, line_cell in SPILLS:
        records.append(
            base
            | {"kind": "spill", "item": item, "cell": text_cell}
            | _spill(cap, text_cell, line_cell, open_line)
        )
    for item, cell in PAPER:
        color = cap.image.pixel(round(cap.edge(cell, "cx")), round(cap.edge(cell, "cy")))
        records.append(base | {"kind": "paper", "item": item, "cell": cell, "color": _hex(color)})
    return records


def zoom_of(state: str) -> str:
    """The zoom field of a window's state (``zoom=85;view=...``), or ``""``."""
    for part in state.split(";"):
        key, _, value = part.partition("=")
        if key == "zoom":
            return value
    return ""


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``freeze_view.tsv``."""
    from capture.screen.run import header, load

    records: list[dict[str, str]] = []
    first: dict[str, str] = {}
    light: dict[str, str] = {}
    sets = [(g.name, False) for g in GROUPS] + [(name, True) for name in LIGHT]
    for name, is_light in sets:
        out = _out(name, is_light)
        refused = out / "refused.txt"
        if refused.exists():
            reason = refused.read_text(encoding="utf-8").strip()
            records.append({"group": name, "kind": "refused", "item": reason})
            continue
        if not (out / "formats.tsv").exists():
            print(f"{name}{' (light)' if is_light else ''}: not captured yet, left out")
            continue
        found = header(out)
        if is_light:
            light = light or found
        else:
            first = first or found
        if not GROUP[name].screen:
            state = next(
                (
                    line.split("\t")[2].rstrip("\n")
                    for line in (out / "formats.tsv").read_text(encoding="utf-8").splitlines()
                    if line.startswith("window\tstate\t")
                ),
                "",
            )
            records.append(
                {"group": name, "kind": "zoom", "item": name.removeprefix("zoom_")}
                | {"zoom": zoom_of(state), "state": state}
            )
            continue
        for cap in load(out):
            records.extend(measure(cap, name, found.get("appearance", "dark")))
    if not records:
        print(f"nothing captured yet: {OUTPUT} not written")
        return
    skip = ("markers", "suite", "workbooks_open_after", "zooms", "appearance")
    with open(OUTPUT, "w", encoding="utf-8") as out_file:
        for key, value in first.items():
            if key not in skip:
                out_file.write(f"# {key}: {value}\n")
        for key in ("excel", "build", "macos", "date"):
            if key in light:
                out_file.write(f"# light {key}: {light[key]}\n")
        out_file.write("# suite: freeze_view\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out_file.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        written = build(Path(argv[1]))
        print(f"{len(written)} workbooks in {argv[1]}")
        return 0
    if argv == ["run"]:
        run()
        return 0
    if argv == ["run", "light"]:
        run(light=True)
        return 0
    if argv == ["run", "zoom"]:
        run(REFUSALS)
        return 0
    if argv[:1] == ["run"] and argv[1:] and all(name in GROUP for name in argv[1:]):
        run(tuple(argv[1:]))
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
