"""Capture how Excel for Mac draws the grid color and the row and column headings (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.headings_grid run grid_x
    EXCEL_EDITION=... uv run --group capture python -m capture.headings_grid run grid_99999
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.headings_grid run
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.headings_grid run light
    uv run --group capture python -m capture.headings_grid read      # the saved captures again
    uv run --group capture python -m capture.headings_grid build PATH  # every group's workbook

Each group is its own workbook and window (a sheet has one grid color and one Normal font, and
a window one scroll position). ``run`` takes every group but ``grid_x`` and ``grid_99999`` in
the dark appearance; ``run grid_x`` and ``run grid_99999`` run those groups alone first, since
Excel may refuse their files (16.115 refused colorId 99999 with a parameter error): the refusal
is kept (``refused.txt``) and written as the group's one row, its ``kind``
``refused`` and Excel's message in ``item``. ``run light`` captures ``grid_default`` and
``grid_10_teal`` again at zoom 100 once macOS is switched to the light appearance (a person's
step: System Settings, Appearance), into ``capture/build/headings_grid/light/``. Each run
refuses to start in the other appearance. The groups:

- ``grid_*`` (zoom 100 and 200 for ``grid_default``, ``grid_10_teal`` and ``grid_12``, 100 for
  the rest): ``<sheetView defaultGridColor="0" colorId>`` with colorId 10 under a palette whose
  entry 10 is teal (``008080``), 12 under the default palette (blue), 64, 65, 99999, and ``x``
  (written into the XML; schema-invalid); ``grid_default`` has neither attribute. Each holds
  C4 filled white, E4 yellow and the yellow block G4:H5, and leaves C8 and its neighbors empty.
- ``rows_*`` (Calibri 11, zoom 100 and 200): the window scrolled so its top row is 100, 1000,
  10000, 1000000 and 9990 (``rows_4_5``: the first rows have 4 digits, the last 5).
  ``grid_*`` give the 2-digit case (rows 1 to about 70).
- ``font_*`` (zoom 100): the Normal style's font Calibri 8, 14, 20, 28 and 36 (the theme's minor
  font) and Aptos Narrow 11, with no ``<sheetFormatPr>``, so Excel sizes rows and columns from it.

The markers are B2 and J10 in ``grid_*``, B2 and D6 in ``font_*``, and the second row and sixth
row below the top in ``rows_*`` (column B and F). Only the markers are read; the grid cells'
places follow from the first marker's size, since every row and column has the default size.
After each view the window is scrolled to the group's top row (``capture(top_row=...)``), and
Excel's visible range is kept (the read ``window``, ``visible.<zoom>``).

Measured, from each capture (``kind`` ``headings``): the row band's width and the column band's
height in sheet points (image pixels over the scale; pixels beside them), each the run of the
heading paper's color from the cells' first boundary outward, read a quarter into the first
marker's row and column; the paper's color; the separators between rows in the row band and
between columns in the column band (the first line of one color across the whole band); the
headings' text color (the pixel farthest in luma from the paper among the first three rows'
or columns' headings); and the colors across the boundary between headings and cells, as runs
(``RRGGBBxN``). In ``grid_*`` (``kind`` ``edge``): the colors across C8's right and bottom
boundaries (open grid) and across every side of C4 (white), E4 (yellow) and G4's inner sides
(yellow block), from 3 screen points before the boundary to 3 after it, averaged over 5 pixels
along it, as runs. ``read`` writes ``tests/captures/headings_grid.tsv``.

Settles in ARCHITECTURE §4.14: the grid color (the palette entry, 64 and past the palette
automatic, a malformed colorId, the 1-device-pixel line, a fill covering it, the headings
keeping their colors) and the headings (the bands at Calibri 11 by the digits of the last row
in the window, at zoom 100 and 200; the bands per Normal font; the colors in the dark and light
appearances).
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.utils.indexed_list import IndexedList

TEAL = "008080"
"""Palette entry 10 in ``grid_10_teal`` (red in the default palette)."""
CALIBRI_11 = ("Calibri", 11.0)


@dataclass(frozen=True)
class Group:
    """One workbook and window: its grid color, Normal font, top row and zooms."""

    name: str
    color_id: int | str | None = None
    """``colorId`` with ``defaultGridColor="0"``; text is written into the XML as it stands."""
    teal: bool = False
    font: tuple[str, float] = CALIBRI_11
    top_row: int = 1
    zooms: tuple[int, ...] = (100, 200)
    grid: bool = False
    """Whether the fills and the measured grid edges are laid out."""

    @property
    def markers(self) -> tuple[str, str]:
        if self.top_row != 1:
            return f"B{self.top_row + 1}", f"F{self.top_row + 6}"
        return ("B2", "J10") if self.grid else ("B2", "D6")


GROUPS = (
    Group("grid_default", grid=True),
    Group("grid_10_teal", 10, teal=True, grid=True),
    Group("grid_12", 12, grid=True),
    Group("grid_64", 64, zooms=(100,), grid=True),
    Group("grid_65", 65, zooms=(100,), grid=True),
    Group("grid_99999", 99999, zooms=(100,), grid=True),
    Group("rows_3", top_row=100),
    Group("rows_4", top_row=1000),
    Group("rows_5", top_row=10000),
    Group("rows_7", top_row=1000000),
    Group("rows_4_5", top_row=9990),
    *(
        Group(f"font_calibri_{size}", font=("Calibri", float(size)), zooms=(100,))
        for size in (8, 14, 20, 28, 36)
    ),
    Group("font_aptos_narrow_11", font=("Aptos Narrow", 11.0), zooms=(100,)),
    Group("grid_x", "x", zooms=(100,), grid=True),
)
GROUP = {g.name: g for g in GROUPS}
REFUSALS = ("grid_99999", "grid_x")
"""Groups whose file Excel may refuse: run on their own, a refusal is the answer."""
LIGHT = ("grid_default", "grid_10_teal")
"""Groups captured again in the light appearance, at zoom 100."""
FILLS = {"C4": "FFFFFF", "E4": "FFFF00", "G4": "FFFF00", "H4": "FFFF00", "G5": "FFFF00",
         "H5": "FFFF00"}  # fmt: skip
EDGES = (
    ("open", "C8", "right"), ("open", "C8", "bottom"),
    *(("white", "C4", side) for side in ("left", "right", "top", "bottom")),
    *(("yellow", "E4", side) for side in ("left", "right", "top", "bottom")),
    ("yellow block", "G4", "right"), ("yellow block", "G4", "bottom"),
)  # fmt: skip
"""The grid boundaries measured in ``grid_*``: (item, cell, side)."""

EXPECTED_GRID = {
    "grid_default": None, "grid_10_teal": TEAL, "grid_12": "0000FF", "grid_64": None,
    "grid_65": None, "grid_99999": None, "grid_x": None,
}  # fmt: skip
"""sheetpeek's ``Sheet.grid_color`` for each grid group (``None``: the automatic gray, drawn
2 device pixels wide; a color is drawn 1 device pixel wide)."""
EXPECTED_DIGITS = {"rows_3": 3, "rows_4": 4, "rows_5": 5, "rows_7": 7, "rows_4_5": 5}
"""The digits of the last row the window shows, as built; 2 for the other groups while the
window shows fewer than 100 rows (the capture records Excel's visible range)."""
EXPECTED_BANDS = {
    ("Calibri", 11.0, 2): (22.0, 21.0), ("Calibri", 11.0, 3): (25.0, 21.0),
    ("Calibri", 11.0, 4): (28.0, 21.0), ("Calibri", 11.0, 5): (31.0, 21.0),
    ("Calibri", 11.0, 7): (37.0, 21.0), ("Calibri", 8.0, 2): (17.33, 19.67),
    ("Calibri", 14.0, 2): (26.67, 22.33), ("Calibri", 20.0, 2): (36.0, 25.0),
    ("Calibri", 28.0, 2): (48.44, 28.56), ("Calibri", 36.0, 2): (60.89, 32.11),
    ("Aptos Narrow", 11.0, 2): (22.0, 21.0),
}  # fmt: skip
"""sheetpeek's row band width and column band height in sheet points, by Normal font and
digits (``backends/html/_headings.bands``, rounded to 0.01)."""
EXPECTED_HEADING_COLORS = {
    "dark": {"paper": "161616", "text": "FFFFFF", "separator": "565656"},
    "light": {"paper": "F7F7F7", "text": "5E5E5E", "separator": "E0DFDD"},
}
"""sheetpeek's heading colors (its HTML stylesheet's ``--sheetpeek-heading-*``)."""

STRIP = 3
"""Screen points measured on each side of a boundary."""
ALONG = 5
"""Device pixels averaged along a boundary."""
PAPER_TOLERANCE = 3
"""How far (per channel) a pixel may be from the heading paper and still count as paper."""
ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "headings_grid"
OUTPUT = ROOT / "tests" / "captures" / "headings_grid.tsv"
COLUMNS = (
    "group", "appearance", "zoom", "kind", "item", "cell", "side", "font", "color_id",
    "top_row", "visible", "row_band", "col_band", "row_band_px", "col_band_px", "band_left_px",
    "paper", "separator_row", "separator_col", "text_row", "text_col", "edge_row", "edge_col",
    "strip",
)  # fmt: skip

Pixel = tuple[int, int, int]


def fill(ws: Any, name: str) -> list[str]:
    """Lay one group out on a worksheet; the cells whose boundaries are measured."""
    group = GROUP[name]
    wb = ws.parent
    if group.font != CALIBRI_11:
        family, size = group.font
        normal = Font(
            name=family, sz=size, family=2, scheme="minor" if family == "Calibri" else None
        )
        wb._fonts = IndexedList([normal])
        for style in wb._named_styles:
            if style.name == "Normal":
                style.font = normal
    if group.teal:
        colors = list(wb._colors)
        colors[10] = "00" + TEAL
        wb._colors = colors
    if isinstance(group.color_id, int):
        ws.sheet_view.defaultGridColor = False
        ws.sheet_view.colorId = group.color_id
    if not group.grid:
        return []
    for cell, color in FILLS.items():
        ws[cell].fill = PatternFill("solid", fgColor="FF" + color)
    return list(dict.fromkeys(cell for _, cell, _ in EDGES))


def postprocess(name: str) -> Any:
    """What a group's saved workbook needs that openpyxl does not write, or ``None``."""
    from capture.screen.run import patch_part

    group = GROUP[name]
    sheet = "xl/worksheets/sheet1.xml"
    if isinstance(group.color_id, str):
        attributes = f'defaultGridColor="0" colorId="{group.color_id}" '

        def crafted(text: str) -> str:
            if text.count("<sheetView ") != 1:
                raise SystemExit("the sheet has not one <sheetView>")
            return text.replace("<sheetView ", "<sheetView " + attributes)

        return lambda path: patch_part(path, {sheet: crafted})
    if group.font != CALIBRI_11:

        def no_format(text: str) -> str:
            found, count = re.subn(r"<sheetFormatPr\b[^>]*/>", "", text)
            if count != 1:
                raise SystemExit("the sheet has not one <sheetFormatPr>")
            return found

        return lambda path: patch_part(path, {sheet: no_format})
    return None


def build(path: Path) -> list[str]:
    """Write every group to a workbook of its own in the folder ``path``; the files' names."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        fill(wb.active, group.name)
        target = path / f"headings_grid_{group.name}.xlsx"
        wb.save(target)
        post = postprocess(group.name)
        if post is not None:
            post(target)
        found.append(target.name)
    return found


def _appearance() -> str:
    from capture.screen.session import environment

    return environment("")["appearance"]


def run(groups: tuple[str, ...] | None = None, *, light: bool = False) -> None:
    """Capture each group's window, then write the TSV (needs Excel).

    Without groups, every group but :data:`REFUSALS`; ``light`` captures :data:`LIGHT` at zoom
    100 into ``light/``. The appearance must be the one the run is for.
    """
    from capture.screen.run import capture
    from capture.screen.session import ExcelError

    wanted = "light" if light else "dark"
    if _appearance() != wanted:
        raise SystemExit(f"switch macOS to the {wanted} appearance first")
    if groups is None:
        groups = LIGHT if light else tuple(g.name for g in GROUPS if g.name not in REFUSALS)
    for name in groups:
        group = GROUP[name]
        try:
            capture(
                lambda ws, name=name: fill(ws, name),
                cells=[],
                out=_out(name, light),
                markers=group.markers,
                zooms=(100,) if light else group.zooms,
                name=f"headings_grid_{name}",
                postprocess=postprocess(name),
                top_row=group.top_row,
            )
        except ExcelError as error:
            if name not in REFUSALS or "got an error" not in str(error):
                raise  # not a file Excel may refuse, or the session's own refusal to start
            record_refusal(name, str(error))
            print(f"{name}: Excel refused the file", file=sys.stderr)
    write()


def _out(name: str, light: bool) -> Path:
    return BUILD_DIR / "light" / name if light else BUILD_DIR / name


def record_refusal(name: str, message: str) -> None:
    """Keep Excel's refusal of a group's file in ``refused.txt``, without the script's path."""
    out = _out(name, False)
    out.mkdir(parents=True, exist_ok=True)
    reason = message.rpartition("execution error: ")[2].strip() or message
    (out / "refused.txt").write_text(reason + "\n", encoding="utf-8")


# -- measuring ---------------------------------------------------------------------------------


def _hex(rgb: Pixel) -> str:
    return "".join(f"{c:02X}" for c in rgb)


def _close(a: Pixel, b: Pixel, tolerance: int = PAPER_TOLERANCE) -> bool:
    return max(abs(x - y) for x, y in zip(a, b, strict=True)) <= tolerance


def _luma(rgb: Pixel) -> float:
    return 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]


def _mean(pixels: list[Pixel]) -> str:
    return "".join(f"{round(sum(p[i] for p in pixels) / len(pixels)):02X}" for i in range(3))


def runs(colors: list[str]) -> str:
    """Colors as runs of one color, ``RRGGBBxN`` each: ``FFFFFFx6 D4D4D4x2 FFFFFFx4``."""
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


def _across(cap: Any, x: int, y: int, vertical: bool) -> str:
    """The colors across a boundary at (x, y): a vertical boundary is crossed along x."""
    per = cap.scale * 100 / cap.zoom  # image pixels a screen point
    reach = round(STRIP * per)
    found = []
    for d in range(-reach, reach):
        if vertical:
            pixels = [cap.image.pixel(x + d, y + k) for k in range(-(ALONG // 2), ALONG // 2 + 1)]
        else:
            pixels = [cap.image.pixel(x + k, y + d) for k in range(-(ALONG // 2), ALONG // 2 + 1)]
        found.append(_mean(pixels))
    return runs(found)


def place(cap: Any, group: Group, cell: str) -> None:
    """Give ``cell`` its geometry from the first marker's: every row and column is default."""
    if cell in cap.geometry:
        return
    first = cap.geometry[group.markers[0]]
    letters, row = coordinate_from_string(group.markers[0])
    column = column_index_from_string(letters)
    their_letters, their_row = coordinate_from_string(cell)
    their_column = column_index_from_string(their_letters)
    cap.geometry[cell] = {
        "left": first["left"] + (their_column - column) * first["width"],
        "top": first["top"] + (their_row - row) * first["height"],
        "width": first["width"],
        "height": first["height"],
    }


def _edge(cap: Any, cell: str, side: str) -> str:
    left, right = cap.edge(cell, "left"), cap.edge(cell, "right")
    top, bottom = cap.edge(cell, "top"), cap.edge(cell, "bottom")
    if side in ("left", "right"):
        return _across(
            cap, round(left if side == "left" else right), round((top + bottom) / 2), True
        )
    return _across(cap, round((left + right) / 2), round(top if side == "top" else bottom), False)


def _headings(cap: Any, group: Group) -> dict[str, str]:
    """The bands' sizes and the headings' colors in one capture."""
    first = cap.geometry[group.markers[0]]
    _, marker_row = coordinate_from_string(group.markers[0])
    sheet_top = first["top"] - (marker_row - group.top_row) * first["height"]
    x_cells = round(cap.grid.x(0))  # column A's left boundary: the window starts at column A
    y_cells = round(cap.grid.y(sheet_top))
    y_quarter = round(cap.grid.y(first["top"] + first["height"] / 4))
    x_quarter = round(cap.grid.x(first["left"] + first["width"] / 4))
    image = cap.image
    paper = image.pixel(x_cells - 1, y_quarter)
    left = x_cells - 1
    while left > 0 and _close(image.pixel(left - 1, y_quarter), paper):
        left -= 1
    top = y_cells - 1
    while top > 0 and _close(image.pixel(x_quarter, top - 1), paper):
        top -= 1
    row_px, col_px = x_cells - left, y_cells - top
    row_height = first["height"] * cap.scale
    column_width = first["width"] * cap.scale

    def lines(spans: range, across: range, vertical: bool) -> str:
        for k in spans:
            found = {image.pixel(k, j) if vertical else image.pixel(j, k) for j in across}
            if len(found) == 1:
                (color,) = found
                if not _close(color, paper):
                    return _hex(color)
        return ""

    def farthest(xs: range, ys: range) -> str:
        best, distance = paper, -1.0
        for y in ys:
            for x in xs:
                p = image.pixel(x, y)
                if abs(_luma(p) - _luma(paper)) > distance:
                    best, distance = p, abs(_luma(p) - _luma(paper))
        return _hex(best)

    # From the middle of the first row and column: the line between the headings and the cells
    # runs on across the other band at their boundaries.
    rows_span = range(
        y_cells + round(row_height / 2), min(image.height, y_cells + round(3 * row_height))
    )
    columns_span = range(
        x_cells + round(column_width / 2), min(image.width, x_cells + round(3 * column_width))
    )
    return {
        "row_band": f"{round(row_px / cap.scale, 3):g}",
        "col_band": f"{round(col_px / cap.scale, 3):g}",
        "row_band_px": str(row_px),
        "col_band_px": str(col_px),
        "band_left_px": str(left),
        "paper": _hex(paper),
        "separator_row": lines(rows_span, range(left + 1, x_cells - 1), False),
        "separator_col": lines(columns_span, range(top + 1, y_cells - 1), True),
        "text_row": farthest(range(left, x_cells), rows_span),
        "text_col": farthest(columns_span, range(top, y_cells)),
        "edge_row": _across(cap, x_cells, y_quarter, True),
        "edge_col": _across(cap, x_quarter, y_cells, False),
    }


def measure(cap: Any, name: str, appearance: str = "dark") -> list[dict[str, str]]:
    """Every record of one group's capture at one zoom."""
    group = GROUP[name]
    base = {
        "group": name,
        "appearance": appearance,
        "zoom": str(cap.zoom),
        "font": f"{group.font[0]} {group.font[1]:g}",
        "color_id": "" if group.color_id is None else str(group.color_id),
    }
    head = base | {
        "kind": "headings",
        "top_row": str(group.top_row),
        "visible": cap.value("window", f"visible.{cap.zoom}") or "",
    }
    records = [head | _headings(cap, group)]
    if group.grid:
        for item, cell, side in EDGES:
            place(cap, group, cell)
            records.append(
                base
                | {"kind": "edge", "item": item, "cell": cell, "side": side}
                | {"strip": _edge(cap, cell, side)}
            )
    return records


def last_row(visible: str) -> int | None:
    """The last row of a visible range's address (``$A$9990:$AQ$10059``), or ``None``."""
    found = re.findall(r"\$?[A-Z]+\$?(\d+)", visible)
    return int(found[-1]) if found else None


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``headings_grid.tsv``."""
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
        out_file.write("# suite: headings_grid\n" + "\t".join(COLUMNS) + "\n")
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
    if argv[:1] == ["run"] and all(name in GROUP for name in argv[1:]):
        run(tuple(argv[1:]))
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
