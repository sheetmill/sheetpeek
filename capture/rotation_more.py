"""Capture more of how Excel for Mac draws angled and stacked text, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.rotation_more run
    uv run --group capture python -m capture.rotation_more read        # the saved captures again
    uv run --group capture python -m capture.rotation_more build PATH  # the workbook only

It extends the ``rotation`` suite (``capture/rotation.py``) with the angled and stacked cases
that suite left out, and writes them with that suite's own ``fill`` and measures them with its
own ``_measure``, so the two read alike; ``rotation``'s cases, windows and TSV are untouched.
As there, each case is one cell in column E, from row 4, with an empty row between cases and
columns A to D and F to I empty but for what a case puts there; rows are fitted by Excel on
opening unless a case sets a height; blue (``0000C0``) text in a bordered cell, so that its ink
and the borders' are told apart, black text elsewhere. Five cases a window, each window a
workbook of its own, opened on its own, at zoom 100, in the dark appearance as the other suites;
the second marker sits in column J below the last case. Cases that set column E's width share a
window only with cases of that width (the stacked ones, 20 characters).

Cases (``ARCHITECTURE.md`` §4.17 and the rule each settles):
- ``stacked``: "Abc" stacked, left, right, center and general, in a 20-character column, so that
  the three placements lie far apart ("left and right sit at the insets"). Measured: each
  character's ink, a band of rows a line (``lines``).
- ``trigger``: text at 45 degrees with the right side alone and the top side alone: which sides
  alone make Excel slant (left alone, bottom alone and both pairs are ``rotation``'s ``slant``
  group; ``_census._boxed`` counts any side). Measured: the ink box, the sides at 25 % and 75 %
  of the row, and the top and bottom sides' extents (``top_from`` and so on: the longest dark
  horizontal run within 2 pt of the edge).
- ``neighbor``: a 45-degree cell with four thin sides whose neighbors D, F, E above and E below
  have their own thick right, left, bottom and top sides on the shared edges ("a neighbor's own
  side on an edge shared with a slanted cell is still drawn upright"); F holds an "x", so that
  the text does not spill across F's side. Measured: the dark pixels across each upright shared
  edge at the cell's middle (``edge_left`` and so on, in points: 0 where nothing is drawn, 1
  thin, 3 thick).
- ``merge``: 45-degree text with four sides in a merge across (E:F, a 70 pt row) and down (E and
  the row below, 40 and 30 pt) ("a merge's borders stay upright and its text is placed as
  without borders"). Measured as ``trigger``, from the merge's edges.
- ``top``: four sides, top-aligned in 90 pt rows at 30, 60, 75 and -45 degrees and a 45 pt row
  at 15 (in 90 pt the box would reach past column I) ("top-aligned text between slanted sides
  is moved differently", captured only at 45). Measured as ``trigger``.
- ``negative``: four sides at -15, -45 and -75 degrees, left, center and right ("at negative
  angles left, center and right are the mirror of the positive angles"). Measured as
  ``trigger``.
- ``vertical``: 45 degrees with vertical justify and distributed in a 90 pt row, with and without
  four sides ("placed as bottom"). Measured: the ink box (and the sides).
- ``underline``: a single accounting underline under text and a number at 45 and 90 degrees
  ("runs under the text's width, turned"). Measured: in the text's turned frame, the longest
  ink run along the line in the band farthest on the descent side (the text has no descenders),
  as its start and end against the rest of the ink's along the line and its distance below that
  ink (``rule_start``, ``rule_end``, ``rule_offset``).
- ``spread``: fill, justify and distributed at 15, 75, -45 and 90 degrees ("placed as left at
  angles other than 45"). Fill rows are fitted; justify and distributed wrap, so their rows are
  set as tall as a fitted row would be (``SPREAD_HEIGHTS``), which keeps the text on one line.
  Measured: the ink box.

Excel reads each case cell (its row's height and the text it shows), A, F and I in row 4 (every
case row's A, F and I take their columns from there), and the row below a merge down. ``read``
writes ``tests/captures/rotation_more.tsv``: one row a case, lengths in points from the cell's (or
merge's) left and top edges, ``ink_right`` and the sides' x past or from its left edge as
``rotation.tsv`` has them.
"""

from __future__ import annotations

import dataclasses
import itertools
import math
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Border, Font, Side

from capture import rotation
from capture.rotation import COLUMN, FIRST_ROW, PER_WINDOW, TEXT

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "rotation_more"
OUTPUT = ROOT / "tests" / "captures" / "rotation_more.tsv"
STACKED = "Abc"
STACKED_WIDTH = 20.0
NUMBER = "12345678901"
MERGE_HEIGHTS = {"across": (70.0, None), "down": (40.0, 30.0)}
"""The rows of a merged case: its own and the one below (``None``: as the other gap rows)."""
TOP_HEIGHTS = {15: 45.0, 30: 90.0, 60: 90.0, 75: 90.0, -45: 90.0}
"""The row of each top-aligned case, by angle."""
SPREAD_HEIGHTS = {15: 40.0, 75: 90.0, -45: 75.0, 90: 90.0}
"""The row of each justified and distributed case, by angle: a little over a fitted row."""
SIDE_SEARCH = 2.0
"""Points either side of a horizontal edge searched for a top or bottom side."""
SIDE_LENGTH = 10.0
"""Points a dark horizontal run must reach to count as a side, not a letter's stroke."""
EDGE_SEARCH = 4.0
"""Points either side of an upright edge searched for a neighbor's side."""
RULE_BAND = 2.0
"""Points from the ink's descent-side extreme within which an underline is looked for."""
RULE_GAP = 0.75
"""Points of ink-free gap across the line that part an underline from the letters' feet."""
COLUMNS = (
    "group", "name", "cell", "rotation", "row_height", "ink_left", "ink_right", "ink_top",
    "ink_bottom", "border_left_25", "border_right_25", "border_left_75", "border_right_75",
    "top_from", "top_to", "bottom_from", "bottom_to", "edge_left", "edge_right", "edge_top",
    "edge_bottom", "lines", "rule_start", "rule_end", "rule_offset", "orientation", "text",
)  # fmt: skip


@dataclass(frozen=True)
class Case:
    """One case: the cell as ``rotation.fill`` writes it, and what this suite adds around it."""

    group: str
    name: str
    cell: rotation.Case
    merge: str = ""  # "across": E:F; "down": E and the row below
    neighbors: bool = False  # thick sides of D, F and E above and below on the shared edges
    underline: bool = False  # a single accounting underline
    measure: tuple[str, ...] = ("ink",)  # also "sides", "edges", "lines", "rule"


def _slant(group: str, name: str, angle: int, **options: Any) -> Case:
    """Text between sides (four unless named), measured with its borders."""
    extra = {key: options.pop(key) for key in ("merge", "neighbors") if key in options}
    measure = ("ink", "sides", "edges") if extra.get("neighbors") else ("ink", "sides")
    return Case(group, name, rotation._slant(name, angle, **options), measure=measure, **extra)


def _plain(group: str, name: str, angle: int, **options: Any) -> Case:
    underline = options.pop("underline", False)
    measure = ("ink", "rule") if underline else ("ink",)
    text = options.pop("text", TEXT)
    base = rotation.Case(group, name, text, rotation._stored(angle), **options)
    return Case(group, name, base, underline=underline, measure=measure)


def cases() -> list[Case]:
    """Every case, in window order."""
    found = [
        Case(
            "stacked",
            f"stacked {horizontal}",
            rotation.Case("stacked", f"stacked {horizontal}", STACKED, 255,
                          horizontal=None if horizontal == "general" else horizontal,
                          width=STACKED_WIDTH),
            measure=("ink", "lines"),
        )
        for horizontal in ("left", "right", "center", "general")
    ]  # fmt: skip
    found += [_slant("trigger", f"45 {side}", 45, borders=side) for side in ("r", "t")]
    found.append(_slant("neighbor", "45 lrtb thick neighbors", 45, neighbors=True))
    found += [
        _slant("merge", f"45 lrtb merged {how}", 45, merge=how, height=MERGE_HEIGHTS[how][0])
        for how in ("across", "down")
    ]
    found += [
        _slant("top", f"{angle} lrtb vertical top", angle, vertical="top", height=height)
        for angle, height in TOP_HEIGHTS.items()
    ]
    found += [
        _slant("negative", f"{angle} lrtb {horizontal}", angle, horizontal=horizontal)
        for angle in (-15, -45, -75)
        for horizontal in ("left", "center", "right")
    ]
    for vertical in ("justify", "distributed"):
        found.append(_plain("vertical", f"45 vertical {vertical}", 45, vertical=vertical,
                            height=90))  # fmt: skip
        found.append(_slant("vertical", f"45 lrtb vertical {vertical}", 45, vertical=vertical,
                            height=90))  # fmt: skip
    found += [
        _plain("underline", f"{angle} accounting {kind}", angle, underline=True,
               text=NUMBER if kind == "number" else TEXT, number=kind == "number")
        for kind in ("text", "number")
        for angle in (45, 90)
    ]  # fmt: skip
    found += [
        _plain(
            "spread",
            f"{angle} {horizontal}",
            angle,
            horizontal=horizontal,
            height=None if horizontal == "fill" else SPREAD_HEIGHTS[angle],
        )
        for horizontal in ("fill", "justify", "distributed")
        for angle in (15, 75, -45, 90)
    ]
    return found


CASES = cases()

_DRAWN = """
stacked left | 47 | 0 | 3,14,6 3,29,6 3,44,5 | - | -
stacked right | 47 | 0 | 112,14,6 112,29,6 113,44,5 | - | -
stacked center | 47 | 0 | 58,14,6 58,29,6 58,44,5 | - | -
stacked general | 47 | 0 | 58,14,6 58,29,6 58,44,5 | - | -
45 r | 70 | 45 | 20.51,67.88,78 | - | 70
45 t | 70 | 45 | 20.51,67.88,78 | - | 70
45 lrtb thick neighbors | 70 | 45 | 20.51,67.88,78 | - | 70 | thick,thick,thick,thick
45 lrtb merged across | 70 | 45 | 10.78,67.88,78 | - | -
45 lrtb merged down | 40 | 45 | 10.78,67.88,78 | - | -
15 lrtb vertical top | 45 | 15 | 91.98,33.81,78 | - | 167.94
30 lrtb vertical top | 90 | 30 | 92.64,51.53,78 | - | 155.88
60 lrtb vertical top | 90 | 60 | 23.07,76.05,78 | - | 51.96
75 lrtb vertical top | 90 | 75 | 15.82,81.19,78 | - | 24.12
-45 lrtb vertical top | 90 | -45 | -44.61,10.78,78 | - | -90
-15 lrtb left | 38 | -15 | -68.77,14.91,78 | - | -141.82
-15 lrtb center | 38 | -15 | -75.11,14.91,78 | - | -141.82
-15 lrtb right | 38 | -15 | -83.39,14.91,78 | - | -141.82
-45 lrtb left | 70 | -45 | -50.91,12.72,78 | - | -70
-45 lrtb center | 70 | -45 | -36.43,12.72,78 | - | -70
-45 lrtb right | 70 | -45 | -22.66,12.72,78 | - | -70
-75 lrtb left | 85 | -75 | -15.74,8.88,78 | - | -22.78
-75 lrtb center | 85 | -75 | 1.96,8.88,78 | - | -22.78
-75 lrtb right | 85 | -75 | 19.15,8.88,78 | - | -22.78
45 vertical justify | 90 | 45 | 10.78,87.88,78 | - | -
45 lrtb vertical justify | 90 | 45 | 20.51,87.88,78 | - | 90
45 vertical distributed | 90 | 45 | 10.78,74.93,78 | - | -
45 lrtb vertical distributed | 90 | 45 | 33.45,74.93,78 | - | 90
45 accounting text | 70 | 45 | 10.78,67.88,78 | 3,-2,3 | -
90 accounting text | 85 | 90 | 46,85,78 | 3,-2,3 | -
45 accounting number | 61 | 45 | 10.78,58.88,66 | 0,0,3 | -
90 accounting number | 73 | 90 | 46,73,66 | 0,0,3 | -
15 fill | 38 | 15 | 5.85,35.1,78 | - | -
75 fill | 85 | 75 | 13.63,84.22,78 | - | -
-45 fill | 70 | -45 | -13.93,12.72,78 | - | -
90 fill | 85 | 90 | 14,85,78 | - | -
15 justify | 40 | 15 | 5.85,37.1,78 | - | -
75 justify | 90 | 75 | 13.63,89.22,78 | - | -
-45 justify | 75 | -45 | -13.93,17.72,78 | - | -
90 justify | 90 | 90 | 14,90,78 | - | -
15 distributed | 40 | 15 | 5.85,37.1,78 | - | -
75 distributed | 90 | 75 | 13.63,89.22,78 | - | -
-45 distributed | 75 | -45 | -13.93,17.72,78 | - | -
90 distributed | 90 | 90 | 30.5,90,78 | - | -
"""


def _triples(text: str) -> tuple[tuple[float, ...], ...]:
    return () if text == "-" else tuple(tuple(map(float, t.split(","))) for t in text.split())


def _table(text: str) -> dict[str, dict[str, Any]]:
    """``EXPECTED``'s entries from lines of name, height, angle, runs, rules, shift and edges."""
    found: dict[str, dict[str, Any]] = {}
    for line in text.strip().split("\n"):
        name, height, angle, runs, rules, shift, *edges = line.split(" | ")
        entry: dict[str, Any] = {
            "height": float(height),
            "angle": float(angle),
            "runs": _triples(runs),
            "rules": _triples(rules),
            "slanted": None if shift == "-" else float(shift),
        }
        if edges:
            entry["edges"] = tuple(None if s == "-" else s for s in edges[0].split(","))
        found[name] = entry
    return found


EXPECTED = _table(_DRAWN)
"""What sheetpeek lays out for each case, taken by running sheetpeek on the built workbook.

Per case, lengths in points from the left and top edges of the case's cell (or merge): the row's
``height`` (a merge's first row), the runs' ``angle``, each run's start of baseline, baseline
and width (``runs``: a stacked case has one a character), each underline's start and end
against its run's start and end and its top below the baseline (``rules``), the slanted box's
shift (``slanted``; ``None``: the sides are drawn upright), and for the neighbor case the style
drawn upright on the cell's left, right, top and bottom edges (``edges``).
"""


def windows() -> list[tuple[str, list[Case]]]:
    """Each window's folder name with its cases: cases of one column width share a window."""
    found: list[tuple[str, list[Case]]] = []
    for width in dict.fromkeys(case.cell.width for case in CASES):
        same = [case for case in CASES if case.cell.width == width]
        prefix = "default" if width is None else f"width_{width:g}"
        found += [
            (f"{prefix}_{k // PER_WINDOW}", same[k : k + PER_WINDOW])
            for k in range(0, len(same), PER_WINDOW)
        ]
    return found


def _row(k: int) -> int:
    return FIRST_ROW + 2 * k


def placements() -> list[tuple[str, str, Case]]:
    """Each case's window (its sheet in :func:`build`'s workbook) and cell, with the case."""
    return [
        (name, f"{COLUMN}{_row(k)}", case)
        for name, chosen in windows()
        for k, case in enumerate(chosen)
    ]


def area(cell: str, case: Case) -> tuple[str, str]:
    """The first and last cell of what a case draws in: the cell, or its merge."""
    row = int(cell[1:])
    if case.merge == "across":
        return cell, f"F{row}"
    if case.merge == "down":
        return cell, f"{COLUMN}{row + 1}"
    return cell, cell


def fill(ws: Any, chosen: list[Case]) -> list[tuple[str, Case]]:
    """Place ``chosen`` on a worksheet; each case's cell address with the case."""
    rotation.fill(ws, [case.cell for case in chosen])
    thick = Side(style="thick", color="FF000000")
    placed = []
    for k, case in enumerate(chosen):
        row = _row(k)
        cell = ws[f"{COLUMN}{row}"]
        if case.underline:
            cell.font = Font(name="Calibri", size=case.cell.size, underline="singleAccounting")
        if case.neighbors:
            ws[f"D{row}"].border = Border(right=thick)
            ws[f"F{row}"].border = Border(left=thick)
            ws[f"F{row}"].value = "x"
            ws[f"{COLUMN}{row - 1}"].border = Border(bottom=thick)
            ws[f"{COLUMN}{row + 1}"].border = Border(top=thick)
        if case.merge:
            first, last = area(cell.coordinate, case)
            below = MERGE_HEIGHTS[case.merge][1]
            if below is not None:
                ws.row_dimensions[row + 1].height = below
            ws.merge_cells(f"{first}:{last}")
        placed.append((cell.coordinate, case))
    return placed


def read_cells(chosen: list[Case]) -> list[str]:
    """The cells Excel reads: each case's, A, F and I in the first row, and a merge's last."""
    found = [f"{c}{FIRST_ROW}" for c in "AFI"]
    for k, case in enumerate(chosen):
        cell = f"{COLUMN}{_row(k)}"
        found += dict.fromkeys(area(cell, case))
    return list(dict.fromkeys(found))


def build(path: Path) -> list[tuple[str, str, Case]]:
    """Write every window's cases to one workbook at ``path``, a sheet a window; each case's
    sheet and cell address with the case."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    placed = []
    for name, chosen in windows():
        ws = wb.create_sheet(name)
        placed += [(name, cell, case) for cell, case in fill(ws, chosen)]
    wb.save(path)
    return placed


def run() -> None:
    """Capture every window, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    for name, chosen in windows():
        last = _row(len(chosen) - 1)
        capture(
            lambda ws, chosen=chosen: fill(ws, chosen),
            cells=read_cells(chosen),
            out=BUILD_DIR / name,
            markers=("B2", f"J{last + 2}"),
            name=f"rotation_more_{name}",
        )
    write()


# -- measuring -----------------------------------------------------------------------------------


def complete(cap: Any, chosen: list[Case]) -> Any:
    """The capture with every case row's A, F and I: their columns from the first row, their
    rows from the case's cell."""
    geometry = dict(cap.geometry)
    for k in range(len(chosen)):
        row = _row(k)
        down = geometry.get(f"{COLUMN}{row}")
        if down is None:
            continue
        for column in "AFI":
            across = geometry.get(f"{column}{FIRST_ROW}")
            if across is not None and f"{column}{row}" not in geometry:
                geometry[f"{column}{row}"] = {
                    "left": across["left"],
                    "width": across["width"],
                    "top": down["top"],
                    "height": down["height"],
                }
    return dataclasses.replace(cap, geometry=geometry)


def merged(cap: Any, cell: str, case: Case) -> Any:
    """The capture with ``cell`` spanning the case's merge, so that its edges are the merge's."""
    first, last = area(cell, case)
    if first == last:
        return cap
    a, b = cap.geometry[first], cap.geometry[last]
    geometry = dict(cap.geometry)
    geometry[cell] = {
        "left": a["left"],
        "top": a["top"],
        "width": b["left"] + b["width"] - a["left"],
        "height": b["top"] + b["height"] - a["top"],
    }
    return dataclasses.replace(cap, geometry=geometry)


def _pts(cap: Any, pixels: float) -> str:
    return f"{pixels / cap.scale:g}"


def _sides(cap: Any, cell: str) -> dict[str, str]:
    """The top and bottom sides' x extents from the cell's left: the longest dark horizontal run
    within ``SIDE_SEARCH`` of each edge, at least ``SIDE_LENGTH`` long."""
    from capture.screen.ink import rule_rows

    row = cell[1:]
    left = cap.edge(cell, "left")
    x0, x1 = round(cap.edge(f"A{row}", "left")) + 1, round(cap.edge(f"I{row}", "right")) - 1
    reach = round(SIDE_SEARCH * cap.scale)
    found: dict[str, str] = {}
    for which in ("top", "bottom"):
        y = round(cap.edge(cell, which))
        box = (x0, y - reach, x1, y + reach + 1)
        hits = rule_rows(cap, box, round(SIDE_LENGTH * cap.scale), rotation._is_dark)
        if hits:
            _, _, start, end = max(hits, key=lambda hit: hit[3] - hit[2])
            found[f"{which}_from"] = _pts(cap, start - left)
            found[f"{which}_to"] = _pts(cap, end - left)
    return found


def _edges(cap: Any, cell: str) -> dict[str, str]:
    """Points of dark pixels across each upright edge of the cell, at the cell's middle."""
    reach = round(EDGE_SEARCH * cap.scale)
    cx, cy = round(cap.edge(cell, "cx")), round(cap.edge(cell, "cy"))
    found: dict[str, str] = {}
    for which in ("left", "right", "top", "bottom"):
        at = round(cap.edge(cell, which))
        if which in ("left", "right"):
            pixels = [cap.image.pixel(x, cy) for x in range(at - reach, at + reach + 1)]
        else:
            pixels = [cap.image.pixel(cx, y) for y in range(at - reach, at + reach + 1)]
        found[f"edge_{which}"] = _pts(cap, sum(rotation._is_dark(p) for p in pixels))
    return found


def _lines(cap: Any, cell: str) -> dict[str, str]:
    """Each stacked line's ink (a band of rows holding ink), ``left:right`` from the cell's left
    and right edges, top line first."""
    from capture.screen.ink import bands, ink

    row = cell[1:]
    left, right = cap.edge(cell, "left"), cap.edge(cell, "right")
    x0, x1 = round(cap.edge(f"A{row}", "left")) + 1, round(cap.edge(f"I{row}", "right")) - 1
    y0, y1 = round(cap.edge(cell, "top")) + 1, round(cap.edge(cell, "bottom")) - 1
    found = []
    for start, length in bands(cap, (x0, y0, x1, y1), rotation._is_dark):
        box = ink(cap, (x0, start, x1, start + length), rotation._is_dark)
        assert box is not None
        found.append(f"{_pts(cap, box[0] - left)}:{_pts(cap, box[2] - right)}")
    return {"lines": ";".join(found)} if found else {}


def _rule(cap: Any, cell: str, angle: float) -> dict[str, str]:
    """An underline turned with the text: in the text's frame (``s`` along the line, ``t``
    across it toward the descent), the longest run along ``s`` in the band of ink farthest on the
    descent side, as its start and end against the rest of the ink's and its distance below it.
    """
    row = cell[1:]
    x0, x1 = round(cap.edge(f"A{row}", "left")) + 1, round(cap.edge(f"I{row}", "right")) - 1
    y0, y1 = round(cap.edge(cell, "top")) + 1, round(cap.edge(cell, "bottom")) - 1
    cos, sin = math.cos(math.radians(angle)), math.sin(math.radians(angle))
    points = [
        (x * cos - y * sin, x * sin + y * cos)
        for y in range(y0, y1)
        for x in range(x0, x1)
        if rotation._is_dark(cap.image.pixel(x, y))
    ]
    if not points:
        return {}
    across = sorted({t for _, t in points}, reverse=True)
    split = None
    for upper, lower in itertools.pairwise(across):
        if upper < across[0] - RULE_BAND * cap.scale:
            break
        if upper - lower >= RULE_GAP * cap.scale:
            split = (upper + lower) / 2
            break
    if split is None:
        return {}
    rule = [(s, t) for s, t in points if t > split]
    rest = [(s, t) for s, t in points if t < split]
    along = [s for s, _ in rule]
    if max(along) - min(along) < SIDE_LENGTH * cap.scale:
        return {}
    return {
        "rule_start": _pts(cap, min(along) - min(s for s, _ in rest)),
        "rule_end": _pts(cap, max(along) - max(s for s, _ in rest)),
        "rule_offset": _pts(cap, sum(t for _, t in rule) / len(rule) - max(t for _, t in rest)),
    }


def measure(cap: Any, cell: str, case: Case) -> dict[str, str]:
    """One case's measurements (``cap`` completed with :func:`complete`)."""
    cap = merged(cap, cell, case)
    found = rotation._measure(cap, cell, case.cell)
    if "sides" in case.measure:
        found |= _sides(cap, cell)
    if "edges" in case.measure:
        found |= _edges(cap, cell)
    if "lines" in case.measure:
        found |= _lines(cap, cell)
    if "rule" in case.measure:
        angle = case.cell.rotation if case.cell.rotation <= 90 else 90 - case.cell.rotation
        found |= _rule(cap, cell, angle)
    return found


def record(cap: Any, cell: str, case: Case) -> dict[str, str]:
    """One case's TSV row: Excel's reads of the cell and the measurements."""
    found = rotation._record(cap.rows, cell, case.cell, measure(cap, cell, case))
    return found | {"group": case.group, "name": case.name}


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``rotation_more.tsv``."""
    from capture.screen.run import header, load

    records: list[dict[str, str]] = []
    first: dict[str, str] = {}
    for name, chosen in windows():
        out = BUILD_DIR / name
        if not (out / "formats.tsv").exists():
            print(f"{name}: not captured yet, left out", file=sys.stderr)
            continue
        first = first or header(out)
        (cap,) = load(out)
        cap = complete(cap, chosen)
        for cell, case in fill(openpyxl.Workbook().active, chosen):
            records.append(record(cap, cell, case))
    if not records:
        print("nothing captured yet; no TSV written", file=sys.stderr)
        return
    with open(OUTPUT, "w", encoding="utf-8") as out_file:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out_file.write(f"# {key}: {value}\n")
        out_file.write("# suite: rotation_more\n" + "\t".join(COLUMNS) + "\n")
        for found in records:
            out_file.write("\t".join(found.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} cases written to {OUTPUT}")


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
