"""Capture how Excel for Mac draws gradient fills, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.gradients run odd
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.gradients run no_stops
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.gradients run not_finite
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.gradients run
    uv run --group capture python -m capture.gradients read         # the saved captures again
    uv run --group capture python -m capture.gradients build PATH   # every group's workbook

ARCHITECTURE.md §4.12, "Gradient fills" and "Fills against gridlines". Each group is its own
workbook and window, in the dark appearance as the other suites; ``main`` at zoom 100 and 200
(a gradient is laid out in the cell, so it should scale with the zoom, unlike a pattern's
tile), the others at zoom 100. Every gradient is written into ``xl/styles.xml`` as the case
gives it (``<gradientFill>`` text with its numbers as ``repr`` text), replacing a placeholder
that openpyxl writes, since openpyxl refuses duplicate stop positions and writes no crafted
degree. Every group has the same grid: columns C, E, ..., O are 7.29 characters wide (about
42 pt, "narrow"), Q, S, U, W and Y 23.29 (about 126 pt, "wide"), the columns between them 1;
row 4 is 126 pt tall, rows 6 and 8 42 pt, rows 5 and 7 6 pt (the widths Excel gives are read
with the capture). So a narrow cell in row 4 is tall (1 by 3), in row 6 square, a wide
cell in row 6 or 8 wide (3 by 1), in row 4 a large square. No case's neighbor to the right or
below is filled, so each gradient owns its far gridlines. The second marker is AA10.

- ``main``: linear white to blue (stops 0 and 1) at 0, 30, 45, 90, 135, 180 and 270 degrees in
  a tall (C4 to O4), a square (C6 to O6) and a wide cell (Q6, S6, U6, W6, Y6, Q8, S8): the
  angle taken in the unit cell, the corners taking the end stops. Path white to blue in large
  squares with the rectangles (left, right, top, bottom) (0, 0, 0, 0) Q4, (0.5, 0.5, 0.5, 0.5)
  S4 and (0.2, 0.8, 0.3, 0.7) U4: the bands. A first stop placed after 0: linear 90 degrees in
  large squares, white 0.25 and blue 0.75 (W4), red 0.5 and blue 1 (Y4); the same at 0 degrees
  in wide cells (U8, W8): what the color is before the first stop.
- ``odd`` (Excel may refuse it, so it runs on its own): duplicate stop positions, red then blue
  both at 0 (Q6), and white 0, red 0.5, blue 0.5 (S6), each in file order, which is not the
  order of their colors; degrees -45 and 405 in wide (U6, W6) and square cells (C6, E6), beside
  their equivalents 315 (Y6, G6) and 45 (Q8, I6) drawn in the same window.
- ``no_stops`` (may be refused): a linear (Q6) and a path gradient (S6) with no ``<stop>``.
- ``not_finite`` (may be refused): degrees ``NaN`` (Q6), ``INF`` (S6) and ``-INF`` (U6), beside
  degree 0 (W6).

A refusal (-50 with alerts off) is kept in ``refused.txt`` and written as the group's one row,
its ``name`` ``refused:`` and Excel's message.

Measured per case, in sheet points at the zoom: the cell's width and height (Excel's geometry);
the colors (``RRGGBB``, each the mean of 3 by 3 device pixels) of an 8 by 8 grid of samples at
the fractions 1/16, 3/16, ..., 15/16 of the cell across and down (``grid``, rows joined by
``/``), and of 24 samples along the horizontal and the vertical middle lines (``profile_x``,
``profile_y``, at 1/48, 3/48, ..., 47/48); where the fill's color at its right (bottom) edge is
told from white and the gridline's gray, how far past the edge it reaches (``fill_right``,
``fill_bottom``: over the far gridline or not); and Excel's display format reads of pattern and
color. ``read`` writes ``tests/captures/gradients.tsv``, one row per group, case and zoom.

The comparison (``tests/test_suite_gradients.py``) evaluates what sheetpeek's HTML draws for
each case (:data:`EXPECTED`, the CSS ``linear-gradient`` and the path's SVG bands as sheetpeek
writes them) at the same fractions, and settles: the angle in the unit cell, the path bands, the
color before a first stop past 0, duplicate stop positions, a degree out of range or not finite,
a gradient without stops, and the fill over the far gridlines.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import GradientFill

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "gradients"
OUTPUT = ROOT / "tests" / "captures" / "gradients.tsv"
GROUPS = ("main", "odd", "no_stops", "not_finite")
REFUSALS = ("odd", "no_stops", "not_finite")
"""Groups whose file Excel may refuse: run on their own, a refusal is recorded as the answer."""
ZOOMS = {"main": (100, 200), "odd": (100,), "no_stops": (100,), "not_finite": (100,)}
MARKERS = ("B2", "AA10")
WHITE, BLUE, RED = "FFFFFF", "0000FF", "FF0000"
NARROW = "CEGIKMO"
WIDE = "QSUWY"
GAPS = "DFHJLNPRTVXZ"
NARROW_WIDTH = 7.29
"""About 42 pt (56 px at 7 px a character plus 5; sheetpeek lays it out at 44 pt)."""
WIDE_WIDTH = 23.29
"""About 126 pt (168 px)."""
GAP_WIDTH = 1
"""About 9 pt (12 px)."""
HEIGHTS = {4: 126, 5: 6, 6: 42, 7: 6, 8: 42}
DEGREES = (0, 30, 45, 90, 135, 180, 270)
WIDE_CELLS = ("Q6", "S6", "U6", "W6", "Y6", "Q8", "S8")
GRID = 8
PROFILE = 24
PATCH = 1
"""Device pixels on each side of a sample's center that its mean takes in (3 by 3)."""
PLACEHOLDER = 7000
"""openpyxl writes the k-th case's fill as a gradient at degree 7000 + k, replaced afterwards."""
COLUMNS = (
    "group", "name", "zoom", "cell", "width", "height", "display_pattern", "display_fill",
    "fill_right", "fill_bottom", "grid", "profile_x", "profile_y",
)  # fmt: skip
READS = {"display_pattern": "display.pattern", "display_fill": "display.fill"}
RAMP = ((0.0, WHITE), (1.0, BLUE))


def _number(value: float) -> str:
    """A number as ``repr`` text, without a trailing ``.0``."""
    text = repr(float(value))
    return text[:-2] if text.endswith(".0") else text


@dataclass(frozen=True)
class Case:
    """One gradient cell: its group, name and cell, and the ``<gradientFill>`` it holds.

    ``degree`` is the attribute's text (``None``: no attribute); ``rect`` the path's left,
    right, top and bottom; ``stops`` the ``<stop>`` elements in file order.
    """

    group: str
    name: str
    cell: str
    kind: str = "linear"
    degree: str | None = None
    stops: tuple[tuple[float, str], ...] = RAMP
    rect: tuple[float, float, float, float] | None = None

    def xml(self) -> str:
        attributes = ' type="path"' if self.kind == "path" else ""
        if self.degree is not None:
            attributes += f' degree="{self.degree}"'
        if self.rect is not None:
            for name, value in zip(("left", "right", "top", "bottom"), self.rect, strict=True):
                attributes += f' {name}="{_number(value)}"'
        stops = "".join(
            f'<stop position="{_number(p)}"><color rgb="FF{c}"/></stop>' for p, c in self.stops
        )
        return f"<gradientFill{attributes}>{stops}</gradientFill>"


def _cases() -> list[Case]:
    found = []
    for k, degree in enumerate(DEGREES):
        text = str(degree) if degree else None
        found.append(Case("main", f"linear {degree} tall", f"{NARROW[k]}4", degree=text))
        found.append(Case("main", f"linear {degree} square", f"{NARROW[k]}6", degree=text))
        found.append(Case("main", f"linear {degree} wide", WIDE_CELLS[k], degree=text))
    for cell, rect in (
        ("Q4", (0.0, 0.0, 0.0, 0.0)),
        ("S4", (0.5, 0.5, 0.5, 0.5)),
        ("U4", (0.2, 0.8, 0.3, 0.7)),
    ):
        label = ",".join(_number(v) for v in rect)
        found.append(Case("main", f"path {label}", cell, kind="path", rect=rect))
    late = ((0.25, WHITE), (0.75, BLUE))
    half = ((0.5, RED), (1.0, BLUE))
    found += [
        Case("main", "first stop 0.25, 90", "W4", degree="90", stops=late),
        Case("main", "first stop 0.5, 90", "Y4", degree="90", stops=half),
        Case("main", "first stop 0.25, 0", "U8", stops=late),
        Case("main", "first stop 0.5, 0", "W8", stops=half),
    ]
    found += [
        Case("odd", "two stops at 0, red then blue", "Q6", stops=((0.0, RED), (0.0, BLUE))),
        Case(
            "odd",
            "white 0, red 0.5, blue 0.5",
            "S6",
            stops=((0.0, WHITE), (0.5, RED), (0.5, BLUE)),
        ),
        Case("odd", "linear -45 wide", "U6", degree="-45"),
        Case("odd", "linear 405 wide", "W6", degree="405"),
        Case("odd", "linear 315 wide", "Y6", degree="315"),
        Case("odd", "linear 45 wide", "Q8", degree="45"),
        Case("odd", "linear -45 square", "C6", degree="-45"),
        Case("odd", "linear 405 square", "E6", degree="405"),
        Case("odd", "linear 315 square", "G6", degree="315"),
        Case("odd", "linear 45 square", "I6", degree="45"),
    ]
    found += [
        Case("no_stops", "linear without stops", "Q6", stops=()),
        Case(
            "no_stops", "path without stops", "S6", kind="path", rect=(0.5, 0.5, 0.5, 0.5), stops=()
        ),
    ]
    found += [
        Case("not_finite", "linear NaN", "Q6", degree="NaN"),
        Case("not_finite", "linear INF", "S6", degree="INF"),
        Case("not_finite", "linear -INF", "U6", degree="-INF"),
        Case("not_finite", "linear 0", "W6"),
    ]
    return found


CASES = _cases()
CONTROLS = {
    "odd": {"U6": "Y6", "W6": "Q8", "C6": "G6", "E6": "I6"},
    "not_finite": {"Q6": "W6", "S6": "W6", "U6": "W6"},
}
"""For a case whose degree is out of range or not finite, the cell drawing its equivalent."""

EXPECTED: dict[str, dict[str, dict[str, Any] | None]] = {
    "main": {
        "C4": {
            "kind": "linear",
            "box": (60, 169),
            "angle": 90.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "C6": {
            "kind": "linear",
            "box": (60, 57),
            "angle": 90.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "Q6": {
            "kind": "linear",
            "box": (187, 57),
            "angle": 90.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "E4": {
            "kind": "linear",
            "box": (59, 169),
            "angle": 101.396,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "E6": {
            "kind": "linear",
            "box": (59, 57),
            "angle": 120.863,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "S6": {
            "kind": "linear",
            "box": (188, 57),
            "angle": 152.294,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "G4": {
            "kind": "linear",
            "box": (60, 169),
            "angle": 109.546,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "G6": {
            "kind": "linear",
            "box": (60, 57),
            "angle": 136.469,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "U6": {
            "kind": "linear",
            "box": (188, 57),
            "angle": 163.133,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "I4": {
            "kind": "linear",
            "box": (60, 169),
            "angle": 180.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "I6": {
            "kind": "linear",
            "box": (60, 57),
            "angle": 180.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "W6": {
            "kind": "linear",
            "box": (187, 57),
            "angle": 180.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "K4": {
            "kind": "linear",
            "box": (59, 169),
            "angle": 250.755,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "K6": {
            "kind": "linear",
            "box": (59, 57),
            "angle": 224.012,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "Y6": {
            "kind": "linear",
            "box": (188, 57),
            "angle": 196.867,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "M4": {
            "kind": "linear",
            "box": (60, 169),
            "angle": 270.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "M6": {
            "kind": "linear",
            "box": (60, 57),
            "angle": 270.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "Q8": {
            "kind": "linear",
            "box": (187, 57),
            "angle": 270.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "O4": {
            "kind": "linear",
            "box": (60, 169),
            "angle": 360.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "O6": {
            "kind": "linear",
            "box": (60, 57),
            "angle": 360.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "S8": {
            "kind": "linear",
            "box": (188, 57),
            "angle": 360.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "Q4": {
            "kind": "path",
            "box": (187, 169),
            "rect": (0.0, 0.0, 0.0, 0.0),
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
            "inner": None,
        },
        "S4": {
            "kind": "path",
            "box": (188, 169),
            "rect": (94.0, 84.5, 94.0, 84.5),
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
            "inner": None,
        },
        "U4": {
            "kind": "path",
            "box": (188, 169),
            "rect": (37.6, 50.699999999999996, 150.4, 118.3),
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
            "inner": "FFFFFF",
        },
        "W4": {
            "kind": "linear",
            "box": (187, 169),
            "angle": 180.0,
            "stops": ((0.0, "0000FF"), (0.25, "FFFFFF"), (0.75, "0000FF")),
        },
        "Y4": {
            "kind": "linear",
            "box": (188, 169),
            "angle": 180.0,
            "stops": ((0.0, "0000FF"), (0.5, "FF0000"), (1.0, "0000FF")),
        },
        "U8": {
            "kind": "linear",
            "box": (188, 57),
            "angle": 90.0,
            "stops": ((0.0, "0000FF"), (0.25, "FFFFFF"), (0.75, "0000FF")),
        },
        "W8": {
            "kind": "linear",
            "box": (187, 57),
            "angle": 90.0,
            "stops": ((0.0, "0000FF"), (0.5, "FF0000"), (1.0, "0000FF")),
        },
    },
    "odd": {
        "Q6": {
            "kind": "linear",
            "box": (187, 57),
            "angle": 90.0,
            "stops": ((0.0, "FF0000"), (0.0, "0000FF")),
        },
        "S6": {
            "kind": "linear",
            "box": (188, 57),
            "angle": 90.0,
            "stops": ((0.0, "FFFFFF"), (0.5, "FF0000"), (0.5, "0000FF")),
        },
        "U6": {
            "kind": "linear",
            "box": (188, 57),
            "angle": 16.867,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "W6": {
            "kind": "linear",
            "box": (187, 57),
            "angle": 163.048,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "Y6": {
            "kind": "linear",
            "box": (188, 57),
            "angle": 16.867,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "Q8": {
            "kind": "linear",
            "box": (187, 57),
            "angle": 163.048,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "C6": {
            "kind": "linear",
            "box": (60, 57),
            "angle": 43.531,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "E6": {
            "kind": "linear",
            "box": (59, 57),
            "angle": 135.988,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "G6": {
            "kind": "linear",
            "box": (60, 57),
            "angle": 43.531,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "I6": {
            "kind": "linear",
            "box": (60, 57),
            "angle": 136.469,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
    },
    "no_stops": {
        "Q6": None,
        "S6": None,
    },
    "not_finite": {
        "Q6": {
            "kind": "linear",
            "box": (188, 57),
            "angle": 90.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "S6": {
            "kind": "linear",
            "box": (187, 57),
            "angle": 90.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "U6": {
            "kind": "linear",
            "box": (188, 57),
            "angle": 90.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
        "W6": {
            "kind": "linear",
            "box": (188, 57),
            "angle": 90.0,
            "stops": ((0.0, "FFFFFF"), (1.0, "0000FF")),
        },
    },
}
"""What sheetpeek's HTML draws for each case, taken by running sheetpeek on the built workbooks.

A linear gradient's box (CSS px, the cell and its far gridlines), CSS angle and stops; a path's box,
rectangle (left, top, right, bottom in px), stops and the rectangle's own color (None: it
has no area); None: no fill. A first stop past 0 shows as the stop sheetpeek adds at 0.
"""


def cases(group: str | None = None) -> list[Case]:
    """The cases of one group, or all of them."""
    return [c for c in CASES if group is None or c.group == group]


def fill(ws: Any, group: str) -> list[Case]:
    """Lay one group out on a worksheet, each gradient as its placeholder; its cases."""
    for column in NARROW:
        ws.column_dimensions[column].width = NARROW_WIDTH
    for column in WIDE:
        ws.column_dimensions[column].width = WIDE_WIDTH
    for column in GAPS:
        ws.column_dimensions[column].width = GAP_WIDTH
    for row, height in HEIGHTS.items():
        ws.row_dimensions[row].height = height
    found = cases(group)
    for case in found:
        k = CASES.index(case)
        ws[case.cell].fill = GradientFill(degree=PLACEHOLDER + k, stop=("FFFFFFFF", "FF0000FF"))
    return found


def styles(text: str, group: str) -> str:
    """``xl/styles.xml`` with each placeholder gradient replaced by its case's own."""
    for case in cases(group):
        k = CASES.index(case)
        pattern = rf'<gradientFill type="linear" degree="{PLACEHOLDER + k}">.*?</gradientFill>'
        text, count = re.subn(pattern, lambda _, case=case: case.xml(), text)
        if count != 1:
            raise ValueError(f"{case.cell}: its placeholder gradient was not found")
    return text


def postprocess(group: str) -> Any:
    """Write a group's gradients into the saved workbook (``patch_part``)."""

    def edit(path: Path) -> None:
        from capture.screen.run import patch_part

        patch_part(path, {"xl/styles.xml": lambda t: styles(t, group)})

    return edit


def build(path: Path) -> list[Case]:
    """Write every group to a workbook of its own in the folder ``path``; every case."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        found.extend(fill(wb.active, group))
        target = path / f"gradients_{group}.xlsx"
        wb.save(target)
        postprocess(group)(target)
    return found


def run(groups: tuple[str, ...] | None = None) -> None:
    """Capture each group's window, then write the TSV (needs Excel).

    Without groups, every group but :data:`REFUSALS`, which run on their own.
    """
    from capture.screen.run import capture
    from capture.screen.session import ExcelError

    if groups is None:
        groups = tuple(g for g in GROUPS if g not in REFUSALS)
    for group in groups:
        try:
            capture(
                lambda ws, group=group: fill(ws, group),
                cells=[c.cell for c in cases(group)],
                out=BUILD_DIR / group,
                markers=MARKERS,
                zooms=ZOOMS[group],
                name=f"gradients_{group}",
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


# -- measuring -----------------------------------------------------------------------------------


def fractions(count: int) -> list[float]:
    """``count`` sample places across a cell: the middles of ``count`` equal parts."""
    return [(2 * k + 1) / (2 * count) for k in range(count)]


def sample(cap: Any, x: float, y: float) -> str:
    """The mean color of the 3 by 3 device pixels around ``(x, y)``, ``RRGGBB``."""
    cx, cy = round(x), round(y)
    pixels = [
        cap.image.pixel(cx + dx, cy + dy)
        for dy in range(-PATCH, PATCH + 1)
        for dx in range(-PATCH, PATCH + 1)
    ]
    return "".join(f"{round(sum(p[i] for p in pixels) / len(pixels)):02X}" for i in range(3))


def _at(cap: Any, cell: str, fx: float, fy: float) -> str:
    left, right = cap.edge(cell, "left"), cap.edge(cell, "right")
    top, bottom = cap.edge(cell, "top"), cap.edge(cell, "bottom")
    return sample(cap, left + fx * (right - left), top + fy * (bottom - top))


def _filled(cap: Any, rgb: tuple[int, int, int]) -> bool:
    """A pixel told from the sheet's white and the gridline's gray: the fill's."""
    from capture.screen.ink import nearest

    palette = {"white": (255, 255, 255)}
    if cap.gridline is not None:
        palette["grid"] = cap.gridline
    return nearest(rgb, palette, tolerance=24) == "?"


def reach(cap: Any, cell: str, edge: str) -> str:
    """How far past the cell's right (bottom) edge its fill runs, in points; ``""`` if unseen.

    Taken along the middle line from 2 device pixels inside the edge, while pixels are told
    from white and the gridline's gray, at most 4 screen points.
    """
    per = cap.scale * 100 / cap.zoom  # image pixels a screen point
    if edge == "right":
        start, at = round(cap.edge(cell, "right")), round(cap.edge(cell, "cy"))

        def pixel(k: int) -> tuple[int, int, int]:
            return cap.image.pixel(start + k, at)
    else:
        start, at = round(cap.edge(cell, "bottom")), round(cap.edge(cell, "cx"))

        def pixel(k: int) -> tuple[int, int, int]:
            return cap.image.pixel(at, start + k)

    if cap.gridline is None or not _filled(cap, pixel(-2)):
        return ""
    k = -2
    while k < round(4 * per) and _filled(cap, pixel(k + 1)):
        k += 1
    return f"{round((k + 1) / cap.scale, 3):g}"


def measure(cap: Any, case: Case) -> dict[str, str] | None:
    """One case's measurements at one zoom, or ``None`` when it is outside the window."""
    from capture.screen.ink import pt, visible

    cell = case.cell
    if not visible(cap, cell):
        return None
    left, right = cap.edge(cell, "left"), cap.edge(cell, "right")
    top, bottom = cap.edge(cell, "top"), cap.edge(cell, "bottom")
    places = fractions(GRID)
    line = fractions(PROFILE)
    return {
        "width": pt(cap, right - left),
        "height": pt(cap, bottom - top),
        "fill_right": reach(cap, cell, "right"),
        "fill_bottom": reach(cap, cell, "bottom"),
        "grid": "/".join(" ".join(_at(cap, cell, fx, fy) for fx in places) for fy in places),
        "profile_x": " ".join(_at(cap, cell, f, 0.5) for f in line),
        "profile_y": " ".join(_at(cap, cell, 0.5, f) for f in line),
    }


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``gradients.tsv``."""
    from capture.screen.run import header, load

    records = []
    first: dict[str, str] = {}
    for group in GROUPS:
        out = BUILD_DIR / group
        refused = out / "refused.txt"
        if refused.exists():
            reason = refused.read_text(encoding="utf-8").strip()
            records.append({"group": group, "name": f"refused: {reason}"})
            continue
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
                        "group": group,
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
        out.write("# suite: gradients\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cases in {argv[1]}")
        return 0
    if argv[:1] == ["run"] and argv[1:] and all(group in GROUPS for group in argv[1:]):
        run(tuple(argv[1:]))
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
