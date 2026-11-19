"""Capture how Excel for Mac draws diagonal borders, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.diagonals run
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.diagonals run large
    uv run --group capture python -m capture.diagonals read         # the saved captures again
    uv run --group capture python -m capture.diagonals build PATH   # every group's workbook

ARCHITECTURE.md §4.12, "Borders" (a diagonal is read only with a style and a direction) and the
diagonals' rule: they run corner to corner, antialiased; hair is drawn as thin, the dashed styles
as short blocks; text and the line both show. A merge's diagonal is another suite's
(``merges_spill``).

Two groups, each its own workbook and window, at zoom 100 and 200, in the dark appearance as the
other suites. ``small`` draws in Excel's default cells (53 by 15 pt on the Mac), ``large`` in
cells about 100 by 40 pt (columns 16.67 characters, rows 40 pt). In each, the 13 line styles,
black, take the case columns C, E, ..., O (the columns between them 1 character wide, the rows
between 6 pt): the first seven styles up in row 4, down in row 6 and both in row 8, the other
six likewise in rows 10, 12 and 14. Row 16 holds the cases with text ("Diagonal" in blue
``0000C0``, so that its ink is told from the line's) and, in ``large``, a diagonal with a style
and no direction (K16) and one with both directions and no style (M16). The second marker is
P18; at zoom 200 the ``large`` window needs about 1850 points across and 800 down to show it,
the ``small`` one about 1200 by 450 (a smaller window needs only the first marker,
``CAPTURING.md``, "One marker, on a small screen"; a case outside the window is left out).

Excel reads only the case cells in row 4, the ones in column C and row 16 (:func:`read_cells`):
every other case cell's geometry is its column's from row 4 and its row's from column C
(:func:`complete`), which keeps the reads to 17 to 20 a window, markers included.

Measured per case and line (``up``, ``down``; a cell with both has a record for each), from
image columns across the line's expected path (it is shallower than 45 degrees in every cell),
a window of 3.5 screen points either side of the corner-to-corner line: where its ink (luma
under 160) starts and ends against the cell's left and right boundaries (``start``, ``end``, in
sheet points at the zoom; negative is to the left); in the columns clear of the gridlines, of
the other line and of the text, the median vertical thickness (``coverage``, the sum of each
pixel's darkness, ``(255 - luma) / 255``) and offset of the ink's center from the
corner-to-corner line (``offset``, down positive), and the most common number of separate ink
runs across the line (``bands``: 2 for a double line); along the line's columns, for the cells
with one direction, the on/off pattern of one period (``pattern``, in sheet points along x, as the
borders suite gives it; a slanted dash covers a fractional number of columns, so a dashed line
often reads ``irregular``) and every whole run between the partial ones at the ends (``runs``,
the same units); the darkest ink. For the text cases, the text's ink box from the
cell's top left (``text_box``) and, inside that box, what each pixel on the line's path shows
(``crossing``: black, blue, white or other). Excel's own reads of the diagonal styles go in
``read_up`` and ``read_down`` where the cell was read. ``read`` writes
``tests/captures/diagonals.tsv``, one row per group, zoom, case and line.

The comparison (``tests/test_suite_diagonals.py``) takes sheetpeek's SVG line for each case
(:data:`EXPECTED`: its box in CSS px, stroke width and dashes in screen points, which lines, a
double's two lines) and settles: corner to corner, the widths, hair against thin, the dashes,
lines fixed in screen points at zoom 200, text and line both showing, and nothing drawn without a
style or a direction.
"""

from __future__ import annotations

import dataclasses
import math
import statistics
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "diagonals"
OUTPUT = ROOT / "tests" / "captures" / "diagonals.tsv"
STYLES = (
    "thin", "medium", "thick", "double", "hair", "dotted", "dashed", "dashDot", "dashDotDot",
    "mediumDashed", "mediumDashDot", "mediumDashDotDot", "slantDashDot",
)  # fmt: skip
DIRECTIONS = ("up", "down", "both")
GROUPS = ("small", "large")
ZOOMS = (100, 200)
CASE_COLUMNS = "CEGIKMO"
GAP_COLUMNS = "DFHJLN"
GAP_WIDTH = 1
LARGE_WIDTH = 16.67
"""About 100 pt on the Mac (6 pt a character; sheetpeek lays it out at 133 px, 99.75 pt)."""
LARGE_HEIGHT = 40
GAP_HEIGHT = 6
EXTRA_ROW = 16
MARKERS = ("B2", "P18")
BLACK = "FF000000"
TEXT = "Diagonal"
TEXT_COLOR = "FF0000C0"
INK = 160
"""Luma below which a pixel is the line's ink (the Mac's gridlines are 212)."""
HALF = 3.5
"""Screen points scanned on either side of the corner-to-corner line."""
CLEAR = 3
"""Image pixels kept clear of the cell's gridlines (each is 2 image pixels from its boundary)."""
COLUMNS = (
    "group", "name", "zoom", "cell", "style", "direction", "line", "width", "height", "read_up",
    "read_down", "ink", "start", "end", "offset", "coverage", "bands", "pattern", "runs",
    "text_box", "crossing",
)  # fmt: skip


@dataclass(frozen=True)
class Case:
    """One cell: its diagonal's style (``None``: none) and flags, and any text."""

    group: str
    name: str
    cell: str
    style: str | None
    up: bool
    down: bool
    text: str = ""
    centered: bool = False

    @property
    def direction(self) -> str:
        return {(True, False): "up", (False, True): "down", (True, True): "both"}.get(
            (self.up, self.down), "none"
        )

    @property
    def lines(self) -> tuple[str, ...]:
        """The lines measured: the ones drawn, or both where the case should draw none."""
        drawn = tuple(line for line, on in (("up", self.up), ("down", self.down)) if on)
        return drawn if drawn and self.style is not None else ("up", "down")


def _place(k: int, direction: str) -> str:
    """The cell of the ``k``-th style drawn in ``direction``."""
    row = 4 + 2 * (3 * (k // len(CASE_COLUMNS)) + DIRECTIONS.index(direction))
    return f"{CASE_COLUMNS[k % len(CASE_COLUMNS)]}{row}"


def _cases() -> list[Case]:
    found = [
        Case(
            group,
            f"{style} {direction}",
            _place(k, direction),
            style,
            up=direction in ("up", "both"),
            down=direction in ("down", "both"),
        )
        for group in GROUPS
        for k, style in enumerate(STYLES)
        for direction in DIRECTIONS
    ]
    row = EXTRA_ROW
    found += [
        Case("large", "thin both, text centered", f"C{row}", "thin", True, True, TEXT, True),
        Case("large", "thick both, text centered", f"E{row}", "thick", True, True, TEXT, True),
        Case("large", "double up, text", f"G{row}", "double", True, False, TEXT),
        Case("large", "dashed down, text", f"I{row}", "dashed", False, True, TEXT),
        Case("large", "thin, no direction", f"K{row}", "thin", False, False),
        Case("large", "no style, up and down", f"M{row}", None, True, True),
        Case("small", "thin both, text", f"C{row}", "thin", True, True, TEXT),
        Case("small", "medium both, text centered", f"E{row}", "medium", True, True, TEXT, True),
        Case("small", "dashDot down, text", f"G{row}", "dashDot", False, True, TEXT),
    ]
    return found


CASES = _cases()

_DRAWN = {
    "small": """
C4 72x21 up 1
C6 72x21 down 1
C8 72x21 both 1
E4 71x21 up 2
E6 71x21 down 2
E8 71x21 both 2
G4 72x21 up 3
G6 72x21 down 3
G8 72x21 both 3
I4 72x21 up 1 double
I6 72x21 down 1 double
I8 72x21 both 1 double
K4 71x21 up 1
K6 71x21 down 1
K8 71x21 both 1
M4 72x21 up 1 1,1
M6 72x21 down 1 1,1
M8 72x21 both 1 1,1
O4 72x21 up 1 2,2
O6 72x21 down 1 2,2
O8 72x21 both 1 2,2
C10 72x21 up 1 9,3,3,3
C12 72x21 down 1 9,3,3,3
C14 72x21 both 1 9,3,3,3
E10 71x21 up 1 9,3,3,3,3,3
E12 71x21 down 1 9,3,3,3,3,3
E14 71x21 both 1 9,3,3,3,3,3
G10 72x21 up 2 9,3
G12 72x21 down 2 9,3
G14 72x21 both 2 9,3
I10 72x21 up 2 9,3,3,3
I12 72x21 down 2 9,3,3,3
I14 72x21 both 2 9,3,3,3
K10 71x21 up 2 9,3,3,3,3,3
K12 71x21 down 2 9,3,3,3,3,3
K14 71x21 both 2 9,3,3,3,3,3
M10 72x21 up 2 9,3,3,3
M12 72x21 down 2 9,3,3,3
M14 72x21 both 2 9,3,3,3
C16 72x21 both 1
E16 71x21 both 2
G16 72x21 down 1 9,3,3,3
""",
    "large": """
C4 134x54 up 1
C6 134x55 down 1
C8 134x54 both 1
E4 135x54 up 2
E6 135x55 down 2
E8 135x54 both 2
G4 134x54 up 3
G6 134x55 down 3
G8 134x54 both 3
I4 134x54 up 1 double
I6 134x55 down 1 double
I8 134x54 both 1 double
K4 135x54 up 1
K6 135x55 down 1
K8 135x54 both 1
M4 134x54 up 1 1,1
M6 134x55 down 1 1,1
M8 134x54 both 1 1,1
O4 134x54 up 1 2,2
O6 134x55 down 1 2,2
O8 134x54 both 1 2,2
C10 134x54 up 1 9,3,3,3
C12 134x55 down 1 9,3,3,3
C14 134x54 both 1 9,3,3,3
E10 135x54 up 1 9,3,3,3,3,3
E12 135x55 down 1 9,3,3,3,3,3
E14 135x54 both 1 9,3,3,3,3,3
G10 134x54 up 2 9,3
G12 134x55 down 2 9,3
G14 134x54 both 2 9,3
I10 134x54 up 2 9,3,3,3
I12 134x55 down 2 9,3,3,3
I14 134x54 both 2 9,3,3,3
K10 135x54 up 2 9,3,3,3,3,3
K12 135x55 down 2 9,3,3,3,3,3
K14 135x54 both 2 9,3,3,3,3,3
M10 134x54 up 2 9,3,3,3
M12 134x55 down 2 9,3,3,3
M14 134x54 both 2 9,3,3,3
C16 134x54 both 1
E16 135x54 both 3
G16 134x54 up 1 double
I16 134x54 down 1 2,2
K16 none
M16 none
""",
}


def _table(text: str) -> dict[str, dict[str, Any] | None]:
    """``EXPECTED``'s entries from lines of cell, box, lines, stroke, dashes and ``double``."""
    found: dict[str, dict[str, Any] | None] = {}
    for line in text.split("\n"):
        if not line:
            continue
        cell, *rest = line.split()
        if rest == ["none"]:
            found[cell] = None
            continue
        box, lines, stroke, *more = rest
        width, height = (int(v) for v in box.split("x"))
        dashes = next((m.replace(",", " ") for m in more if m != "double"), "")
        found[cell] = {
            "box": (width, height),
            "up": lines in ("up", "both"),
            "down": lines in ("down", "both"),
            "double": "double" in more,
            "stroke": float(stroke),
            "dashes": dashes,
        }
    return found


EXPECTED = {group: _table(text) for group, text in _DRAWN.items()}
"""What sheetpeek's HTML draws for each case, taken by running sheetpeek on the built workbooks.

Per cell: the line's ``box`` (CSS px: the cell and its far gridlines), ``stroke`` (CSS px, which
are screen points) and ``dashes`` (the SVG dash array, ``""`` for solid), whether it draws
``up`` and ``down``, and ``double`` (two lines 1 px either side); ``None``: nothing drawn.
``_DRAWN`` holds them one cell a line: cell, box, lines, stroke, dashes (commas), ``double``.
"""


def cases(group: str | None = None) -> list[Case]:
    """The cases of one group, or all of them."""
    return [c for c in CASES if group is None or c.group == group]


def read_cells(group: str) -> list[str]:
    """The cells Excel reads: the case cells in row 4, in column C and in the last row."""
    found = [c.cell for c in cases(group)]  # one-letter columns
    return [
        cell for cell in found if cell[1:] == "4" or cell[0] == "C" or cell[1:] == str(EXTRA_ROW)
    ]


def fill(ws: Any, group: str) -> list[Case]:
    """Lay one group out on a worksheet; its cases."""
    if group == "large":
        for column in CASE_COLUMNS:
            ws.column_dimensions[column].width = LARGE_WIDTH
    for column in GAP_COLUMNS:
        ws.column_dimensions[column].width = GAP_WIDTH
    for row in range(4, EXTRA_ROW + 1, 2):
        if group == "large":
            ws.row_dimensions[row].height = LARGE_HEIGHT
        ws.row_dimensions[row + 1].height = GAP_HEIGHT
    found = cases(group)
    for case in found:
        cell = ws[case.cell]
        side = Side(style=case.style, color=BLACK if case.style else None)
        cell.border = Border(diagonal=side, diagonalUp=case.up, diagonalDown=case.down)
        if case.text:
            cell.value = case.text
            cell.font = Font(color=TEXT_COLOR)
            if case.centered:
                cell.alignment = Alignment(horizontal="center", vertical="center")
    return found


def build(path: Path) -> list[Case]:
    """Write every group to a workbook of its own in the folder ``path``; every case."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        found.extend(fill(wb.active, group))
        wb.save(path / f"diagonals_{group}.xlsx")
    return found


def run(groups: tuple[str, ...] = GROUPS) -> None:
    """Capture each group's window at both zooms, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    for group in groups:
        capture(
            lambda ws, group=group: fill(ws, group),
            cells=read_cells(group),
            out=BUILD_DIR / group,
            markers=MARKERS,
            zooms=ZOOMS,
            name=f"diagonals_{group}",
        )
    write()


# -- measuring -----------------------------------------------------------------------------------


def complete(cap: Any, group: str) -> Any:
    """The capture with every case cell's geometry: its column's from row 4, its row's from C."""
    geometry = dict(cap.geometry)
    for case in cases(group):
        if case.cell in geometry:
            continue
        column, row = case.cell[0], case.cell[1:]
        across, down = geometry.get(f"{column}4"), geometry.get(f"C{row}")
        if across is None or down is None:
            continue
        geometry[case.cell] = {
            "left": across["left"],
            "width": across["width"],
            "top": down["top"],
            "height": down["height"],
        }
    return dataclasses.replace(cap, geometry=geometry)


def _pixel(cap: Any, x: int, y: int) -> tuple[int, int, int]:
    """An image pixel, or white outside the image, which a line's window may reach past."""
    if 0 <= x < cap.image.width and 0 <= y < cap.image.height:
        return cap.image.pixel(x, y)
    return (255, 255, 255)


def is_text(rgb: tuple[int, int, int]) -> bool:
    """The blue text's ink, antialiased or not: blue well above red and green."""
    return rgb[2] - max(rgb[0], rgb[1]) > 40


def _ideal(cap: Any, cell: str, line: str) -> Any:
    """The corner-to-corner line's y (image pixels) at an image x."""
    left, right = cap.edge(cell, "left"), cap.edge(cell, "right")
    top, bottom = cap.edge(cell, "top"), cap.edge(cell, "bottom")
    slope = (bottom - top) / (right - left)

    def y(x: float) -> float:
        return bottom - (x - left) * slope if line == "up" else top + (x - left) * slope

    return y


def _columns(cap: Any, case: Case, line: str) -> list[dict[str, Any]]:
    """Each image column across the line's path, from a window before the cell to one after."""
    from capture.screen.run import luma

    cell = case.cell
    left, right = cap.edge(cell, "left"), cap.edge(cell, "right")
    top, bottom = cap.edge(cell, "top"), cap.edge(cell, "bottom")
    half = math.ceil(HALF * cap.scale * 100 / cap.zoom)
    y = _ideal(cap, cell, line)
    other = _ideal(cap, cell, "down" if line == "up" else "up")
    both = case.up and case.down and case.style is not None
    found = []
    for x in range(round(left) - half, round(right) + half):
        center = y(x + 0.5)
        y0, y1 = round(center) - half, round(center) + half + 1
        pixels = [_pixel(cap, x, yy) for yy in range(y0, y1)]
        text = [is_text(p) for p in pixels]
        flags = [not t and luma(p) < INK for p, t in zip(pixels, text, strict=True)]
        darkness = [
            0.0 if t else max(0.0, 255 - luma(p)) / 255 for p, t in zip(pixels, text, strict=True)
        ]
        clear = (
            left + CLEAR <= x < right - CLEAR
            and y0 >= top + CLEAR
            and y1 <= bottom - CLEAR
            and not any(text)
            and not (both and abs(center - other(x + 0.5)) < 2 * half + 2)
        )
        found.append(
            {
                "x": x,
                "center": center,
                "y0": y0,
                "pixels": pixels,
                "flags": flags,
                "darkness": darkness,
                "clear": clear,
            }
        )
    return found


def _bands(flags: list[bool]) -> int:
    return sum(1 for a, b in zip([False, *flags], flags, strict=False) if b and not a)


def all_runs(flags: list[bool], scale: float) -> str:
    """Every run of a scanline but the partial ones at its ends: ``on 3 off 1 on 3 ...``, points."""
    from capture.screen.run import runs

    inner = runs(flags)[1:-1]
    return " ".join(f"{'on' if on else 'off'} {length / scale:g}" for on, _, length in inner)


def measure_line(cap: Any, case: Case, line: str) -> dict[str, str]:
    """One line of a case at one zoom (see the module's docstring)."""
    from capture.screen.ink import pattern, pt

    cell = case.cell
    left, right = cap.edge(cell, "left"), cap.edge(cell, "right")
    columns = _columns(cap, case, line)
    found: dict[str, str] = {}
    inked = [c for c in columns if any(c["flags"])]
    if inked:
        found["start"] = pt(cap, inked[0]["x"] - left)
        found["end"] = pt(cap, inked[-1]["x"] + 1 - right)
        darkest = min(
            (p for c in inked for p, f in zip(c["pixels"], c["flags"], strict=True) if f),
            key=sum,
        )
        found["ink"] = "".join(f"{v:02X}" for v in darkest)
    measured = [c for c in inked if c["clear"]]
    if measured:
        coverage = [sum(c["darkness"]) for c in measured]
        offsets = [
            sum((c["y0"] + k + 0.5) * d for k, d in enumerate(c["darkness"])) / sum(c["darkness"])
            - c["center"]
            for c in measured
            if sum(c["darkness"]) > 0
        ]
        found["coverage"] = pt(cap, statistics.median(coverage))
        found["offset"] = pt(cap, statistics.median(offsets))
        found["bands"] = str(Counter(_bands(c["flags"]) for c in measured).most_common(1)[0][0])
    if not (case.up and case.down) and case.style is not None:
        flags = [any(c["flags"]) for c in columns if left + CLEAR <= c["x"] < right - CLEAR]
        if any(flags):
            found["pattern"] = pattern(flags, cap.scale)
            found["runs"] = all_runs(flags, cap.scale)
    if case.text:
        found.update(_text(cap, case, line))
    return found


def _text(cap: Any, case: Case, line: str) -> dict[str, str]:
    """The text's ink box from the cell's top left, and what the line's path shows across it."""
    from capture.screen.ink import cells_box, ink, pt
    from capture.screen.run import luma

    cell = case.cell
    box = ink(cap, cells_box(cap, cell, inset=CLEAR), is_text)
    if box is None:
        return {}
    left, top = cap.edge(cell, "left"), cap.edge(cell, "top")
    x0, y0, x1, y1 = box
    y = _ideal(cap, cell, line)
    shown: Counter[str] = Counter()
    for x in range(x0, x1):
        at = round(y(x + 0.5) - 0.5)
        if not y0 <= at < y1:
            continue
        rgb = cap.image.pixel(x, at)
        if is_text(rgb):
            shown["blue"] += 1
        elif luma(rgb) < INK:
            shown["black"] += 1
        elif min(rgb) >= 230:
            shown["white"] += 1
        else:
            shown["other"] += 1
    return {
        "text_box": " ".join(pt(cap, v) for v in (x0 - left, y0 - top, x1 - left, y1 - top)),
        "crossing": " ".join(f"{k} {shown[k]}" for k in ("black", "blue", "white", "other")),
    }


def measure(cap: Any, group: str) -> list[dict[str, str]]:
    """Every case's records at one zoom; cases outside the window are left out."""
    from capture.screen.ink import pt, visible

    cap = complete(cap, group)
    records = []
    for case in cases(group):
        if case.cell not in cap.geometry or not visible(cap, case.cell):
            continue
        left, right = cap.edge(case.cell, "left"), cap.edge(case.cell, "right")
        top, bottom = cap.edge(case.cell, "top"), cap.edge(case.cell, "bottom")
        records.extend(
            {
                "group": group,
                "name": case.name,
                "zoom": str(cap.zoom),
                "cell": case.cell,
                "style": case.style or "",
                "direction": case.direction,
                "line": line,
                "width": pt(cap, right - left),
                "height": pt(cap, bottom - top),
                "read_up": cap.value(case.cell, "own.border.diag_up.style") or "",
                "read_down": cap.value(case.cell, "own.border.diag_down.style") or "",
                **measure_line(cap, case, line),
            }
            for line in case.lines
        )
    return records


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``diagonals.tsv``."""
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
            records.extend(measure(cap, group))
    with open(OUTPUT, "w", encoding="utf-8") as out_file:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out_file.write(f"# {key}: {value}\n")
        out_file.write("# suite: diagonals\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out_file.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
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
