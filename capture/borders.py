"""Capture how Excel for Mac draws each border line style, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.screen run borders
    uv run --group capture python -m capture.screen read borders     # the saved captures again
    uv run --group capture python -m capture.borders build PATH      # the workbook only

Each of the 13 ``ST_BorderStyle`` line styles is drawn black on all four sides of one cell (20
characters wide, 40 pt high), the cells apart, in columns C, E, G and I by rows 4, 6, 8 and 10;
a 14th cell with no border is the gridline control. One window, at zoom 100 and 200; the second
marker is K12, which at zoom 200 shows in a window about 1650 points wide. A narrower window
needs only the first marker at zoom 100 and 200, where Excel draws at the nominal scale
(`CAPTURING.md`, "One marker, on a small screen").

Measured on each edge, 12 image pixels clear of the corners: every device row (across the line)
that carries ink, as its offset from the cell's boundary in points (0 is the first pixel after
the boundary, negative before it, both down or to the right of the sheet), so a line's width,
double's three rows and where each sits against the gridline; along each such row, the on/off
run lengths of one period in points (``on 3 off 1``, ``solid``) and where its first ink starts
from the cell's left or top edge, so that dash patterns and any stagger between device rows can
be read; the darkest ink (hair may be gray), and what the gaps show (white, or the gridline's
gray). The control cell's rows are its gridline's, found by the gridline's gray instead of ink.

``read`` writes ``tests/captures/borders.tsv``: one row per zoom, style, edge and inked device
row, lengths in sheet points at the row's zoom (a screen point is 1 at zoom 100, 0.5 at 200).
"""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Border, Side

STYLES = (
    "none", "thin", "medium", "thick", "double", "hair", "dotted", "dashed", "dashDot",
    "dashDotDot", "mediumDashed", "mediumDashDot", "mediumDashDotDot", "slantDashDot",
)  # fmt: skip
SHEET_COLUMNS = ("C", "E", "G", "I")
CELLS = {style: f"{SHEET_COLUMNS[k % 4]}{4 + 2 * (k // 4)}" for k, style in enumerate(STYLES)}
"""Each style's cell."""
MARKERS = ("B2", "K12")
ZOOMS = (100, 200)
EDGES = ("bottom", "right", "top", "left")
CLEAR = 12
"""Image pixels kept clear of each corner, so that the other sides' lines are not scanned."""
HALF = 10
"""Image pixels scanned on each side of a boundary."""
ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "borders"
OUTPUT = ROOT / "tests" / "captures" / "borders.tsv"
COLUMNS = ("zoom", "style", "cell", "edge", "offset", "pattern", "first_on", "ink", "gaps")


def cases() -> dict[str, str]:
    """Every style with its cell, ``none`` (the gridline control) first."""
    return dict(CELLS)


def fill(ws: Any) -> list[tuple[str, str]]:
    """Lay the styles out on a worksheet; each style's cell address with the style."""
    for column in SHEET_COLUMNS:
        ws.column_dimensions[column].width = 20
    for column in ("D", "F", "H"):
        ws.column_dimensions[column].width = 3
    for row in (4, 6, 8, 10):
        ws.row_dimensions[row].height = 40
        ws.row_dimensions[row + 1].height = 10
    for style, cell in CELLS.items():
        if style != "none":
            side = Side(style=style, color="FF000000")
            ws[cell].border = Border(left=side, right=side, top=side, bottom=side)
    return [(cell, style) for style, cell in CELLS.items()]


def build(path: Path) -> list[tuple[str, str]]:
    """Write the styles to one workbook at ``path``; each style's cell address with the style."""
    wb = openpyxl.Workbook()
    wb.active.title = "Borders"
    placed = fill(wb.active)
    wb.save(path)
    return placed


def run() -> None:
    """Capture the window at both zooms, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    capture(
        fill,
        cells=list(CELLS.values()),
        out=BUILD_DIR,
        markers=MARKERS,
        zooms=ZOOMS,
        name="borders",
    )
    write()


def _is_ink(rgb: tuple[int, int, int]) -> bool:
    from capture.screen.run import luma

    return luma(rgb) < 180


def _is_grid(rgb: tuple[int, int, int], grid: tuple[int, int, int]) -> bool:
    """Within 10 per channel of the platform's gridline gray (``Capture.gridline``)."""
    return all(abs(c - g) <= 10 for c, g in zip(rgb, grid, strict=True))


def _line(cap: Any, along: str, fixed: int, a0: int, a1: int) -> list[tuple[int, int, int]]:
    if along == "x":
        return [cap.image.pixel(a, fixed) for a in range(a0, a1)]
    return [cap.image.pixel(fixed, a) for a in range(a0, a1)]


def _measure_edge(cap: Any, cell: str, edge: str, control: bool) -> list[dict[str, str]]:
    """One record per device row near ``edge`` that carries ink (or, for the control, gridline)."""
    from capture.screen.ink import nearest, pattern, pt

    left, right = round(cap.edge(cell, "left")), round(cap.edge(cell, "right"))
    top, bottom = round(cap.edge(cell, "top")), round(cap.edge(cell, "bottom"))
    along = "x" if edge in ("top", "bottom") else "y"
    boundary = {"top": top, "bottom": bottom, "left": left, "right": right}[edge]
    start, end = (left, right) if along == "x" else (top, bottom)
    a0, a1 = start + CLEAR, end - CLEAR
    grid = cap.gridline
    shades = (
        {"white": (255, 255, 255)} if grid is None else {"white": (255, 255, 255), "grid": grid}
    )

    def test(rgb: tuple[int, int, int]) -> bool:
        return _is_grid(rgb, grid) if control and grid is not None else _is_ink(rgb)

    found = []
    for offset in range(-HALF, HALF):
        pixels = _line(cap, along, boundary + offset, a0, a1)
        flags = [test(p) for p in pixels]
        if not any(flags):
            continue
        marked = Counter(p for p, f in zip(pixels, flags, strict=True) if f)
        gaps = Counter(nearest(p, shades, 20) for p, f in zip(pixels, flags, strict=True) if not f)
        darkest = min(marked, key=sum)
        found.append(
            {
                "offset": pt(cap, offset),
                "pattern": pattern(flags, cap.scale),
                "first_on": pt(cap, CLEAR + flags.index(True)),
                "ink": "".join(f"{c:02X}" for c in darkest),
                "gaps": " ".join(name for name, _ in gaps.most_common()),
            }
        )
    return found


def measure(cap: Any) -> list[dict[str, str]]:
    """Every style's records at one zoom; the gridline control only where its gray is known."""
    from capture.screen.window import NO_GRIDLINE

    styles = dict(CELLS)
    if cap.gridline is None:
        del styles["none"]
        print(f"borders: the gridline control is skipped at zoom {cap.zoom}: {NO_GRIDLINE}")
    return [
        {"zoom": str(cap.zoom), "style": style, "cell": cell, "edge": edge, **found}
        for style, cell in styles.items()
        for edge in EDGES
        for found in _measure_edge(cap, cell, edge, style == "none")
    ]


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``borders.tsv``."""
    from capture.screen.run import header, load

    first = header(BUILD_DIR)
    records = [record for cap in load(BUILD_DIR) for record in measure(cap)]
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: borders\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cells in {argv[1]}")
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
