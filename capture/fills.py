"""Capture how Excel for Mac draws each pattern fill, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.screen run fills
    uv run --group capture python -m capture.screen read fills     # the saved captures again
    uv run --group capture python -m capture.fills build PATH      # the workbook only

Each of the 19 ``ST_PatternType`` values twice: black on white (``fgColor`` 000000, ``bgColor``
FFFFFF) in rows 4, 8, 12 and 16, and blue on yellow (0000FF on FFFF00) two rows below, in columns
C, E, G, I and K (12 characters wide, rows 24 pt) with narrow columns between. ``none`` is the
control: an unstyled cell. Row 20 holds ``darkUp`` in the adjacent cells C20, D20 and E20, for
the tile's anchoring and the gridlines between patterned cells. One window, at zoom 100 and 200;
the second marker is M22, so at zoom 200 the window must be about 1350 points wide to show it.

Measured per cell, 4 image pixels inside its edges: each pixel classed as foreground or
background (nearest color); the foreground's share; the repeating tile's period in points across
and down (sheet points at the zoom, as every length here); whether the pattern is drawn in whole
screen points (2 image pixels on a 2x display at any zoom; every pixel of a screen point the
same), and then the tile one character a screen point (``#`` foreground, ``.`` background, rows
joined by ``/``), anchored once at the sheet's origin and once at the cell's top left, each moved
to the nearest screen-point boundary (``snap_px``: the move at the sheet's origin, in image
pixels). When it is not drawn in whole screen points, the tile is given in image pixels and
``tile_unit`` says so. For the colored cells, the fill's extent against the cell's edges (does
a pattern cover its gridlines like a solid fill?) and the pattern and color ``display format``
reads; for the row 20 cells, how many gridline-gray pixels the middle row shows across the cell
and 2 pixels past each edge.

``read`` writes ``tests/captures/fills.tsv``: one row per zoom and cell.
"""

from __future__ import annotations

import math
import sys
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import PatternFill

PATTERNS = (
    "none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical",
    "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical",
    "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625",
)  # fmt: skip
SHEET_COLUMNS = ("C", "E", "G", "I", "K")
SETS = {"bw": ("000000", "FFFFFF", 0), "color": ("0000FF", "FFFF00", 2)}
"""Foreground, background and row offset of each color set."""
ANCHOR = ("C20", "D20", "E20")
MARKERS = ("B2", "M22")
ZOOMS = (100, 200)
INSET = 4
ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "fills"
OUTPUT = ROOT / "tests" / "captures" / "fills.tsv"
COLUMNS = (
    "zoom", "pattern", "colors", "cell", "fg_share", "unclassified", "period_x", "period_y",
    "whole_points", "snap_px", "tile_unit", "tile_sheet", "tile_cell", "fill_left", "fill_right",
    "fill_top", "fill_bottom", "gridline_pixels", "display_pattern", "display_fill",
)  # fmt: skip


def _slot(k: int, colors: str) -> str:
    return f"{SHEET_COLUMNS[k % 5]}{4 + 4 * (k // 5) + SETS[colors][2]}"


def cases() -> list[tuple[str, str, str]]:
    """Every measured cell as ``(pattern, colors, cell)``: both sets, then the anchor row."""
    found = [
        (pattern, colors, _slot(k, colors)) for k, pattern in enumerate(PATTERNS) for colors in SETS
    ]
    found.extend(("darkUp", "anchor", cell) for cell in ANCHOR)
    return found


def fill(ws: Any) -> list[tuple[str, str, str]]:
    """Lay the patterns out on a worksheet; every measured cell as ``cases`` gives it."""
    for column in SHEET_COLUMNS:
        ws.column_dimensions[column].width = 12
    for column in ("D", "F", "H", "J"):
        ws.column_dimensions[column].width = 2
    for block in range(4):
        for row, height in ((4, 24), (5, 7), (6, 24), (7, 7)):
            ws.row_dimensions[row + 4 * block].height = height
    ws.row_dimensions[20].height = 24
    found = cases()
    for pattern, colors, cell in found:
        if colors == "anchor":
            ws[cell].fill = PatternFill(pattern, fgColor="FF000000", bgColor="FFFFFFFF")
        elif pattern != "none":
            fg, bg, _ = SETS[colors]
            ws[cell].fill = PatternFill(pattern, fgColor="FF" + fg, bgColor="FF" + bg)
    return found


def build(path: Path) -> list[tuple[str, str, str]]:
    """Write the patterns to one workbook at ``path``; every measured cell."""
    wb = openpyxl.Workbook()
    wb.active.title = "Fills"
    placed = fill(wb.active)
    wb.save(path)
    return placed


def run() -> None:
    """Capture the window at both zooms, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    capture(
        fill,
        cells=[cell for _, _, cell in cases()],
        out=BUILD_DIR,
        markers=MARKERS,
        zooms=ZOOMS,
        name="fills",
    )
    write()


def _rgb(hex_color: str) -> tuple[int, int, int]:
    return int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)


def _bitmap(
    cap: Any, cell: str, fg: str, bg: str
) -> tuple[tuple[int, int], list[list[bool]], float]:
    """The cell's pixels as foreground flags, from its top left plus the inset.

    Returns the bitmap's origin in image pixels, the bitmap, and the share of pixels near neither
    color.
    """
    from capture.screen.ink import cells_box, nearest

    x0, y0, x1, y1 = cells_box(cap, cell, inset=INSET)
    palette = {"fg": _rgb(fg), "bg": _rgb(bg)}
    rows, unknown = [], 0
    for y in range(y0, y1):
        row = []
        for x in range(x0, x1):
            name = nearest(cap.image.pixel(x, y), palette, tolerance=90)
            unknown += name == "?"
            row.append(name == "fg")
        rows.append(row)
    return (x0, y0), rows, unknown / max(1, (x1 - x0) * (y1 - y0))


def _period(rows: list[list[bool]], axis: str, limit: int) -> int | None:
    """The smallest shift (1 % of pixels may disagree) that maps the bitmap onto itself."""
    height, width = len(rows), len(rows[0]) if rows else 0
    for p in range(1, limit + 1):
        if axis == "x":
            if p >= width:
                return None
            total = height * (width - p)
            pairs = ((row[x], row[x + p]) for row in rows for x in range(width - p))
        else:
            if p >= height:
                return None
            total = (height - p) * width
            pairs = ((rows[y][x], rows[y + p][x]) for y in range(height - p) for x in range(width))
        allowed, bad = 0.01 * total, 0
        for a, b in pairs:
            if a != b:
                bad += 1
                if bad > allowed:
                    break
        else:
            return p
    return None


def _tile_pixels(
    rows: list[list[bool]], origin: tuple[int, int], anchor: tuple[int, int], px: int, py: int
) -> str:
    """The tile in image pixels with its (0, 0) at ``anchor``."""
    ox, oy = origin
    ax, ay = anchor
    lines = []
    for ty in range(py):
        y = next(y for y in range(len(rows)) if (oy + y - ay) % py == ty)
        xs = [next(x for x in range(len(rows[0])) if (ox + x - ax) % px == tx) for tx in range(px)]
        lines.append("".join("#" if rows[y][x] else "." for x in xs))
    return "/".join(lines)


def _points(
    rows: list[list[bool]], origin: tuple[int, int], anchor: tuple[int, int], step: int
) -> dict[tuple[int, int], set[bool]]:
    """The values of the pixels of each screen point, points counted from ``anchor``."""
    ox, oy = origin
    ax, ay = anchor
    found: dict[tuple[int, int], set[bool]] = {}
    for y, row in enumerate(rows):
        ky = (oy + y - ay) // step
        for x, value in enumerate(row):
            found.setdefault(((ox + x - ax) // step, ky), set()).add(value)
    return found


def _tile_points(points: dict[tuple[int, int], set[bool]], qx: int, qy: int) -> str:
    """The tile one character a point, its (0, 0) at the points' origin."""
    tile = {}
    for (kx, ky), values in sorted(points.items()):
        tile.setdefault((kx % qx, ky % qy), next(iter(values)))
    return "/".join(
        "".join("#" if tile.get((tx, ty)) else "." for tx in range(qx)) for ty in range(qy)
    )


def _tiles(cap: Any, cell: str, fg: str, bg: str) -> dict[str, str]:
    from capture.screen.ink import pt

    origin, rows, unknown = _bitmap(cap, cell, fg, bg)
    count = sum(len(row) for row in rows)
    share = sum(map(sum, rows)) / max(1, count)
    found = {"fg_share": f"{round(share, 4):g}", "unclassified": f"{round(unknown, 4):g}"}
    if share in (0.0, 1.0):
        found["tile_sheet"] = found["tile_cell"] = "uniform"
        return found
    limit = round(12 * cap.scale)
    px, py = _period(rows, "x", limit), _period(rows, "y", limit)
    found["period_x"] = pt(cap, px) if px else ""
    found["period_y"] = pt(cap, py) if py else ""
    if not (px and py):
        return found
    sheet = (round(cap.grid.x(0)), round(cap.grid.y(0)))
    own = (round(cap.edge(cell, "left")), round(cap.edge(cell, "top")))
    screen = cap.scale * 100 / cap.zoom  # image pixels a screen point
    step = round(screen)
    qx, qy = math.lcm(px, step), math.lcm(py, step)
    phase = _phase(rows, origin, step) if abs(screen - step) < 0.05 else None
    found["whole_points"] = "yes" if phase else "no"
    if phase:
        found["tile_unit"] = "screen_pt"
        snapped = {name: _snap(at, phase, step) for name, at in (("sheet", sheet), ("cell", own))}
        found["snap_px"] = ",".join(
            str(s - a) for s, a in zip(snapped["sheet"], sheet, strict=True)
        )
        for name, at in snapped.items():
            found[f"tile_{name}"] = _tile_points(
                _points(rows, origin, at, step), qx // step, qy // step
            )
    else:
        found["tile_unit"] = "px"
        found["tile_sheet"] = _tile_pixels(rows, origin, sheet, px, py)
        found["tile_cell"] = _tile_pixels(rows, origin, own, px, py)
    return found


def _phase(rows: list[list[bool]], origin: tuple[int, int], step: int) -> tuple[int, int] | None:
    """Where screen points start (image x and y modulo ``step``), if every point is one color."""
    for dy in range(step):
        for dx in range(step):
            points = _points(rows, origin, (dx, dy), step)
            if all(len(values) == 1 for values in points.values()):
                return dx, dy
    return None


def _snap(at: tuple[int, int], phase: tuple[int, int], step: int) -> tuple[int, int]:
    """The point boundary nearest ``at`` on each axis."""
    snapped = []
    for a, p in zip(at, phase, strict=True):
        shift = (p - a) % step
        snapped.append(a + (shift - step if shift > step / 2 else shift))
    return snapped[0], snapped[1]


def _fill_box(cap: Any, cell: str) -> dict[str, str]:
    """The blue-and-yellow fill's extent against the cell's edges, in points."""
    from capture.screen.ink import nearest, pt

    palette = {"blue": (0, 0, 255), "yellow": (255, 255, 0), "white": (255, 255, 255)}
    palette["black"] = (0, 0, 0)
    if cap.gridline is not None:
        palette["grid"] = cap.gridline
    left, right = round(cap.edge(cell, "left")), round(cap.edge(cell, "right"))
    top, bottom = round(cap.edge(cell, "top")), round(cap.edge(cell, "bottom"))
    cx, cy = (left + right) // 2, (top + bottom) // 2
    pad = 6
    colored = ("blue", "yellow")
    xs = [
        x
        for x in range(left - pad, right + pad)
        if nearest(cap.image.pixel(x, cy), palette) in colored
    ]
    ys = [
        y
        for y in range(top - pad, bottom + pad)
        if nearest(cap.image.pixel(cx, y), palette) in colored
    ]
    if not xs or not ys:
        return {}
    return {
        "fill_left": pt(cap, xs[0] - left),
        "fill_right": pt(cap, xs[-1] + 1 - right),
        "fill_top": pt(cap, ys[0] - top),
        "fill_bottom": pt(cap, ys[-1] + 1 - bottom),
    }


def _gridline_pixels(cap: Any, cell: str) -> str:
    from capture.screen.ink import nearest

    left, right = round(cap.edge(cell, "left")), round(cap.edge(cell, "right"))
    y = round(cap.edge(cell, "cy"))
    if cap.gridline is None:
        return ""  # the platform's gray is not recorded (window.NO_GRIDLINE)
    grid = {"grid": cap.gridline}
    count = sum(
        nearest(cap.image.pixel(x, y), grid, 8) == "grid" for x in range(left - 2, right + 2)
    )
    return str(count)


def measure(cap: Any) -> list[dict[str, str]]:
    """Every cell's record at one zoom."""
    records = []
    for pattern, colors, cell in cases():
        fg, bg = ("000000", "FFFFFF") if colors == "anchor" else SETS[colors][:2]
        record = {"zoom": str(cap.zoom), "pattern": pattern, "colors": colors, "cell": cell}
        record.update(_tiles(cap, cell, fg, bg))
        if colors == "color":
            record.update(_fill_box(cap, cell))
            record["display_pattern"] = cap.value(cell, "display.pattern") or ""
            record["display_fill"] = cap.value(cell, "display.fill") or ""
        if colors == "anchor":
            record["gridline_pixels"] = _gridline_pixels(cap, cell)
        records.append(record)
    return records


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``fills.tsv``."""
    from capture.screen.run import header, load

    first = header(BUILD_DIR)
    records = [record for cap in load(BUILD_DIR) for record in measure(cap)]
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: fills\n" + "\t".join(COLUMNS) + "\n")
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
