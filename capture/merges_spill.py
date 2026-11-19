"""Capture how Excel for Mac draws borders around merges and where text spills, from the screen.

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.merges_spill run
    uv run --group capture python -m capture.merges_spill read        # the saved captures again
    uv run --group capture python -m capture.merges_spill build PATH  # the workbook only

ARCHITECTURE.md §4.12 ("Merges" and "Spilled text removes every vertical border it crosses"),
with §4.1 and §3.6 for where text spills and which gridlines it hides. One workbook, one window,
at zoom 100 and 200, in the dark appearance as the other suites. Columns C to J are 14
characters wide (84 pt in sheetpeek's layout), the merge rows 30 pt high and the spill rows
24 pt; spacer rows keep the cases' sides from meeting. Borders are red ``FF0000``, blue
``0000FF`` and green ``00B050`` in thin, medium, thick and double (styles whose geometry the
suite ``borders`` captured), never black, so that black text crossing a boundary is not taken
for a border. The cases:

- ``mixed`` (merge C4:E5): every edge cell carries outer sides of its own, each unlike its
  neighbors' (style and color), and inner sides on the boundaries inside the merge. Checks
  "the merge's outline is the outer sides of its edge cells, each on its own segment" and "sides
  inside the merge are not drawn".
- ``anchor_only`` (merge C8:E9): only the anchor has sides, all four thick red. Checks "the
  anchor's [sides] only on its own [segment]": its top and left cover C8's segments alone, and
  its right and bottom, inside the merge, are not drawn.
- ``neighbor`` (merge H4:I5): each of the eight outline segments has a neighbor's facing side,
  in blue, that wins on four of them and loses on the other four by style. Checks "a neighbor's
  side competes with each segment by the rule above".
- ``merge_diagonal`` (merge H8:J9): the anchor's diagonals, up and down, medium red. Checks "a
  merge's diagonals are its anchor's, across the whole merge" (and not across the anchor's own
  cell).
- ``covered_diagonal`` (merge C11:D12): the covered D12 has both diagonals, thin green; the
  anchor none. Checks "a covered cell's diagonal is not drawn", in D12's box or the merge's.
- ``spill_borders`` (row 14): left-aligned text in C14 spills over D14 to F14 and stops at G14's
  value; sides on every vertical boundary it crosses, horizontal sides on the cells it crosses,
  and a left side on G14. Checks "spilled text removes every vertical border it crosses, over the
  whole row height, on either cell", "the horizontal sides of the cells it crosses stay" and
  "where the spill stops at a value, that cell's left side is drawn".
- ``spill_fills`` (row 16): C16 (gold fill) spills over D16 (sky fill, thick red left and right)
  into E16, whose right side the text does not reach. Checks the same removal over fills, what
  shows where a removed side was (the fill or white, §3.6's crossed boundaries), and that a side
  the text does not reach stays (a spill takes whole columns, as far as the text needs, §4.1).
- ``spill_right`` (row 18): right-aligned text in F18 spills left over E18 and D18 and stops at
  C18's value, whose right side is drawn: the mirror of ``spill_borders``.
- ``spill_center`` (row 20): centered text in E20 takes one column each way; the sides it
  crosses go, the outer ones stay.
- ``spill_end`` (row 22): text that ends inside D22 removes C22's right side and keeps D22's.

Measured, for every edge in ``EDGES`` (a grid cell's top, bottom, left or right, also a covered
cell's), from the boundary's image row (or column) across ``HALF`` image pixels each way, along
the edge less ``CLEAR`` pixels at each end: each pixel named by the nearest color of ``NAMES``
(white, the gridline gray, black, the border colors, the fills); the border color covering most
of the length (``color``) and the share of the length it covers (``coverage``); the device rows
(offsets in points from the boundary, as the suite ``borders`` writes them) in which that color
covers half the length or more (``rows``) and the on/off pattern of the first (``pattern``); and
what the boundary row itself shows most (``shows``: a fill, white, or the gridline). For every
diagonal in ``DIAGONALS``, the share of four points along the box's corner-to-corner line (35, 45,
55 and 65 %) that have the diagonal's color within ``REACH`` image pixels (``coverage``). The
cells' boxes come from reference cells (row 1 for the columns, column A for the rows), so that
no merged cell's geometry has to be read: 19 cells.

``read`` writes ``tests/captures/merges_spill.tsv``, one row per zoom and edge or diagonal.
"""

from __future__ import annotations

import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

RED, BLUE, GREEN = "FF0000", "0000FF", "00B050"
GOLD, SKY = "FFE699", "9BC2E6"
INK = {"red": RED, "blue": BLUE, "green": GREEN}
"""The border colors, by the names the records use."""
FILLS = {"gold": GOLD, "sky": SKY}
NAMES = {"white": "FFFFFF", "grid": "D4D4D4", "black": "000000", **INK, **FILLS}
"""Every color a pixel is named by (the nearest); the gridline's gray is the run's own."""

COLUMN_WIDTH = 14
"""Characters, for columns C to J."""
HEIGHTS = {
    3: 20, 4: 30, 5: 30, 6: 20, 7: 8, 8: 30, 9: 30, 10: 8, 11: 30, 12: 30, 13: 8,
    14: 24, 15: 8, 16: 24, 17: 8, 18: 24, 19: 8, 20: 24, 21: 8, 22: 24,
}  # fmt: skip
MARKERS = ("B2", "K23")
ZOOMS = (100, 200)
REF_ROW, REF_COLUMN = 1, "A"
CLEAR = 12
"""Image pixels kept clear of each end of an edge, so that the sides meeting it are not read."""
HALF = 10
"""Image pixels read on each side of a boundary."""
REACH = 3
"""Image pixels around each point of a diagonal searched for its color."""
SPOTS = (0.35, 0.45, 0.55, 0.65)
"""Where along a diagonal its color is looked for: clear of the corners, where the sides meet, and
of the middle, where the other diagonal crosses it."""
ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "merges_spill"
OUTPUT = ROOT / "tests" / "captures" / "merges_spill.tsv"
COLUMNS = (
    "zoom", "kind", "case", "cell", "edge", "statement", "color", "coverage", "rows", "pattern",
    "shows",
)  # fmt: skip

LONG = "spilled text runs over the borders of every cell it crosses until a value stops it here"
"""360 pt of Calibri 11 (``xlformat.text_width``): from C over D, E and F, stopped at G's value;
right-aligned from F, over E and D, stopped at C's."""
MID = "spilled text over a fill and over a border here"
"""196 pt: from C over D into E, whose right side it does not reach."""
CENTER = "centered text over both sides"
"""130 pt: one column each way from E."""
SHORT = "a text one column and a half"
"""126 pt: from C into D."""

CASES = {
    "mixed": "merge C4:E5, mixed outer sides on every edge cell and sides inside",
    "anchor_only": "merge C8:E9, only the anchor has sides, all four thick red",
    "neighbor": "merge H4:I5 with a neighbor's side against each outline segment",
    "merge_diagonal": "merge H8:J9, the anchor's diagonals up and down, medium red",
    "covered_diagonal": "merge C11:D12, the covered D12's diagonals, thin green",
    "spill_borders": "C14 spills over D14:F14 to G14's value, sides on every boundary",
    "spill_fills": "C16 (gold) spills over D16 (sky, thick red sides) into E16",
    "spill_right": "right-aligned F18 spills over E18:D18 to C18's value",
    "spill_center": "centered E20 takes D20 and F20",
    "spill_end": "C22 spills into D22 only",
}
MERGES = ("C4:E5", "C8:E9", "H4:I5", "H8:J9", "C11:D12")


@dataclass(frozen=True)
class Placed:
    """A cell's contents: its sides (``l``, ``r``, ``t``, ``b``: style and color), value, fill."""

    sides: tuple[tuple[str, str, str], ...] = ()
    value: str | None = None
    fill: str | None = None
    align: str | None = None
    diagonal: tuple[str, str] | None = None
    """Style and color of both diagonals, up and down."""


def _p(sides: str = "", **kw: Any) -> Placed:
    """``"l thick red, r thin blue"`` as a :class:`Placed`."""
    found = []
    for part in filter(None, (s.strip() for s in sides.split(","))):
        edge, style, color = part.split()
        found.append((edge, style, color))
    return Placed(tuple(found), **kw)


CELLS: dict[str, Placed] = {
    # mixed: outer sides each unlike their neighbors, and inner sides
    "C4": _p("t thick red, l thick red, r thin blue, b thin blue"),
    "D4": _p("t medium green, l double red, b thick blue"),
    "E4": _p("t thin blue, r medium red, l thick green, b medium green"),
    "C5": _p("l double blue, b thick green, t thick red, r medium blue"),
    "D5": _p("b medium red, t thin green, r thick red"),
    "E5": _p("b thin red, r thick blue"),
    # neighbor: red outline segments, blue neighbors winning on half of them
    "H4": _p("t thin red, l medium red"),
    "I4": _p("t thick red, r thin red"),
    "H5": _p("l thin red, b thin red"),
    "I5": _p("r medium red, b medium red"),
    "H3": _p("b thick blue"),
    "I3": _p("b thin blue"),
    "G4": _p("r thin blue"),
    "G5": _p("r thick blue"),
    "J4": _p("l medium blue"),
    "J5": _p("l thin blue"),
    "H6": _p("t thick blue"),
    "I6": _p("t thin blue"),
    # anchor_only
    "C8": _p("l thick red, r thick red, t thick red, b thick red"),
    # merge_diagonal and covered_diagonal
    "H8": _p(diagonal=("medium", "red")),
    "D12": _p(diagonal=("thin", "green")),
    # spill_borders
    "C14": _p("r thin blue", value=LONG),
    "D14": _p("l medium red, r thick blue, t medium green, b thick red"),
    "E14": _p("r thin green, b medium blue"),
    "F14": _p("t thin red"),
    "G14": _p("l medium red", value="stop"),
    # spill_fills
    "C16": _p(value=MID, fill="gold"),
    "D16": _p("l thick red, r thick red", fill="sky"),
    "E16": _p("r thin blue"),
    # spill_right
    "C18": _p("r medium green", value="stop"),
    "D18": _p("r thin blue"),
    "E18": _p("l thick red"),
    "F18": _p("l thin red, r thin blue", value=LONG, align="right"),
    # spill_center
    "D20": _p("l thin red, r medium blue, t thin blue"),
    "E20": _p("r thick red", value=CENTER, align="center"),
    "F20": _p("r thin green, b medium red"),
    # spill_end
    "C22": _p("r thin blue", value=SHORT),
    "D22": _p("r thick red"),
}

OUTLINE = "outline: the outer sides of the edge cells, each on its own segment"
INNER = "sides inside the merge are not drawn"
ANCHOR = "the anchor's sides only on its own segments"
NEIGHBOR = "a neighbor's side competes with each segment"
CROSSED = "spilled text removes every vertical border it crosses, over the whole row height"
STAYS = "the horizontal sides of the cells spilled text crosses stay"
BLOCKER = "where the spill stops at a value, that cell's facing side is drawn"
REACHED = "a side past the spill's last column stays"
MERGE_DIAGONAL = "a merge's diagonals are its anchor's, across the whole merge"
COVERED_DIAGONAL = "a covered cell's diagonal is not drawn"


def _edges(case: str, statement: str, spec: str) -> list[tuple[str, str, str, str]]:
    """``"C4 top, D4 top"`` as ``(case, cell, edge, statement)`` records."""
    found = []
    for part in spec.split(", "):
        cell, edge = part.split()
        found.append((case, cell, edge, statement))
    return found


EDGES: list[tuple[str, str, str, str]] = [
    *_edges("mixed", OUTLINE, "C4 top, D4 top, E4 top, C5 bottom, D5 bottom, E5 bottom"),
    *_edges("mixed", OUTLINE, "C4 left, C5 left, E4 right, E5 right"),
    *_edges("mixed", INNER, "C4 right, D4 right, C5 right, D5 right, C4 bottom, D4 bottom"),
    *_edges("mixed", INNER, "E4 bottom"),
    *_edges("anchor_only", ANCHOR, "C8 top, D8 top, E8 top, C8 left, C9 left, E8 right"),
    *_edges("anchor_only", ANCHOR, "E9 right, C9 bottom, D9 bottom, E9 bottom"),
    *_edges("anchor_only", INNER, "C8 right, C8 bottom"),
    *_edges("neighbor", NEIGHBOR, "H4 top, I4 top, H4 left, H5 left, I4 right, I5 right"),
    *_edges("neighbor", NEIGHBOR, "H5 bottom, I5 bottom"),
    *_edges("spill_borders", CROSSED, "C14 right, D14 right, E14 right"),
    *_edges("spill_borders", BLOCKER, "G14 left"),
    *_edges("spill_borders", STAYS, "D14 top, D14 bottom, E14 bottom, F14 top"),
    *_edges("spill_fills", CROSSED, "C16 right, D16 right"),
    *_edges("spill_fills", REACHED, "E16 right"),
    *_edges("spill_fills", STAYS, "D16 top, C16 bottom"),
    *_edges("spill_right", CROSSED, "D18 right, E18 right"),
    *_edges("spill_right", BLOCKER, "C18 right"),
    *_edges("spill_right", REACHED, "F18 right"),
    *_edges("spill_center", CROSSED, "D20 right, E20 right"),
    *_edges("spill_center", REACHED, "D20 left, F20 right"),
    *_edges("spill_center", STAYS, "D20 top, F20 bottom"),
    *_edges("spill_end", CROSSED, "C22 right"),
    *_edges("spill_end", REACHED, "D22 right"),
]
"""Every edge measured: its case, its grid cell (a covered one too), the edge, the statement."""

DIAGONALS: list[tuple[str, str, str, str, str]] = [
    (case, box, direction, color, statement)
    for case, box, color, statement in (
        ("merge_diagonal", "H8:J9", "red", MERGE_DIAGONAL),
        ("merge_diagonal", "H8:H8", "red", MERGE_DIAGONAL),
        ("covered_diagonal", "D12:D12", "green", COVERED_DIAGONAL),
        ("covered_diagonal", "C11:D12", "green", COVERED_DIAGONAL),
    )
    for direction in ("up", "down")
]
"""Every diagonal looked for: its case, the box it would run across, up or down, its color."""

EXPECTED: dict[tuple[str, str], str] = {
    ("C4", "top"): "thick FF0000", ("D4", "top"): "medium 00B050", ("E4", "top"): "thin 0000FF",
    ("C5", "bottom"): "thick 00B050", ("D5", "bottom"): "medium FF0000",
    ("E5", "bottom"): "thin FF0000", ("C4", "left"): "thick FF0000",
    ("C5", "left"): "double 0000FF", ("E4", "right"): "medium FF0000",
    ("E5", "right"): "thick 0000FF", ("C4", "right"): "none", ("D4", "right"): "none",
    ("C5", "right"): "none", ("D5", "right"): "none", ("C4", "bottom"): "none",
    ("D4", "bottom"): "none", ("E4", "bottom"): "none", ("C8", "top"): "thick FF0000",
    ("D8", "top"): "none", ("E8", "top"): "none", ("C8", "left"): "thick FF0000",
    ("C9", "left"): "none", ("E8", "right"): "none", ("E9", "right"): "none",
    ("C9", "bottom"): "none", ("D9", "bottom"): "none", ("E9", "bottom"): "none",
    ("C8", "right"): "none", ("C8", "bottom"): "none", ("H4", "top"): "thick 0000FF",
    ("I4", "top"): "thick FF0000", ("H4", "left"): "medium FF0000",
    ("H5", "left"): "thick 0000FF", ("I4", "right"): "medium 0000FF",
    ("I5", "right"): "medium FF0000", ("H5", "bottom"): "thick 0000FF",
    ("I5", "bottom"): "medium FF0000", ("C14", "right"): "none", ("D14", "right"): "none",
    ("E14", "right"): "none", ("G14", "left"): "medium FF0000",
    ("D14", "top"): "medium 00B050", ("D14", "bottom"): "thick FF0000",
    ("E14", "bottom"): "medium 0000FF", ("F14", "top"): "thin FF0000",
    ("C16", "right"): "none", ("D16", "right"): "none", ("E16", "right"): "thin 0000FF",
    ("D16", "top"): "none", ("C16", "bottom"): "none", ("D18", "right"): "none",
    ("E18", "right"): "none", ("C18", "right"): "medium 00B050",
    ("F18", "right"): "thin 0000FF", ("D20", "right"): "none", ("E20", "right"): "none",
    ("D20", "left"): "thin FF0000", ("F20", "right"): "thin 00B050",
    ("D20", "top"): "thin 0000FF", ("F20", "bottom"): "medium FF0000",
    ("C22", "right"): "none", ("D22", "right"): "thick FF0000",
    ("H8:J9", "up"): "drawn", ("H8:J9", "down"): "drawn", ("H8:H8", "up"): "none",
    ("H8:H8", "down"): "none", ("D12:D12", "up"): "none", ("D12:D12", "down"): "none",
    ("C11:D12", "up"): "none", ("C11:D12", "down"): "none",
}  # fmt: skip
"""What sheetpeek draws on each edge (``style RRGGBB`` or ``none``) and each diagonal's box
(``drawn`` or ``none``), from :func:`sheetpeek_drawn`; the capture shows where Excel differs."""


def references() -> list[str]:
    """The cells whose geometry is read: row 1 for each column, column A for each row."""
    cells = [cell for _, cell, _, _ in EDGES]
    for _, box, *_ in DIAGONALS:
        cells += box.split(":")
    columns = sorted({_split(c)[0] for c in cells}, key=column_index_from_string)
    rows = sorted({_split(c)[1] for c in cells})
    return [f"{c}{REF_ROW}" for c in columns] + [f"{REF_COLUMN}{r}" for r in rows]


def _split(cell: str) -> tuple[str, int]:
    letters, row = coordinate_from_string(cell)
    return letters, row


def fill(ws: Any) -> list[tuple[str, str]]:
    """Lay the cases out on a worksheet; each case with its cells."""
    for k in range(3, 11):
        ws.column_dimensions[get_column_letter(k)].width = COLUMN_WIDTH
    for row, height in HEIGHTS.items():
        ws.row_dimensions[row].height = height
    # Merged first: openpyxl copies an anchor's border to the merge's edges when it merges.
    for merge in MERGES:
        ws.merge_cells(merge)
    for address, placed in CELLS.items():
        cell = ws[address]
        if placed.value is not None:
            cell.value = placed.value
        if placed.fill is not None:
            cell.fill = PatternFill("solid", fgColor="FF" + FILLS[placed.fill])
        if placed.align is not None:
            cell.alignment = Alignment(horizontal=placed.align)
        sides = {
            {"l": "left", "r": "right", "t": "top", "b": "bottom"}[e]: Side(s, color="FF" + INK[c])
            for e, s, c in placed.sides
        }
        if placed.diagonal is not None:
            style, color = placed.diagonal
            sides["diagonal"] = Side(style, color="FF" + INK[color])
        if sides:
            diagonal = placed.diagonal is not None
            cell.border = Border(**sides, diagonalUp=diagonal, diagonalDown=diagonal)
    return [(case, CASES[case]) for case in CASES]


def build(path: Path) -> list[tuple[str, str]]:
    """Write the cases to one workbook at ``path``; each case with its description."""
    wb = openpyxl.Workbook()
    wb.active.title = "Merges"
    placed = fill(wb.active)
    wb.save(path)
    return placed


def run() -> None:
    """Capture the window at both zooms, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    capture(
        fill,
        cells=references(),
        out=BUILD_DIR,
        markers=MARKERS,
        zooms=ZOOMS,
        name="merges_spill",
    )
    write()


def sheetpeek_drawn(path: Path) -> dict[tuple[str, str], str]:
    """What sheetpeek draws on each edge and diagonal of :data:`EDGES` and :data:`DIAGONALS`."""
    import sheetpeek

    rendered = sheetpeek.read(path).layout("excel")
    borders = rendered.borders
    assert borders is not None
    rows = {r.index: i for i, r in enumerate(rendered.rows)}
    columns = {c.index: p for p, c in enumerate(rendered.columns)}

    def at(cell: str) -> tuple[int, int]:
        letters, row = _split(cell)
        return rows[row], columns[column_index_from_string(letters)]

    found: dict[tuple[str, str], str] = {}
    for _, cell, edge, _ in EDGES:
        i, p = at(cell)
        if edge in ("top", "bottom"):
            side = borders.horizontal[i + (edge == "bottom")][p]
        else:
            side = borders.vertical[i][p + (edge == "right")]
        found[cell, edge] = "none" if side is None else f"{side.style} {side.drawn_color}"
    for _, box, direction, _, _ in DIAGONALS:
        first, last = box.split(":")
        (top, left), (bottom, right) = at(first), at(last)
        hits = [
            d
            for d in borders.diagonals
            if (d.top, d.left, d.bottom, d.right) == (top, left, bottom, right)
            and (d.up if direction == "up" else d.down)
        ]
        found[box, direction] = "drawn" if hits else "none"
    return found


# -- measuring -----------------------------------------------------------------------------------


def _box(cap: Any, cell: str) -> tuple[float, float, float, float]:
    """A grid cell's left, top, right and bottom in image pixels, from the reference cells."""
    letters, row = _split(cell)
    column, line = f"{letters}{REF_ROW}", f"{REF_COLUMN}{row}"
    return (
        cap.edge(column, "left"),
        cap.edge(line, "top"),
        cap.edge(column, "right"),
        cap.edge(line, "bottom"),
    )


def _palette(cap: Any) -> dict[str, tuple[int, int, int]]:
    found = {name: _rgb(hex6) for name, hex6 in NAMES.items()}
    if cap.gridline is not None:
        found["grid"] = cap.gridline
    return found


def _rgb(hex6: str) -> tuple[int, int, int]:
    return int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)


def _name(rgb: tuple[int, int, int], palette: dict[str, tuple[int, int, int]]) -> str:
    """The nearest named color: every pixel gets one."""
    return min(
        palette, key=lambda n: sum((a - b) ** 2 for a, b in zip(rgb, palette[n], strict=True))
    )


def measure_edge(cap: Any, case: str, cell: str, edge: str, statement: str) -> dict[str, str]:
    """One edge's record: the border color drawn on it, where, how much, and what else shows."""
    from capture.screen.ink import pattern, pt

    palette = _palette(cap)
    left, top, right, bottom = (round(v) for v in _box(cap, cell))
    along_x = edge in ("top", "bottom")
    boundary = {"top": top, "bottom": bottom, "left": left, "right": right}[edge]
    a0, a1 = (left, right) if along_x else (top, bottom)
    a0, a1 = a0 + CLEAR, a1 - CLEAR

    def pixel(a: int, offset: int) -> tuple[int, int, int]:
        x, y = (a, boundary + offset) if along_x else (boundary + offset, a)
        return cap.image.pixel(x, y)

    named = {
        offset: [_name(pixel(a, offset), palette) for a in range(a0, a1)]
        for offset in range(-HALF, HALF)
    }
    length = a1 - a0
    record = {
        "zoom": str(cap.zoom),
        "kind": "edge",
        "case": case,
        "cell": cell,
        "edge": edge,
        "statement": statement,
        "shows": Counter(named[0]).most_common(1)[0][0],
    }
    coverage = {
        color: sum(any(named[o][k] == color for o in named) for k in range(length)) / length
        for color in INK
    }
    color = max(INK, key=lambda c: coverage[c])
    if coverage[color] == 0:
        return record | {"coverage": "0"}
    rows = [o for o in sorted(named) if named[o].count(color) >= length / 2]
    record |= {"color": color, "coverage": f"{coverage[color]:.3f}"}
    if rows:
        record["rows"] = " ".join(pt(cap, o) for o in rows)
        record["pattern"] = pattern([n == color for n in named[rows[0]]], cap.scale)
    return record


def measure_diagonal(
    cap: Any, case: str, box: str, direction: str, color: str, statement: str
) -> dict[str, str]:
    """One diagonal's record: the share of :data:`SPOTS` along it that show its color."""
    palette = _palette(cap)
    first, last = box.split(":")
    left, top, _, _ = _box(cap, first)
    _, _, right, bottom = _box(cap, last)
    hits = 0
    for t in SPOTS:
        x = left + t * (right - left)
        y = bottom - t * (bottom - top) if direction == "up" else top + t * (bottom - top)
        cx, cy = round(x), round(y)
        if any(
            _name(cap.image.pixel(cx + dx, cy + dy), palette) == color
            for dy in range(-REACH, REACH + 1)
            for dx in range(-REACH, REACH + 1)
        ):
            hits += 1
    return {
        "zoom": str(cap.zoom),
        "kind": "diagonal",
        "case": case,
        "cell": box,
        "edge": direction,
        "statement": statement,
        "color": color,
        "coverage": f"{hits / len(SPOTS):g}",
    }


def measure(cap: Any) -> list[dict[str, str]]:
    """Every edge's and diagonal's record at one zoom."""
    return [measure_edge(cap, *edge) for edge in EDGES] + [
        measure_diagonal(cap, *diagonal) for diagonal in DIAGONALS
    ]


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``merges_spill.tsv``."""
    from capture.screen.run import header, load

    first = header(BUILD_DIR)
    records = [record for cap in load(BUILD_DIR) for record in measure(cap)]
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: merges_spill\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


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
