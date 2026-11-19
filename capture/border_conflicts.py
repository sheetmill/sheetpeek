"""Capture which of two meeting border sides Excel for Mac draws, from the screen (route C).

    export EXCEL_EDITION="Microsoft 365"
    uv run --group capture python -m capture.border_conflicts run start_end
    uv run --group capture python -m capture.border_conflicts run
    uv run --group capture python -m capture.border_conflicts read         # saved captures again
    uv run --group capture python -m capture.border_conflicts build PATH   # the workbooks only

ARCHITECTURE.md §4.12, "Conflicts" (the style order, the darker color by Rec. 601 luma, equal
luma to the larger ``RRGGBB``, colors compared after resolution), "Hidden rows and columns" (the
visible neighbors' facing sides compete on one edge) and "Borders" (``start`` and ``end`` read
as ``left`` and ``right``). Each group is its own workbook and window, at zoom 100, in the dark
appearance as the other suites (the sheet stays white).

Every case is one edge where two cells meet, each giving that edge a side of its own. The cells
of a group form a dense grid in which every cell has four sides, so each interior edge is a case
of its own: a vertical edge is the left cell's right side against the right cell's left side, a
horizontal edge the upper cell's bottom side against the lower cell's top side. The grid's outer
top and left edges have no neighbor and are the references: each line style drawn alone, in the
same window, to which a shared edge's line is matched. Groups:

- ``styles``: a 14 x 13 grid (C4:P16, columns 10 characters wide, rows 56 pt), every side
  black: all 156 ordered pairs of the 13 styles on vertical edges and again on horizontal ones
  (78 pairs x both orders x both orientations); the top edges of C4:O4 and the left edges of
  C4:C16 are the 13 styles alone.
- ``colors``: a 7 x 6 grid (C4:I9, rows 30 pt) of thin sides in pairs of colors, in both orders
  and orientations: pairs that order differently by Rec. 601 luma than by Rec. 709 luma, by the
  largest channel, by the mean of the channels and by CIE L*; four pairs of exactly equal Rec. 601
  luma; and a theme color (accent 1), the same with a tint of 0.4, and an indexed color (10)
  against an ``rgb`` color of a different hue a few luma steps darker and lighter than their
  resolved RGB. (A theme color against the ``rgb`` of the same RGB would draw the same whatever
  wins, so it is not a case.) The top edges of C4:H4 and the left edges of C4:C9 are six styles
  alone. Beside it, across hidden column L, K's right side meets M's left side (rows 4, 6, 8),
  and across hidden row 13, the bottom sides of C12, E12 and G12 meet the top sides of C14,
  E14 and G14, the hidden cells' own sides double or thick.
- ``start_end``: borders whose XML holds ``<start>`` or ``<end>`` instead of ``<left>`` and
  ``<right>``, alone and beside a ``<left>`` or ``<right>`` of another style (written start,
  end, left, right in that order), on C4 to C12, rows 40 pt with empty columns B and D beside
  them; E4 to E14 hold the reference styles. Excel may refuse or repair the file (-50 with
  alerts off): it runs on its own, first, and a refusal is kept (``refused.txt``) and written as
  the group's one row, its ``name`` ``refused:`` and Excel's message.

Measured on each case's edge, 6 points clear of its corners: every device row across the edge,
4 points either side of the boundary, that carries ink (a pixel far from white and from the
gridline's gray), as its offset from the boundary in points (``offsets``, as ``borders.tsv``
gives them); on the row with the most ink, the on/off pattern (``capture.screen.ink.pattern``),
the distinct lengths of its whole on and off runs in points (``ons``, ``offs``) counted with the
pixels near the line's own color, so that slantDashDot's antialiased gaps count as gaps, and the
longest sequence of short dashes (3 points or less) between longer ones (``short_seq``, 1 for
dashDot, 2 for dashDotDot, capped at 3); the share of the rows' pixels that are ink
(``coverage``) and that are near the line's own color (``solid``); the line's color (the most
common ink pixel, ``RRGGBB``). Each edge is then matched to the nearest reference of its
orientation in the same window (:func:`identify`): ``drawn`` is that style, ``none`` without
ink, and ``distance`` how far it is. ``read`` writes ``tests/captures/border_conflicts.tsv``,
one row per group, edge and zoom.

Only one cell per column and one per row of each group is read from Excel (its geometry); every
other cell's edges are taken from its column's and its row's.
"""

from __future__ import annotations

import re
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Border
from openpyxl.styles import Side as OpenpyxlSide
from openpyxl.styles.colors import Color
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "border_conflicts"
OUTPUT = ROOT / "tests" / "captures" / "border_conflicts.tsv"
GROUPS = ("start_end", "styles", "colors")
REFUSALS = ("start_end",)
"""Groups whose file Excel may refuse: run on their own, a refusal is recorded as the answer."""
MARKERS = {"styles": ("B2", "R18"), "colors": ("B2", "O16"), "start_end": ("B2", "G16")}
ZOOM = 100
STYLES = (
    "thin", "medium", "thick", "double", "hair", "dotted", "dashed", "dashDot", "dashDotDot",
    "mediumDashed", "mediumDashDot", "mediumDashDotDot", "slantDashDot",
)  # fmt: skip
"""The 13 line styles, in the order of ``borders.py``."""
CLEAR = 6
"""Points kept clear of each corner, so that the lines across the edge's ends are not scanned."""
HALF = 4
"""Points scanned on each side of a boundary."""
SHORT = 3
"""A dash of this many points or fewer is short (a dashDot's dot); the long dashes are 9."""
COLUMNS = (
    "group", "kind", "name", "zoom", "cell", "side", "first", "second", "offsets", "pattern",
    "ons", "offs", "short_seq", "coverage", "solid", "color", "drawn", "distance",
)  # fmt: skip


@dataclass(frozen=True)
class Line:
    """One cell's side: its style and color, ``rgb`` (``RRGGBB``) as the color resolves.

    ``spec`` is how the file gives the color: ``RRGGBB``, ``theme4``, ``theme4+0.4`` (a tint) or
    ``indexed10``; ``rgb`` is the resolved color, as sheetpeek resolves it under openpyxl's
    default theme (the suite's tests check it).
    """

    style: str
    spec: str = "000000"
    rgb: str = "000000"

    def openpyxl(self) -> OpenpyxlSide:
        return OpenpyxlSide(style=self.style, color=_color(self.spec))

    def __str__(self) -> str:
        return self.style if self.spec == "000000" else f"{self.style} {self.spec}"


def _color(spec: str) -> Color:
    if spec.startswith("theme"):
        theme, _, tint = spec.removeprefix("theme").partition("+")
        return Color(theme=int(theme), tint=float(tint) if tint else 0.0)
    if spec.startswith("indexed"):
        return Color(indexed=int(spec.removeprefix("indexed")))
    return Color(rgb="FF" + spec)


@dataclass(frozen=True)
class Edge:
    """One measured edge: ``side`` of ``cell``; ``first`` the left or upper cell's side on it,
    ``second`` the right or lower cell's (``None``: that cell has no side there).

    ``kind`` is ``ref`` (a style alone), ``style``, ``color``, ``hidden`` or ``start_end``.
    """

    group: str
    kind: str
    name: str
    cell: str
    side: str
    first: Line | None
    second: Line | None

    @property
    def vertical(self) -> bool:
        return self.side in ("left", "right")


def _thin(spec: str, rgb: str | None = None) -> Line:
    return Line("thin", spec, rgb or spec)


COLOR_PAIRS = (
    ("601 vs 709", _thin("40C080"), _thin("FF8000")),
    ("601 vs 709", _thin("00FF00"), _thin("40C0C0")),
    ("601 vs max", _thin("004080"), _thin("A00000")),
    ("601 vs max and L*", _thin("604080"), _thin("FF0000")),
    ("601 vs mean", _thin("80A080"), _thin("FF8000")),
    ("601 vs mean", _thin("00FF40"), _thin("A0A060")),
    ("601 vs L*", _thin("00C0C0"), _thin("80A040")),
    ("equal luma", _thin("003223"), _thin("1400F0")),
    ("equal luma", _thin("2D5A69"), _thin("FA0505")),
    ("equal luma", _thin("23B923"), _thin("3787F0")),
    ("equal luma", _thin("5055B9"), _thin("D22D32")),
    ("theme vs darker rgb", _thin("theme4", "4F81BD"), _thin("F84800")),
    ("theme vs lighter rgb", _thin("theme4", "4F81BD"), _thin("F85800")),
    ("tinted theme vs darker rgb", _thin("theme4+0.4", "95B3D7"), _thin("60F000")),
    ("tinted theme vs lighter rgb", _thin("theme4+0.4", "95B3D7"), _thin("E8B800")),
    ("indexed vs darker rgb", _thin("indexed10", "FF0000"), _thin("0848F8")),
    ("indexed vs lighter rgb", _thin("indexed10", "FF0000"), _thin("1050F8")),
)
"""Thin sides in two colors: a label, then the two sides."""
COLOR_REFERENCES = ("thin", "medium", "thick", "double", "dashed", "hair")
RED, BLUE = Line("thick", "FF0000", "FF0000"), Line("thick", "0000FF", "0000FF")
HIDDEN = (
    ("thick red, double hidden, thick blue", RED, Line("double"), BLUE),
    ("medium, double hidden, thin", Line("medium"), Line("double"), Line("thin")),
    ("thin, thick hidden, dashed", Line("thin"), Line("thick"), Line("dashed")),
)
"""Across a hidden column or row: the first cell's side, the hidden cell's, the second's."""
START_END = (
    ("C4", "start thick", {"start": "thick"}),
    ("C6", "end medium", {"end": "medium"}),
    ("C8", "start thick, left thin", {"start": "thick", "left": "thin"}),
    ("C10", "end thick, right dashed", {"end": "thick", "right": "dashed"}),
    ("C12", "start double, end mediumDashed", {"start": "double", "end": "mediumDashed"}),
)
"""Each crafted border: its cell, its name and its elements' styles (black)."""
START_END_REFERENCES = ("thin", "medium", "thick", "double", "dashed", "mediumDashed")
TAGS = ("hair", "dotted", "dashDot", "dashDotDot", "slantDashDot")
"""Placeholder ``vertical`` styles that mark each crafted border until it is written."""


def _grid(columns: int, rows: int) -> tuple[list[str], list[int]]:
    return [get_column_letter(3 + k) for k in range(columns)], [4 + k for k in range(rows)]


def _pairs(
    group: str, kind: str, letters: list[str], numbers: list[int], vertical: list[Any],
    horizontal: list[Any], references: tuple[str, ...],
) -> list[Edge]:  # fmt: skip
    """A grid's edges: the references on its outer top and left edges, then the pairs."""
    found = []
    for k, style in enumerate(references):
        line = Line(style)
        found.append(Edge(group, "ref", f"ref h {style}", f"{letters[k]}4", "top", None, line))
        found.append(Edge(group, "ref", f"ref v {style}", f"C{numbers[k]}", "left", None, line))
    slots = [(f"{c}{r}", "right") for r in numbers for c in letters[:-1]]
    for (cell, side), (name, a, b) in zip(slots, vertical, strict=False):
        found.append(Edge(group, kind, f"v {name}", cell, side, a, b))
    slots = [(f"{c}{r}", "bottom") for r in numbers[:-1] for c in letters]
    for (cell, side), (name, a, b) in zip(slots, horizontal, strict=False):
        found.append(Edge(group, kind, f"h {name}", cell, side, a, b))
    return found


def _edges() -> list[Edge]:
    found = []
    letters, numbers = _grid(14, 13)
    pairs = [(f"{a}|{b}", Line(a), Line(b)) for a in STYLES for b in STYLES if a != b]
    flat = [(n.replace("|", "/"), a, b) for n, a, b in pairs]
    assert len(pairs) <= min(13 * 13, 12 * 14)  # the grid's vertical and horizontal edges
    found += _pairs("styles", "style", letters, numbers, pairs, flat, STYLES)
    letters, numbers = _grid(7, 6)
    pairs = []
    for label, a, b in COLOR_PAIRS:
        pairs += [(f"{label}: {a.spec}|{b.spec}", a, b), (f"{label}: {b.spec}|{a.spec}", b, a)]
    flat = [(n.replace("|", "/"), a, b) for n, a, b in pairs]
    assert len(pairs) <= min(6 * 6, 5 * 7)
    found += _pairs("colors", "color", letters, numbers, pairs, flat, COLOR_REFERENCES)
    for k, (name, a, _, b) in enumerate(HIDDEN):
        found.append(Edge("colors", "hidden", f"column {name}", f"K{4 + 2 * k}", "right", a, b))
        found.append(Edge("colors", "hidden", f"row {name}", f"{'CEG'[k]}12", "bottom", a, b))
    for cell, name, elements in START_END:
        left = elements.get("left", elements.get("start"))
        right = elements.get("right", elements.get("end"))
        for side, style in (("left", left), ("right", right)):
            line = Line(style) if style else None
            first, second = (None, line) if side == "left" else (line, None)
            found.append(
                Edge("start_end", "start_end", f"{name}: {side}", cell, side, first, second)
            )
    for k, style in enumerate(START_END_REFERENCES):
        line = Line(style)
        found.append(
            Edge("start_end", "ref", f"ref v {style}", f"E{4 + 2 * k}", "left", None, line)
        )
    return found


EDGES = _edges()

ORDER = (
    "double", "thick", "medium", "mediumDashed", "mediumDashDot", "slantDashDot",
    "mediumDashDotDot", "thin", "dashed", "dotted", "dashDot", "dashDotDot", "hair",
)  # fmt: skip
"""The style order sheetpeek draws by, strongest first: its expectation for the style pairs."""
EXPECTED = {
    "colors": {
        "40C080|FF8000": "thin 40C080", "00FF00|40C0C0": "thin 00FF00",
        "004080|A00000": "thin A00000", "604080|FF0000": "thin FF0000",
        "80A080|FF8000": "thin 80A080", "00FF40|A0A060": "thin A0A060",
        "00C0C0|80A040": "thin 00C0C0", "003223|1400F0": "thin 1400F0",
        "2D5A69|FA0505": "thin FA0505", "23B923|3787F0": "thin 3787F0",
        "5055B9|D22D32": "thin D22D32", "theme4|F84800": "thin F84800",
        "theme4|F85800": "thin 4F81BD", "theme4+0.4|60F000": "thin 60F000",
        "theme4+0.4|E8B800": "thin 95B3D7", "indexed10|0848F8": "thin 0848F8",
        "indexed10|1050F8": "thin FF0000",
    },
    "hidden": {
        "thick red, double hidden, thick blue": "thick 0000FF",
        "medium, double hidden, thin": "medium 000000",
        "thin, thick hidden, dashed": "thin 000000",
    },
    "start_end": {
        "start thick: left": "thick 000000", "start thick: right": "none",
        "end medium: left": "none", "end medium: right": "medium 000000",
        "start thick, left thin: left": "thin 000000", "start thick, left thin: right": "none",
        "end thick, right dashed: left": "none", "end thick, right dashed: right": "dashed 000000",
        "start double, end mediumDashed: left": "double 000000",
        "start double, end mediumDashed: right": "mediumDashed 000000",
    },
}  # fmt: skip
"""What sheetpeek draws, ``style RRGGBB`` or ``none``: for each color pair (in either order and
orientation), each pair across a hidden column or row, and each crafted border's edge."""


def expected(edge: Edge) -> str:
    """What sheetpeek draws on an edge: a reference its own style, a style pair the stronger by
    :data:`ORDER`, the rest by :data:`EXPECTED`."""
    a, b = edge.first, edge.second
    if edge.kind == "ref" and b is not None:
        return f"{b.style} {b.rgb}"
    if edge.kind == "style" and a is not None and b is not None:
        return f"{min(a.style, b.style, key=ORDER.index)} 000000"
    if edge.kind == "color" and a is not None and b is not None:
        pairs = EXPECTED["colors"]
        return pairs.get(f"{a.spec}|{b.spec}") or pairs[f"{b.spec}|{a.spec}"]
    if edge.kind == "hidden":
        return EXPECTED["hidden"][edge.name.split(" ", 1)[1]]
    return EXPECTED[edge.group][edge.name]


def edges(group: str | None = None) -> list[Edge]:
    """Every edge of ``group``, or of every group, in order."""
    return [e for e in EDGES if group is None or e.group == group]


# -- the workbooks -------------------------------------------------------------------------------


def _sides(group: str) -> dict[str, dict[str, Line]]:
    """Each cell's sides, from the group's edges (the hidden cells' too)."""
    found: dict[str, dict[str, Line]] = {}
    other = {"right": "left", "bottom": "top", "left": "right", "top": "bottom"}
    for edge in edges(group):
        if edge.kind == "start_end":
            continue
        neighbor = _neighbor(edge.cell, edge.side, 2 if edge.kind == "hidden" else 1)
        if edge.side in ("right", "bottom"):
            near, far = edge.first, edge.second
        else:
            near, far = edge.second, edge.first
        if near is not None:
            found.setdefault(edge.cell, {})[edge.side] = near
        if far is not None and neighbor is not None:
            found.setdefault(neighbor, {})[other[edge.side]] = far
    if group == "colors":
        for k, (_, _, hidden, _) in enumerate(HIDDEN):
            found[f"L{4 + 2 * k}"] = {"left": hidden, "right": hidden}
            found[f"{'CEG'[k]}13"] = {"top": hidden, "bottom": hidden}
    return found


def _neighbor(cell: str, side: str, step: int) -> str | None:
    letters, row = coordinate_from_string(cell)
    column = column_index_from_string(letters)
    dc, dr = {"right": (step, 0), "left": (-step, 0), "bottom": (0, step), "top": (0, -step)}[side]
    if column + dc < 1 or row + dr < 1:
        return None
    return f"{get_column_letter(column + dc)}{row + dr}"


def fill(ws: Any, group: str) -> list[Edge]:
    """Lay one group out on a worksheet; its edges."""
    if group == "styles":
        letters, numbers = _grid(14, 13)
        height = 56.0
    elif group == "colors":
        letters, numbers = [*_grid(7, 6)[0], "K", "L", "M"], [*_grid(7, 6)[1], 12, 13, 14]
        height = 30.0
        ws.column_dimensions["L"].hidden = True
        ws.row_dimensions[13].hidden = True
    else:
        letters, numbers = ["B", "C", "D", "E"], [4, 6, 8, 10, 12, 14]
        height = 40.0
    for letter in letters:
        ws.column_dimensions[letter].width = 10
    for number in numbers:
        ws.row_dimensions[number].height = height
    for cell, sides in _sides(group).items():
        ws[cell].border = Border(**{side: line.openpyxl() for side, line in sides.items()})
    if group == "start_end":
        for k, (cell, _, _) in enumerate(START_END):
            ws[cell].border = Border(vertical=OpenpyxlSide(style=TAGS[k]))
    return edges(group)


def crafted(text: str) -> str:
    """The styles with each placeholder border replaced by its ``start``/``end`` border."""
    for k, (_, _, elements) in enumerate(START_END):
        parts = "".join(
            f'<{name} style="{elements[name]}"><color rgb="FF000000"/></{name}>'
            if name in elements
            else f"<{name}/>"
            for name in ("start", "end", "left", "right", "top", "bottom", "diagonal")
            if name in elements or name not in ("start", "end", "left", "right")
        )
        text, count = re.subn(
            rf'<border><vertical style="{TAGS[k]}"\s*/></border>', f"<border>{parts}</border>", text
        )
        if count != 1:
            raise ValueError(f"placeholder border {TAGS[k]} not found once")
    return text


def postprocess(group: str) -> Any:
    """The ``start_end`` group's crafted borders; nothing for the others."""
    if group != "start_end":
        return None

    def edit(path: Path) -> None:
        from capture.screen.run import patch_part

        patch_part(path, {"xl/styles.xml": crafted})

    return edit


def reads(group: str) -> list[str]:
    """One cell of each column and of each row the group's edges use: the cells Excel reads."""
    found = [coordinate_from_string(e.cell) for e in edges(group)]
    letters = sorted({c for c, _ in found}, key=column_index_from_string)
    numbers = sorted({r for _, r in found})
    count = max(len(letters), len(numbers))
    return [
        f"{letters[min(k, len(letters) - 1)]}{numbers[min(k, len(numbers) - 1)]}"
        for k in range(count)
    ]


def build(path: Path) -> list[Edge]:
    """Write every group to a workbook of its own in the folder ``path``; every edge."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        wb.active.title = "Borders"
        found.extend(fill(wb.active, group))
        target = path / f"border_conflicts_{group}.xlsx"
        wb.save(target)
        edit = postprocess(group)
        if edit is not None:
            edit(target)
    return found


def run(groups: tuple[str, ...] | None = None) -> None:
    """Capture each group's window, then write the TSV (needs Excel).

    Without groups, every group but :data:`REFUSALS`, which run on their own first.
    """
    from capture.screen.run import capture
    from capture.screen.session import ExcelError

    if groups is None:
        groups = tuple(g for g in GROUPS if g not in REFUSALS)
    for group in groups:
        try:
            capture(
                lambda ws, group=group: fill(ws, group),
                cells=reads(group),
                out=BUILD_DIR / group,
                markers=MARKERS[group],
                zooms=(ZOOM,),
                name=f"border_conflicts_{group}",
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


def bounds(cap: Any, cell: str) -> dict[str, float]:
    """A cell's edges in image pixels, from a read cell of its column and one of its row."""
    letters, row = coordinate_from_string(cell)
    column = next((c for c in cap.geometry if coordinate_from_string(c)[0] == letters), None)
    line = next((c for c in cap.geometry if coordinate_from_string(c)[1] == row), None)
    if column is None or line is None:
        raise KeyError(f"{cell}: no cell of its column or row was read")
    return {
        "left": cap.edge(column, "left"),
        "right": cap.edge(column, "right"),
        "top": cap.edge(line, "top"),
        "bottom": cap.edge(line, "bottom"),
    }


def _hex(rgb: tuple[int, int, int]) -> str:
    return "".join(f"{c:02X}" for c in rgb)


def _distance(a: tuple[int, int, int], b: tuple[int, int, int]) -> float:
    return float(sum((x - y) ** 2 for x, y in zip(a, b, strict=True)) ** 0.5)


def _is_ink(rgb: tuple[int, int, int], grid: tuple[int, int, int] | None) -> bool:
    """Far from white, and from the gridline's gray where it is known."""
    if _distance(rgb, (255, 255, 255)) <= 60:
        return False
    return grid is None or _distance(rgb, grid) > 30


def _lengths(found: list[tuple[bool, int, int]], on: bool, scale: float) -> str:
    return " ".join(
        f"{v:g}" for v in sorted({round(n / scale * 2) / 2 for o, _, n in found if o == on})
    )


def _short_seq(found: list[tuple[bool, int, int]], scale: float) -> int:
    longest = current = 0
    for on, _, n in found:
        if not on:
            continue
        if n / scale <= SHORT + 0.25:
            current += 1
            longest = max(longest, current)
        else:
            current = 0
    return min(longest, 3)


def measure_edge(cap: Any, edge: Edge) -> dict[str, str]:
    """One edge's line: the rows that carry ink, their pattern, runs, shares and color."""
    from capture.screen.ink import pattern, pt
    from capture.screen.run import runs

    box = bounds(cap, edge.cell)
    boundary = round(box[edge.side])
    clear, half = round(CLEAR * cap.scale), round(HALF * cap.scale)
    a0, a1 = (box["top"], box["bottom"]) if edge.vertical else (box["left"], box["right"])
    along = range(round(a0) + clear, round(a1) - clear)
    grid = cap.gridline
    rows = []
    for offset in range(-half, half):
        fixed = boundary + offset
        pixels = [
            cap.image.pixel(fixed, a) if edge.vertical else cap.image.pixel(a, fixed) for a in along
        ]
        flags = [_is_ink(p, grid) for p in pixels]
        if any(flags):
            rows.append((offset, pixels, flags))
    found: dict[str, str] = {"offsets": " ".join(pt(cap, o) for o, _, _ in rows)}
    if not rows:
        return found | {"drawn": "none"}
    inked = Counter(p for _, pixels, flags in rows for p, f in zip(pixels, flags, strict=True) if f)
    color = inked.most_common(1)[0][0]
    total = sum(len(pixels) for _, pixels, _ in rows)
    near = [[_distance(p, color) <= 60 for p in pixels] for _, pixels, _ in rows]
    k = max(range(len(rows)), key=lambda i: sum(rows[i][2]))  # the row with the most ink
    inner = runs(near[k])[1:-1]  # its whole runs, the partial ones at the ends dropped
    return found | {
        "pattern": pattern(rows[k][2], cap.scale),
        "ons": _lengths(inner, True, cap.scale),
        "offs": _lengths(inner, False, cap.scale),
        "short_seq": str(_short_seq(inner, cap.scale)),
        "coverage": f"{sum(sum(f) for _, _, f in rows) / total:.2f}",
        "solid": f"{sum(sum(n) for n in near) / total:.2f}",
        "color": _hex(color),
    }


def _set(text: str) -> set[str]:
    return set(text.split())


def _jaccard(a: set[str], b: set[str]) -> float:
    return 0.0 if not (a | b) else 1 - len(a & b) / len(a | b)


def distance(found: dict[str, str], reference: dict[str, str]) -> float:
    """How unlike a measured line is a reference line: 0 for the same rows, runs and shares.

    Each device row present in one and not the other counts 10; the run lengths count up to 2
    each (on and off); the shares 2 a unit (a dash's phase moves them by up to 0.15 on a short
    edge); the short-dash sequence 2 a step (dashDot and dashDotDot differ in little else).
    """
    offsets = len(_set(found.get("offsets", "")) ^ _set(reference.get("offsets", "")))
    score = 10.0 * offsets
    score += 2 * _jaccard(_set(found.get("ons", "")), _set(reference.get("ons", "")))
    score += 2 * _jaccard(_set(found.get("offs", "")), _set(reference.get("offs", "")))
    for key in ("coverage", "solid"):
        score += 2 * abs(float(found.get(key) or 0) - float(reference.get(key) or 0))
    score += 2 * abs(int(found.get("short_seq") or 0) - int(reference.get("short_seq") or 0))
    return score


def identify(
    found: dict[str, str], references: dict[str, dict[str, str]]
) -> list[tuple[float, str]]:
    """The references by distance from a measured line, nearest first: ``[(distance, style)]``."""
    if not found.get("offsets"):
        return [(0.0, "none")]
    return sorted((distance(found, ref), style) for style, ref in references.items())


def measure(cap: Any, group: str) -> list[dict[str, str]]:
    """Every edge of a group at one zoom, each matched to the references of its orientation."""
    found = []
    references: dict[bool, dict[str, dict[str, str]]] = {True: {}, False: {}}
    for edge in edges(group):
        record = measure_edge(cap, edge)
        if edge.kind == "ref" and edge.second is not None:
            references[edge.vertical][edge.second.style] = record
        found.append((edge, record))
    records = []
    for edge, record in found:
        ranked = identify(record, references[edge.vertical])
        if ranked:
            score, style = ranked[0]
            record = {**record, "drawn": style, "distance": f"{score:.2f}"}
        records.append(
            {
                "group": group,
                "kind": edge.kind,
                "name": edge.name,
                "zoom": str(cap.zoom),
                "cell": edge.cell,
                "side": edge.side,
                "first": str(edge.first) if edge.first else "",
                "second": str(edge.second) if edge.second else "",
                **record,
            }
        )
    return records


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``border_conflicts.tsv``."""
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
            records.extend(measure(cap, group))
    with open(OUTPUT, "w", encoding="utf-8") as out_file:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out_file.write(f"# {key}: {value}\n")
        out_file.write("# suite: border_conflicts\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out_file.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} edges in {argv[1]}")
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
