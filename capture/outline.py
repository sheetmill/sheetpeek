"""Capture how Excel for Mac draws the outline gutter, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.outline run flag_yes
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.outline run
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.outline run light
    uv run --group capture python -m capture.outline read      # the saved captures again
    uv run --group capture python -m capture.outline build PATH  # every group's workbook

``run`` takes the dark appearance and ``run light`` the light one: each refuses to start in the
other, so a person switches macOS's appearance (System Settings > Appearance) once between them.
``run flag_yes`` goes first and alone: Excel may refuse that file (-50 with alerts off), which is
the answer, kept in ``refused.txt`` and written as the group's one row.

Each group is its own workbook and window (outline settings belong to a sheet); the second
marker sits below and right of what is measured. Groups:

- ``main``: ``ARCHITECTURE.md`` §4.11's sheet. Rows 4 to 17 grouped up to level 3 (rows 9 and 10
  hidden under the collapsed row 11, rows 14 to 16 hidden under the collapsed row 17), columns D
  to J up to level 2 (H and I hidden under the collapsed J); summaries below and right; no
  ``outlineLevelRow`` or ``outlineLevelCol``. Zoom 100 and 200, in both appearances.
- ``heights``: rows grouped to level 2 at 15, 30 and 60 pt (ticks, dots and minus buttons at
  each height, bars at 30 and 60, a plus at 60), columns D and E grouped under F, D 72 and F 114
  pt wide. Zoom 100 and 200.
- ``summaries``: ``summaryBelow="0"`` and ``summaryRight="0"``: rows 4 to 6 and columns D to F
  grouped to level 2 under the summary row 3 and column C, expanded; rows 9 and 10 and columns I
  and J hidden under the collapsed row 8 and column H. Zoom 100 and 200.
- ``declared`` and ``declared_small``: ``main`` with ``outlineLevelRow``/``outlineLevelCol`` 3/2
  (what Excel writes) and 1/1 (smaller than the levels). Zoom 100.
- ``view_off``, ``props_off``, ``view_on_props_off``, ``view_off_props_on``: ``main`` with
  ``showOutlineSymbols`` in the ``<sheetView>``, in the ``<outlinePr>``, or in both, disagreeing.
  Zoom 100.
- ``flag_yes``: ``main`` with ``<sheetView showOutlineSymbols="yes">`` (crafted). Zoom 100.

Measured, in screen points (image pixels over the display's scale), the gutter being found from
the window, not assumed: across row 1 near its top, from column A's left edge to the window's
edge, the row headings' band and the band beyond it (the row gutter: its width and color); up
column A near its left edge, the column headings' band and the band beyond it, taken as the
column gutter when its color is the row gutter's. Then, in each measured row's (column's)
stretch of the gutter, the ink (pixels more than 24 from the gutter's color in some channel),
cut into pieces across the lanes and along the row: a ``minus`` or ``plus`` button (lines at
least 9 points across; its border, paper and stroke colors and the stroke's size), a ``tick``
(4 to 9 across), a ``bar`` (thinner), or a lone ``dot`` (at most 4 square). Each piece is
written with its offset ``across`` from the gutter's outer edge (the window's edge for rows, the
gutter's top for columns) and ``along`` from the row's top (the column's left), and its size.
The level buttons are measured the same way where the row gutter meets the column headings and
where the column gutter meets the row headings, with the colors along their middle. ``read``
writes ``tests/captures/outline.tsv``, with an ``appearance`` column.

The capture settles §4.11's rules: the buttons (minus while a detail row shows, plus when all
are hidden or the summary is collapsed), the dots, the summaries on the left and above, the
depth from the levels with ``outlineLevelRow`` ignored, ``showOutlineSymbols`` in the view and in
``<outlinePr>`` and a value other than 0, 1, false and true, the geometry (gutter, lanes, bars,
ticks, buttons and dots at fixed screen sizes; a button centered on its row) and the colors in
both appearances.
"""

from __future__ import annotations

import itertools
import sys
from collections import Counter
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import Outline as OutlineProperties

# row -> (level, hidden, collapsed, height or None); column -> (level, hidden, collapsed, width)
MAIN_ROWS = {
    4: (1, False, False, None),
    5: (2, False, False, None),
    6: (3, False, False, None),
    7: (2, False, False, None),
    8: (1, False, False, None),
    9: (2, True, False, None),
    10: (2, True, False, None),
    11: (1, False, True, None),
    14: (1, True, False, None),
    15: (1, True, False, None),
    16: (1, True, False, None),
    17: (0, False, True, None),
}
MAIN_COLUMNS = {
    "D": (1, False, False, None),
    "E": (2, False, False, None),
    "F": (1, False, False, None),
    "H": (1, True, False, None),
    "I": (1, True, False, None),
    "J": (0, False, True, None),
}
HEIGHTS_ROWS = {
    4: (1, False, False, 15),
    5: (2, False, False, 30),
    6: (2, False, False, 60),
    7: (1, False, False, 30),
    8: (0, False, False, 60),
    10: (1, True, False, None),
    11: (1, True, False, None),
    12: (0, False, True, 60),
    14: (1, False, False, 60),
    15: (1, False, False, 30),
    16: (0, False, False, 15),
}
HEIGHTS_COLUMNS = {
    "D": (1, False, False, 13),
    "E": (1, False, False, None),
    "F": (0, False, False, 21),
}
SUMMARIES_ROWS = {
    3: (0, False, False, None),
    4: (1, False, False, None),
    5: (2, False, False, None),
    6: (2, False, False, None),
    8: (0, False, True, None),
    9: (1, True, False, None),
    10: (1, True, False, None),
}
SUMMARIES_COLUMNS = {
    "D": (1, False, False, None),
    "E": (2, False, False, None),
    "F": (2, False, False, None),
    "H": (0, False, True, None),
    "I": (1, True, False, None),
    "J": (1, True, False, None),
}


@dataclass(frozen=True)
class Group:
    """One workbook: its sheet, its settings, what is measured and where the second marker is."""

    rows: dict[int, tuple[int, bool, bool, int | None]]
    columns: dict[str, tuple[int, bool, bool, int | None]]
    last_row: int
    last_column: str
    marker: str
    zooms: tuple[int, ...] = (100,)
    below: bool = True
    right: bool = True
    view: str | None = None
    """``showOutlineSymbols`` of the ``<sheetView>``, as written; ``None`` leaves it out."""
    props: str | None = None
    """``showOutlineSymbols`` of the ``<outlinePr>``; ``None`` leaves it out."""
    declared: tuple[int, int] | None = None
    """``outlineLevelRow`` and ``outlineLevelCol``; ``None`` leaves them out."""
    detail: bool = False
    """Every row and column is measured (else only the gutter and level buttons)."""


def _main(**settings: Any) -> Group:
    return Group(MAIN_ROWS, MAIN_COLUMNS, 18, "K", "L20", **settings)


GROUPS: dict[str, Group] = {
    "main": _main(zooms=(100, 200), detail=True),
    "heights": Group(HEIGHTS_ROWS, HEIGHTS_COLUMNS, 17, "G", "H18", (100, 200), detail=True),
    "summaries": Group(
        SUMMARIES_ROWS, SUMMARIES_COLUMNS, 12, "K", "M13", (100, 200), below=False, right=False,
        detail=True,
    ),
    "declared": _main(declared=(3, 2)),
    "declared_small": _main(declared=(1, 1)),
    "view_off": _main(view="0"),
    "props_off": _main(props="0"),
    "view_on_props_off": _main(view="1", props="0"),
    "view_off_props_on": _main(view="0", props="1"),
    "flag_yes": _main(view="yes"),
}  # fmt: skip
REFUSALS = ("flag_yes",)
"""Groups whose file Excel may refuse: run on their own, a refusal is the answer."""
LIGHT = ("main",)
"""The groups ``run light`` captures again in the light appearance."""
APPEARANCES = ("dark", "light")
INK = 24
"""A gutter pixel is ink when one channel is more than this from the gutter's color."""
BAND = 6
"""Neighboring pixels within this of a band's first color belong to the band."""
WIDE, MID, DOT = 9, 4, 4
"""Screen points across: a button's lines, a tick's lines; a dot's largest side."""
ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "outline"
OUTPUT = ROOT / "tests" / "captures" / "outline.tsv"
COLUMNS = (
    "appearance", "group", "zoom", "axis", "index", "size", "expected", "kind", "across",
    "along", "across_size", "along_size", "ink", "paper", "stroke", "stroke_size", "strip",
)  # fmt: skip

EXPECTED: dict[str, Any] = {
    "main": {
        "lanes": (4, 3),
        "rows": {
            4: "1:start 2:dot", 5: "1:bar 2:start 3:dot", 6: "1:bar 2:bar 3:start 4:dot",
            7: "1:bar 2:bar 3:open", 8: "1:bar 2:open", 11: "1:bar 2:closed", 12: "1:open",
            17: "1:closed",
        },
        "columns": {
            "D": "1:start 2:dot", "E": "1:bar 2:start 3:dot", "F": "1:bar 2:open",
            "G": "1:open", "J": "1:closed",
        },
    },
    "heights": {
        "lanes": (3, 2),
        "rows": {
            4: "1:start 2:dot", 5: "1:bar 2:start 3:dot", 6: "1:bar 2:bar 3:dot",
            7: "1:bar 2:open", 8: "1:open", 12: "1:closed", 14: "1:start 2:dot",
            15: "1:bar 2:dot", 16: "1:open",
        },
        "columns": {"D": "1:start 2:dot", "E": "1:bar 2:dot", "F": "1:open"},
    },
    "summaries": {
        "lanes": (3, 3),
        "rows": {
            3: "1:open", 4: "1:bar 2:open", 5: "1:bar 2:bar 3:dot", 6: "1:start 2:start 3:dot",
            8: "1:closed",
        },
        "columns": {
            "C": "1:open", "D": "1:bar 2:open", "E": "1:bar 2:bar 3:dot",
            "F": "1:start 2:start 3:dot", "H": "1:closed",
        },
    },
    **dict.fromkeys(
        ("declared", "declared_small", "props_off", "view_on_props_off", "flag_yes"),
        "the same as main",
    ),
    **dict.fromkeys(("view_off", "view_off_props_on")),
}  # fmt: skip
"""What sheetpeek lays out for each group: per row and column, ``lane:kind`` of each mark (the
gutters' lanes, rows and columns); ``"the same as main"``; ``None`` draws no outline."""


def fill(ws: Any, name: str) -> Group:
    """Lay one group's outline out on a worksheet."""
    group = GROUPS[name]
    for row in range(1, group.last_row + 1):
        ws.cell(row, 1, row)
    for column in range(2, openpyxl.utils.column_index_from_string(group.last_column) + 1):
        ws.cell(1, column, get_column_letter(column))
    for row, (level, hidden, collapsed, height) in group.rows.items():
        dimension = ws.row_dimensions[row]
        dimension.outlineLevel, dimension.hidden, dimension.collapsed = level, hidden, collapsed
        if height is not None:
            dimension.height = height
    for letter, (level, hidden, collapsed, width) in group.columns.items():
        dimension = ws.column_dimensions[letter]
        dimension.outlineLevel, dimension.hidden, dimension.collapsed = level, hidden, collapsed
        if width is not None:
            dimension.width = width
    if not (group.below and group.right) or group.props in ("0", "1"):
        properties = OutlineProperties(summaryBelow=group.below, summaryRight=group.right)
        if group.props is not None:
            properties.showOutlineSymbols = group.props == "1"
        ws.sheet_properties.outlinePr = properties
    if group.view in ("0", "1"):
        ws.sheet_view.showOutlineSymbols = group.view == "1"
    return group


def postprocess(name: str) -> Callable[[Path], None] | None:
    """What a group's saved workbook needs that openpyxl does not write, or ``None``: a
    ``showOutlineSymbols`` other than 0 and 1, and ``outlineLevelRow``/``outlineLevelCol``
    (openpyxl writes the row's alone)."""
    group = GROUPS[name]
    view = None if group.view in (None, "0", "1") else group.view
    declared = group.declared
    if view is None and declared is None:
        return None

    def crafted(path: Path) -> None:
        import re

        from capture.screen.run import patch_part

        def edit(text: str) -> str:
            if view is not None:
                if text.count("<sheetView ") != 1:
                    raise SystemExit("expected one <sheetView> in the sheet")
                text = text.replace("<sheetView ", f'<sheetView showOutlineSymbols="{view}" ')
            if declared is not None:
                text, count = re.subn(
                    r"<sheetFormatPr ([^>]*?)\s*/>",
                    lambda m: (
                        f'<sheetFormatPr {m[1]} outlineLevelRow="{declared[0]}" '
                        f'outlineLevelCol="{declared[1]}" />'
                    ),
                    re.sub(r'\s*outlineLevel(?:Row|Col)="\d+"', "", text),
                )
                if count != 1:
                    raise SystemExit("expected one <sheetFormatPr> in the sheet")
            return text

        patch_part(path, {"xl/worksheets/sheet1.xml": edit})

    return crafted


def build(path: Path) -> list[str]:
    """Write every group to a workbook of its own in the folder ``path``; the group names."""
    path.mkdir(parents=True, exist_ok=True)
    for name in GROUPS:
        wb = openpyxl.Workbook()
        fill(wb.active, name)
        target = path / f"outline_{name}.xlsx"
        wb.save(target)
        post = postprocess(name)
        if post is not None:
            post(target)
    return list(GROUPS)


def measured(group: Group) -> tuple[list[int], list[str]]:
    """The rows and the column letters whose gutter stretch is measured."""
    if not group.detail:
        return [1], ["A"]
    last = openpyxl.utils.column_index_from_string(group.last_column)
    return list(range(1, group.last_row + 1)), [get_column_letter(c) for c in range(1, last + 1)]


def cells(group: Group) -> list[str]:
    """The cells whose geometry is read: column A for each row, row 1 for each column."""
    rows, columns = measured(group)
    return list(dict.fromkeys([*(f"A{r}" for r in rows), *(f"{c}1" for c in columns)]))


def _appearance() -> str:
    from capture.screen.session import environment

    return environment("unknown")["appearance"]


def run(names: tuple[str, ...] | None = None, appearance: str = "dark") -> None:
    """Capture each group's window in ``appearance``, then write the TSV (needs Excel).

    Without names: every group but :data:`REFUSALS` in the dark, :data:`LIGHT` in the light.
    """
    from capture.screen.run import capture
    from capture.screen.session import ExcelError

    found = _appearance()
    if found != appearance:
        raise SystemExit(f"macOS shows the {found} appearance; this run needs the {appearance}")
    if names is None:
        names = LIGHT if appearance == "light" else tuple(g for g in GROUPS if g not in REFUSALS)
    for name in names:
        group = GROUPS[name]
        try:
            capture(
                lambda ws, name=name: fill(ws, name),
                cells=cells(group),
                out=BUILD_DIR / appearance / name,
                markers=("B2", group.marker),
                zooms=group.zooms,
                name=f"outline_{name}",
                postprocess=postprocess(name),
            )
        except ExcelError as error:
            if name not in REFUSALS:
                raise
            record_refusal(appearance, name, str(error))
            print(f"{name}: Excel refused the file", file=sys.stderr)
    write()


def record_refusal(appearance: str, name: str, message: str) -> None:
    """Keep Excel's refusal of a group's file in ``refused.txt``, without the script's path."""
    out = BUILD_DIR / appearance / name
    out.mkdir(parents=True, exist_ok=True)
    reason = message.rpartition("execution error: ")[2].strip() or message
    (out / "refused.txt").write_text(reason + "\n", encoding="utf-8")


# -- measuring ------------------------------------------------------------------------------------

RGB = tuple[int, int, int]


def _hex(rgb: RGB) -> str:
    return "".join(f"{c:02X}" for c in rgb)


def _near(a: RGB, b: RGB, tolerance: int) -> bool:
    return max(abs(x - y) for x, y in zip(a, b, strict=True)) <= tolerance


def _mode(pixels: list[RGB]) -> str:
    return _hex(Counter(pixels).most_common(1)[0][0]) if pixels else ""


def _bands(colors: list[RGB]) -> list[tuple[RGB, int]]:
    """Runs of pixels within :data:`BAND` of each run's first color: ``[(color, length)]``."""
    found: list[list[Any]] = []
    for color in colors:
        if found and _near(color, found[-1][0], BAND):
            found[-1][1] += 1
        else:
            found.append([color, 1])
    return [(color, length) for color, length in found]


def runs_text(colors: list[RGB]) -> str:
    """A strip as runs of one exact color, ``RRGGBBxN`` each, as ``tables`` writes them."""
    found: list[list[Any]] = []
    for color in colors:
        if found and found[-1][0] == color:
            found[-1][1] += 1
        else:
            found.append([color, 1])
    return " ".join(f"{_hex(color)}x{count}" for color, count in found)


def _split(bands: list[tuple[RGB, int]]) -> tuple[int, tuple[RGB, int] | None, int]:
    """The headings' extent (with the thin bands before it), the next wide band and the thin
    bands between: from a strip that starts at a cell edge and runs outward."""
    i, headings = 0, 0
    while i < len(bands) and bands[i][1] <= 2:
        headings += bands[i][1]
        i += 1
    if i == len(bands):
        return headings, None, 0
    headings += bands[i][1]
    i += 1
    between = 0
    while i < len(bands) and bands[i][1] <= 2:
        between += bands[i][1]
        i += 1
    return headings, (bands[i] if i < len(bands) else None), between


@dataclass
class Gutter:
    """Where the gutters lie in a capture, in image pixels, and their color."""

    per: float
    """Image pixels a screen point."""
    color: RGB | None
    row_outer: int
    """The row gutter's outer edge (x); equal to ``row_inner`` when there is none."""
    row_inner: int
    """The row headings' outer edge (x)."""
    column_outer: int
    column_inner: int
    """The column headings' outer edge (y)."""
    cell_x: int
    cell_y: int
    """Column A's left and row 1's top."""
    row_strip: list[RGB]
    column_strip: list[RGB]
    column_candidate: tuple[RGB, int] | None


def gutter(cap: Any) -> Gutter:
    """Find the headings and the gutters left of and above them, from row 1 and column A."""
    per = cap.scale * 100 / cap.zoom
    left, top = round(cap.edge("A1", "left")), round(cap.edge("A1", "top"))
    height = cap.edge("A1", "bottom") - cap.edge("A1", "top")
    width = cap.edge("A1", "right") - cap.edge("A1", "left")
    y = round(top + 0.15 * height)
    row_strip = [cap.image.pixel(x, y) for x in range(left - 1, -1, -1)]
    headings, candidate, between = _split(_bands(row_strip))
    row_inner = left - headings
    color: RGB | None = None
    row_outer = row_inner
    if candidate is not None:
        color = candidate[0]
        row_outer = row_inner - between - candidate[1]
    x = round(left + 0.15 * width)
    reach = max(0, top - 1 - round(160 * per))
    column_strip = [cap.image.pixel(x, yy) for yy in range(top - 1, reach - 1, -1)]
    headings, column_candidate, between = _split(_bands(column_strip))
    column_inner = top - headings
    column_outer = column_inner
    if column_candidate is not None and color is not None and _near(column_candidate[0], color, 12):
        column_outer = column_inner - between - column_candidate[1]
    return Gutter(
        per, color, row_outer, row_inner, column_outer, column_inner, left, top,
        row_strip, column_strip, column_candidate,
    )  # fmt: skip


@dataclass
class Piece:
    kind: str
    across: int
    along: int
    across_end: int
    along_end: int
    """Image pixels, ends exclusive, relative to the region's corner."""
    ink: str = ""
    paper: str = ""
    stroke: str = ""
    stroke_size: str = ""


def pieces(
    pixel: Callable[[int, int], RGB],
    across: tuple[int, int],
    along: tuple[int, int],
    paper: RGB,
    per: float,
) -> list[Piece]:
    """The ink in a region of the gutter, cut into buttons, ticks, bars and dots.

    ``pixel(a, b)`` is the image's pixel at ``a`` across the lanes and ``b`` along the row (or
    column); the region runs over ``across`` and ``along`` (image pixels, ends exclusive).
    """
    a0, a1 = across
    b0, b1 = along
    span = a1 - a0
    if span <= 0 or b1 <= b0:
        return []
    ink: dict[tuple[int, int], RGB] = {}
    for b in range(b0, b1):
        line = {}
        for a in range(a0, a1):
            rgb = pixel(a, b)
            if not _near(rgb, paper, INK):
                line[(a, b)] = rgb
        if len(line) < 0.97 * span:  # a line across the whole gutter is a separator
            ink.update(line)
    columns = sorted({a for a, _ in ink})
    groups: list[list[int]] = []
    for a in columns:
        if groups and a == groups[-1][-1] + 1:
            groups[-1].append(a)
        else:
            groups.append([a])
    found = []
    for group in groups:
        for first, last in _boxes(group, ink, per):
            mine = {(a, b): rgb for (a, b), rgb in ink.items() if first <= a <= last}
            found.extend(_cut(mine, per, (a0, b0)))
    return found


def _boxes(group: list[int], ink: dict[tuple[int, int], RGB], per: float) -> list[tuple[int, int]]:
    """A run of inked columns as boxes that touch (level buttons side by side), split at their
    borders: columns in the first column's color over at least 80 % of the run's length."""
    first, last = group[0], group[-1]
    if last - first + 1 < 1.5 * WIDE * per:
        return [(first, last)]
    along = [b for (a, b) in ink if first <= a <= last]
    length = max(along) - min(along) + 1
    border = Counter(rgb for (a, _), rgb in ink.items() if a == first).most_common(1)[0][0]
    count = Counter(
        a for (a, _), rgb in ink.items() if first <= a <= last and _near(rgb, border, INK)
    )
    borders = [a for a in group if count[a] >= 0.8 * length]
    runs: list[list[int]] = []
    for a in borders:
        if runs and a <= runs[-1][-1] + 1:
            runs[-1].append(a)
        else:
            runs.append([a])
    found = []
    for left, right in itertools.pairwise(runs):
        if right[0] - left[-1] >= 6 * per:  # a button's inside lies between two borders
            found.append((left[0], right[-1]))
    if not found:
        return [(first, last)]
    # a run shared by two touching buttons is the right border of one and the left of the next
    boxes = []
    for k, (start, end) in enumerate(found):
        if k and start <= boxes[-1][1]:
            middle = (start + boxes[-1][1]) // 2
            boxes[-1] = (boxes[-1][0], middle)
            start = middle + 1
        boxes.append((start, end))
    return boxes


def _cut(ink: dict[tuple[int, int], RGB], per: float, corner: tuple[int, int]) -> list[Piece]:
    """One lane's ink cut along the row into pieces by how wide each line is."""
    lines: dict[int, list[int]] = {}
    for a, b in ink:
        lines.setdefault(b, []).append(a)
    extent = {b: (min(xs), max(xs) + 1) for b, xs in lines.items()}
    a_min = min(a for a, _ in ink)
    a_max = max(a for a, _ in ink) + 1
    b_min, b_max = min(lines), max(lines) + 1
    if a_max - a_min <= DOT * per and b_max - b_min <= DOT * per:
        return [_piece("dot", ink, (a_min, b_min, a_max, b_max), corner)]

    def kind(b: int) -> str:
        if b not in extent:
            return ""
        width = extent[b][1] - extent[b][0]
        return "button" if width >= WIDE * per else "tick" if width >= MID * per else "bar"

    kinds = {b: kind(b) for b in range(b_min, b_max)}
    # a button's rounded corners make its first and last lines narrower: give them to the button
    reach = round(1.5 * per)
    for b in [b for b, k in kinds.items() if k == "button"]:
        for step in (-1, 1):
            for k in range(1, reach + 1):
                if kinds.get(b + step * k) == "tick":
                    kinds[b + step * k] = "button"
                else:
                    break
    found = []
    start = b_min
    for b in range(b_min, b_max + 1):
        if b == b_max or kinds[b] != kinds[start]:
            if kinds[start]:
                part = {(a, bb): rgb for (a, bb), rgb in ink.items() if start <= bb < b}
                if part:
                    box = (
                        min(a for a, _ in part),
                        start,
                        max(a for a, _ in part) + 1,
                        b,
                    )
                    found.append(_piece(kinds[start], part, box, corner))
            start = b
    return found


def _piece(
    kind: str, ink: dict[tuple[int, int], RGB], box: tuple[int, int, int, int], corner: Any
) -> Piece:
    a0, b0, a1, b1 = box
    piece = Piece(kind, a0 - corner[0], b0 - corner[1], a1 - corner[0], b1 - corner[1])
    if kind != "button":
        piece.ink = _mode(list(ink.values()))
        return piece
    border = [
        rgb
        for (a, b), rgb in ink.items()
        if a == a0 and b0 + (b1 - b0) // 4 <= b < b1 - (b1 - b0) // 4
    ]
    piece.ink = _mode(border)
    inset = max(2, (a1 - a0) // 6)
    inner = {
        (a, b): rgb
        for (a, b), rgb in ink.items()
        if a0 + inset <= a < a1 - inset and b0 + inset <= b < b1 - inset
    }
    if not inner:
        return piece
    paper = Counter(inner.values()).most_common(1)[0][0]
    piece.paper = _hex(paper)
    strokes = {k: rgb for k, rgb in inner.items() if not _near(rgb, paper, 40)}
    if not strokes:
        piece.kind = "button"
        return piece
    sa0, sa1 = min(a for a, _ in strokes), max(a for a, _ in strokes) + 1
    sb0, sb1 = min(b for _, b in strokes), max(b for _, b in strokes) + 1
    piece.stroke = _mode(list(strokes.values()))
    piece.stroke_size = f"{sa1 - sa0},{sb1 - sb0}"  # image pixels; written in points
    # A plus strokes both ways, a minus one way: compare the extents without the gutter's axis
    # (across and along swap between the row and the column gutters).
    short, long = sorted((sb1 - sb0, sa1 - sa0))
    piece.kind = "plus" if short >= long / 2 and short > 2 else "minus"
    return piece


def _spt(pixels: float, per: float) -> str:
    return f"{round(pixels / per, 2):g}"


def _record(cap: Any, found: Gutter, piece: Piece, base: dict[str, str]) -> dict[str, str]:
    per = found.per
    stroke = ""
    if piece.stroke_size:
        w, h = (int(v) for v in piece.stroke_size.split(","))
        stroke = f"{_spt(w, per)},{_spt(h, per)}"
    return base | {
        "kind": piece.kind,
        "across": _spt(piece.across, per),
        "along": _spt(piece.along, per),
        "across_size": _spt(piece.across_end - piece.across, per),
        "along_size": _spt(piece.along_end - piece.along, per),
        "ink": piece.ink,
        "paper": piece.paper,
        "stroke": piece.stroke,
        "stroke_size": stroke,
    }


def _expected(name: str, axis: str, index: Any) -> str:
    expected = EXPECTED[name]
    if expected == "the same as main":
        expected = EXPECTED["main"]
    if expected is None:
        return ""
    return str(expected[axis].get(index, ""))


def measure(cap: Any, name: str) -> list[dict[str, str]]:
    """One group's records at one zoom: the gutters, the level buttons, every measured row and
    column."""
    group = GROUPS[name]
    found = gutter(cap)
    per = found.per
    zoom = str(cap.zoom)
    records = []
    paper = found.color
    rows_on = found.row_outer < found.row_inner
    columns_on = found.column_outer < found.column_inner
    records.append({
        "group": name, "zoom": zoom, "axis": "gutter", "index": "rows",
        "across_size": _spt(found.row_inner - found.row_outer, per) if rows_on else "",
        "along_size": _spt(found.cell_x - found.row_inner, per),
        "ink": _hex(paper) if paper else "",
        "strip": runs_text(found.row_strip[: round(160 * per)]),
    })  # fmt: skip
    candidate = found.column_candidate
    records.append({
        "group": name, "zoom": zoom, "axis": "gutter", "index": "columns",
        "across_size": _spt(found.column_inner - found.column_outer, per) if columns_on else "",
        "along_size": _spt(found.cell_y - found.column_inner, per),
        "ink": _hex(candidate[0]) if candidate else "",
        "strip": runs_text(found.column_strip),
    })  # fmt: skip
    if paper is None:
        return records
    image = cap.image

    def by_row(a: int, b: int) -> RGB:
        return image.pixel(a, b)

    def by_column(a: int, b: int) -> RGB:
        return image.pixel(b, a)

    levels = []
    if rows_on:
        span = (found.row_outer, found.row_inner)
        band = (found.column_inner, found.cell_y)
        middle = (band[0] + band[1]) // 2
        strip = runs_text([image.pixel(a, middle) for a in range(*span)])
        levels.append(("rows", by_row, span, band, strip))
    if columns_on:
        span = (found.column_outer, found.column_inner)
        band = (found.row_inner, found.cell_x)
        middle = (band[0] + band[1]) // 2
        strip = runs_text([image.pixel(middle, a) for a in range(*span)])
        levels.append(("columns", by_column, span, band, strip))
    for index, pixel, span, band, strip in levels:
        base = {"group": name, "zoom": zoom, "axis": "levels", "index": index, "strip": strip}
        base["size"] = _spt(band[1] - band[0], per)
        got = pieces(pixel, span, band, paper, per)
        for piece in got:
            if piece.kind in ("minus", "plus", "button"):
                piece.kind = "level"  # its digit is no stroke
        records += [_record(cap, found, p, base) for p in got] or [base | {"kind": "none"}]
    rows, columns = measured(group)
    if not group.detail:
        return records
    for row in rows if rows_on else []:
        cell = f"A{row}"
        top, bottom = round(cap.edge(cell, "top")), round(cap.edge(cell, "bottom"))
        if bottom - top <= 0:
            continue  # hidden
        base = {
            "group": name, "zoom": zoom, "axis": "rows", "index": str(row),
            "size": _spt(cap.edge(cell, "bottom") - cap.edge(cell, "top"), per),
            "expected": _expected(name, "rows", row),
        }  # fmt: skip
        got = pieces(by_row, (found.row_outer, found.row_inner), (top, bottom), paper, per)
        records += [_record(cap, found, p, base) for p in got] or [base | {"kind": "none"}]
    for letter in columns if columns_on else []:
        cell = f"{letter}1"
        left, right = round(cap.edge(cell, "left")), round(cap.edge(cell, "right"))
        if right - left <= 0:
            continue
        base = {
            "group": name, "zoom": zoom, "axis": "columns", "index": letter,
            "size": _spt(cap.edge(cell, "right") - cap.edge(cell, "left"), per),
            "expected": _expected(name, "columns", letter),
        }  # fmt: skip
        span = (found.column_outer, found.column_inner)
        got = pieces(by_column, span, (left, right), paper, per)
        records += [_record(cap, found, p, base) for p in got] or [base | {"kind": "none"}]
    return records


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``outline.tsv``."""
    from capture.screen.run import header, load

    records = []
    first: dict[str, str] = {}
    for appearance in APPEARANCES:
        for name in GROUPS:
            out = BUILD_DIR / appearance / name
            refused = out / "refused.txt"
            if refused.exists():
                reason = refused.read_text(encoding="utf-8").strip()
                records.append({
                    "appearance": appearance, "group": name, "kind": f"refused: {reason}",
                })  # fmt: skip
                continue
            if not (out / "formats.tsv").exists():
                print(f"{appearance} {name}: not captured yet, left out", file=sys.stderr)
                continue
            found = header(out)
            first = first or found
            shown = {"appearance": found.get("appearance", appearance)}
            records += [shown | record for cap in load(out) for record in measure(cap, name)]
    with open(OUTPUT, "w", encoding="utf-8") as out_file:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after", "zooms", "appearance"):
                out_file.write(f"# {key}: {value}\n")
        out_file.write("# suite: outline\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out_file.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        names = build(Path(argv[1]))
        print(f"{len(names)} workbooks in {argv[1]}")
        return 0
    if argv == ["run", "light"]:
        run(appearance="light")
        return 0
    if argv[:1] == ["run"] and all(name in GROUPS for name in argv[1:]):
        run(tuple(argv[1:]) or None)
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
