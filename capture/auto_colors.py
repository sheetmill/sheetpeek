"""Capture how Excel for Mac draws automatic, system and missing colors, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.auto_colors run empty
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.auto_colors run
    uv run --group capture python -m capture.auto_colors read           # the saved captures again
    uv run --group capture python -m capture.auto_colors build PATH     # every group's workbook

ARCHITECTURE.md §3.3 (a color element without attributes), §4.12 (automatic fill and border
colors, a missing fill color element) and §4.14 (automatic colors: indexed 64, 65 and 81).
Each group is its own workbook and window, at zoom 100, in the dark appearance as the other
suites. Every case is one cell, 12 characters wide and 24 pt high, with an empty narrow column
and a 7 pt row on each side, so that its edges are its own. The workbooks carry no
``<colors>`` element: the default palette applies, as in a file Excel writes. A color the case
removes or empties is written by openpyxl as a placeholder ``rgb`` and rewritten in
``xl/styles.xml`` before Excel opens the file (``postprocess``). Groups:

- ``main``: solid fills whose ``fgColor`` is indexed 64, ``auto``, indexed 65, indexed 81, or
  missing (with a yellow ``bgColor``, and with none), indexed 64 and ``auto`` over a yellow
  ``bgColor``, and a blue control; ``lightGrid`` fills (foreground 7 of 16 pixels, so the two
  colors cannot be confused) in indexed 64 on 65, 65 on 64, ``auto`` on ``auto``, 81 on yellow,
  blue on 64, a missing ``fgColor`` on yellow, blue with ``bgColor`` missing, both missing, and
  blue on yellow; fonts ("HH") in ``auto``, indexed 64, 81, no color element, blue, and indexed
  65 over a gray fill (white text would not show on white); medium borders on all four sides
  in ``auto``, indexed 64, no color element, indexed 65, 81, and blue.
- ``empty``: ``<fgColor/>``, ``<bgColor/>`` and ``<color/>`` with no attributes on a solid fill
  (alone and over yellow), on each color of a ``lightGrid`` fill, on a font and on a border, and
  a blue solid control. Excel may refuse or repair the file (-50 with alerts off): it runs on
  its own, first, and a refusal is kept (``refused.txt``) and written as the group's one row,
  its ``name`` ``refused:`` and Excel's message.

Measured per case: the two most common colors inside the cell, 4 image pixels in from its
edges, with their shares (``interior``, ``RRGGBB:share``); per edge (left, right, top, bottom),
whether the gridline's gray shows within 3 image pixels of the boundary at the edge's middle
(``gridlines``, ``1`` shows; a fill covers its gridlines, no fill leaves them); for a font, the
text's ink against the cell's own most common color (``ink_color``, the mean of the tenth of
pixels farthest from it; ``ink_pixels``); for a border, the colors across each edge from 3
screen points before it to 3 after, as runs (``edge_*``, as the ``tables`` suite writes them);
and Excel's own reads: ``display format``'s fill, pattern and font color, and the left side's
style and color. ``read`` writes ``tests/captures/auto_colors.tsv``, one row per group and case.
"""

from __future__ import annotations

import dataclasses
import re
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Border, Color, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "auto_colors"
OUTPUT = ROOT / "tests" / "captures" / "auto_colors.tsv"
GROUPS = ("main", "empty")
REFUSALS = ("empty",)
"""Groups whose file Excel may refuse: run on their own, a refusal is recorded as the answer."""
ZOOM = 100
SLOTS = ("C", "E", "G", "I", "K", "M")
"""The case columns, 12 characters wide; the columns between them are 2 wide."""
INSET = 4
REACH = 3
"""Image pixels on each side of a boundary searched for the gridline's gray."""
BLUE, YELLOW, GRAY = "0000FF", "FFFF00", "808080"
GRID_SHARE = 7 / 16
"""The foreground's share of a ``lightGrid`` tile."""
COLUMNS = (
    "group", "name", "kind", "cell", "zoom", "interior", "gridlines", "ink_color", "ink_pixels",
    "edge_left", "edge_right", "edge_top", "edge_bottom", "display_fill", "display_pattern",
    "display_font_color", "border_style", "border_color",
)  # fmt: skip
READS = {
    "display_fill": "display.fill",
    "display_pattern": "display.pattern",
    "display_font_color": "display.font_color",
    "border_style": "own.border.left.style",
    "border_color": "own.border.left.color",
}
EDGES = ("left", "right", "top", "bottom")


@dataclass(frozen=True)
class Case:
    """One measured cell.

    Colors are written ``rgb:RRGGBB``, ``auto``, ``indexed:N``, ``missing`` (no element) or
    ``empty`` (an element without attributes).
    """

    group: str
    name: str
    kind: str
    """``fill``, ``font`` or ``border``."""
    pattern: str | None = None
    fg: str = "missing"
    bg: str = "missing"
    color: str = "missing"
    """A font's or a border's color."""
    under: str | None = None
    """A solid fill under a font, ``RRGGBB``."""
    cell: str = ""


def _fill(group: str, name: str, pattern: str, fg: str, bg: str = "missing") -> Case:
    return Case(group, name, "fill", pattern=pattern, fg=fg, bg=bg)


_SPECS = [
    _fill("main", "solid fg indexed 64", "solid", "indexed:64"),
    _fill("main", "solid fg auto", "solid", "auto"),
    _fill("main", "solid fg indexed 65", "solid", "indexed:65"),
    _fill("main", "solid fg indexed 81", "solid", "indexed:81"),
    _fill("main", "solid fg missing, bg yellow", "solid", "missing", f"rgb:{YELLOW}"),
    _fill("main", "solid fg and bg missing", "solid", "missing"),
    _fill("main", "solid fg indexed 64, bg yellow", "solid", "indexed:64", f"rgb:{YELLOW}"),
    _fill("main", "solid fg auto, bg yellow", "solid", "auto", f"rgb:{YELLOW}"),
    _fill("main", "solid blue", "solid", f"rgb:{BLUE}"),
    _fill("main", "lightGrid 64 on 65", "lightGrid", "indexed:64", "indexed:65"),
    _fill("main", "lightGrid 65 on 64", "lightGrid", "indexed:65", "indexed:64"),
    _fill("main", "lightGrid auto on auto", "lightGrid", "auto", "auto"),
    _fill("main", "lightGrid 81 on yellow", "lightGrid", "indexed:81", f"rgb:{YELLOW}"),
    _fill("main", "lightGrid blue on 64", "lightGrid", f"rgb:{BLUE}", "indexed:64"),
    _fill("main", "lightGrid fg missing, bg yellow", "lightGrid", "missing", f"rgb:{YELLOW}"),
    _fill("main", "lightGrid blue, bg missing", "lightGrid", f"rgb:{BLUE}", "missing"),
    _fill("main", "lightGrid fg and bg missing", "lightGrid", "missing", "missing"),
    _fill("main", "lightGrid blue on yellow", "lightGrid", f"rgb:{BLUE}", f"rgb:{YELLOW}"),
    Case("main", "font auto", "font", color="auto"),
    Case("main", "font indexed 64", "font", color="indexed:64"),
    Case("main", "font indexed 81", "font", color="indexed:81"),
    Case("main", "font without color", "font", color="missing"),
    Case("main", "font blue", "font", color=f"rgb:{BLUE}"),
    Case("main", "font indexed 65 over gray", "font", color="indexed:65", under=GRAY),
    Case("main", "border auto", "border", color="auto"),
    Case("main", "border indexed 64", "border", color="indexed:64"),
    Case("main", "border without color", "border", color="missing"),
    Case("main", "border indexed 65", "border", color="indexed:65"),
    Case("main", "border indexed 81", "border", color="indexed:81"),
    Case("main", "border blue", "border", color=f"rgb:{BLUE}"),
    _fill("empty", "solid <fgColor/>", "solid", "empty"),
    _fill("empty", "solid <fgColor/>, bg yellow", "solid", "empty", f"rgb:{YELLOW}"),
    _fill("empty", "lightGrid blue on <bgColor/>", "lightGrid", f"rgb:{BLUE}", "empty"),
    _fill("empty", "lightGrid <fgColor/> on yellow", "lightGrid", "empty", f"rgb:{YELLOW}"),
    _fill("empty", "solid blue control", "solid", f"rgb:{BLUE}"),
    Case("empty", "font <color/>", "font", color="empty"),
    Case("empty", "border <color/>", "border", color="empty"),
]


def _place(specs: list[Case]) -> list[Case]:
    """Each case its cell: six to a row from row 4, every other row; each kind on fresh rows."""
    placed: list[Case] = []
    for group in GROUPS:
        row, slot, kind = 4, 0, None
        for case in (c for c in specs if c.group == group):
            if (kind is not None and case.kind != kind) or slot == len(SLOTS):
                row, slot = row + 2, 0
            kind = case.kind
            placed.append(dataclasses.replace(case, cell=f"{SLOTS[slot]}{row}"))
            slot += 1
    return placed


CASES = _place(_SPECS)


def cases(group: str | None = None) -> list[Case]:
    """Every case of ``group``, or of every group, in order."""
    return [c for c in CASES if group is None or c.group == group]


def _last_row(group: str) -> int:
    return max(int(c.cell[1:]) for c in cases(group))


MARKERS = {g: ("B2", f"O{_last_row(g) + 2}") for g in GROUPS}
"""The second marker right of the last case column and below the last case row."""

EXPECTED = {
    "solid fg indexed 64": "none",
    "solid fg auto": "none",
    "solid fg indexed 65": "solid FFFFFF",
    "solid fg indexed 81": "solid 000000",
    "solid fg missing, bg yellow": "solid 000000",
    "solid fg and bg missing": "solid 000000",
    "solid fg indexed 64, bg yellow": "none",
    "solid fg auto, bg yellow": "none",
    "solid blue": "solid 0000FF",
    "lightGrid 64 on 65": "lightGrid 000000/FFFFFF",
    "lightGrid 65 on 64": "lightGrid FFFFFF/FFFFFF",
    "lightGrid auto on auto": "lightGrid 000000/FFFFFF",
    "lightGrid 81 on yellow": "lightGrid 000000/FFFF00",
    "lightGrid blue on 64": "lightGrid 0000FF/FFFFFF",
    "lightGrid fg missing, bg yellow": "lightGrid 000000/FFFF00",
    "lightGrid blue, bg missing": "lightGrid 0000FF/FFFFFF",
    "lightGrid fg and bg missing": "lightGrid 000000/FFFFFF",
    "lightGrid blue on yellow": "lightGrid 0000FF/FFFF00",
    "font auto": "000000",
    "font indexed 64": "000000",
    "font indexed 81": "000000",
    "font without color": "000000",
    "font blue": "0000FF",
    "font indexed 65 over gray": "000000",
    "border auto": "medium 000000",
    "border indexed 64": "medium 000000",
    "border without color": "medium 000000",
    "border indexed 65": "medium 000000",
    "border indexed 81": "medium 000000",
    "border blue": "medium 0000FF",
    "solid <fgColor/>": "solid 000000",
    "solid <fgColor/>, bg yellow": "solid 000000",
    "lightGrid blue on <bgColor/>": "lightGrid 0000FF/FFFFFF",
    "lightGrid <fgColor/> on yellow": "lightGrid 000000/FFFF00",
    "solid blue control": "solid 0000FF",
    "font <color/>": "000000",
    "border <color/>": "medium 000000",
}
"""What sheetpeek draws for each case, by name: a fill as ``none``, ``solid RRGGBB`` or
``lightGrid FG/BG`` (the colors drawn, automatic resolved); a font's color drawn; a border's
style and color drawn on all four sides."""


# -- the workbooks -------------------------------------------------------------------------------


def placeholder(case: Case, which: str) -> str:
    """The ``rgb`` openpyxl writes for a color :func:`postprocess` removes or empties."""
    index = CASES.index(case)
    return f"FF0A{index:02X}{('fg', 'bg', 'color').index(which):02X}"


def _color(case: Case, which: str) -> Color | None:
    spec = getattr(case, which)
    if spec in ("missing", "empty"):
        return Color(rgb=placeholder(case, which))
    if spec == "auto":
        return Color(auto=True)
    kind, _, value = spec.partition(":")
    return Color(indexed=int(value)) if kind == "indexed" else Color(rgb="FF" + value)


def fill(ws: Any, group: str) -> list[Case]:
    """Lay one group out on a worksheet; its cases."""
    for k, column in enumerate(SLOTS):
        ws.column_dimensions[column].width = 12
        ws.column_dimensions[openpyxl.utils.get_column_letter(4 + 2 * k)].width = 2
    found = cases(group)
    for row in range(4, _last_row(group) + 1):
        ws.row_dimensions[row].height = 24 if row % 2 == 0 else 7
    for case in found:
        cell = ws[case.cell]
        if case.kind == "fill":
            cell.fill = PatternFill(
                case.pattern, fgColor=_color(case, "fg"), bgColor=_color(case, "bg")
            )
        elif case.kind == "font":
            cell.value = "HH"
            cell.font = Font(
                name="Calibri", size=11, family=2, scheme="minor", color=_color(case, "color")
            )
            if case.under is not None:
                cell.fill = PatternFill("solid", fgColor="FF" + case.under)
        else:
            side = Side("medium", color=_color(case, "color"))
            cell.border = Border(left=side, right=side, top=side, bottom=side)
    return found


def rewrite(text: str, group: str) -> str:
    """``styles.xml`` with each placeholder removed or emptied, and no ``<colors>``."""
    text = re.sub(r"<colors>.*?</colors>", "", text, flags=re.S)
    for case in cases(group):
        for which in ("fg", "bg") if case.kind == "fill" else ("color",):
            spec = getattr(case, which)
            if spec not in ("missing", "empty"):
                continue
            pattern = rf'<(fgColor|bgColor|color) rgb="{placeholder(case, which)}"\s*/>'
            text, count = re.subn(pattern, "" if spec == "missing" else r"<\1/>", text)
            if count == 0:
                raise ValueError(f"{case.name}: the placeholder for {which} was not written")
    return text


def postprocess(group: str) -> Any:
    """Rewrite a group's placeholder colors in ``xl/styles.xml``."""

    def edit(path: Path) -> None:
        from capture.screen.run import patch_part

        patch_part(path, {"xl/styles.xml": lambda t: rewrite(t, group)})

    return edit


def _cells(group: str) -> list[str]:
    return [c.cell for c in cases(group)]


def build(path: Path) -> list[Case]:
    """Write every group to a workbook of its own in the folder ``path``; every case."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        found.extend(fill(wb.active, group))
        target = path / f"auto_colors_{group}.xlsx"
        wb.save(target)
        postprocess(group)(target)
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
                cells=_cells(group),
                out=BUILD_DIR / group,
                markers=MARKERS[group],
                zooms=(ZOOM,),
                name=f"auto_colors_{group}",
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


def _hex(rgb: tuple[int, int, int]) -> str:
    return "".join(f"{c:02X}" for c in rgb)


def _distance(a: tuple[int, int, int], b: tuple[int, int, int]) -> float:
    return float(sum((x - y) ** 2 for x, y in zip(a, b, strict=True)) ** 0.5)


def interior(cap: Any, cell: str) -> str:
    """The two most common colors inside the cell (``INSET`` px in), ``RRGGBB:share`` each."""
    from capture.screen.ink import cells_box

    x0, y0, x1, y1 = cells_box(cap, cell, inset=INSET)
    counts = Counter(cap.image.pixel(x, y) for y in range(y0, y1) for x in range(x0, x1))
    total = sum(counts.values())
    return " ".join(
        f"{_hex(rgb)}:{round(n / total, 4):g}" for rgb, n in counts.most_common(2) if total
    )


def gridlines(cap: Any, cell: str) -> str:
    """Per edge (left, right, top, bottom), ``1`` where the gridline's gray shows at its middle."""
    gray = cap.gridline
    if gray is None:
        return ""
    cx, cy = round(cap.edge(cell, "cx")), round(cap.edge(cell, "cy"))
    flags = []
    for edge in EDGES:
        at = round(cap.edge(cell, edge))
        if edge in ("left", "right"):
            pixels = [cap.image.pixel(x, cy) for x in range(at - REACH, at + REACH)]
        else:
            pixels = [cap.image.pixel(cx, y) for y in range(at - REACH, at + REACH)]
        flags.append("1" if any(_distance(p, gray) <= 8 for p in pixels) else "0")
    return "".join(flags)


def text_ink(cap: Any, cell: str) -> dict[str, str]:
    """The text's color: the mean of the tenth of pixels farthest from the cell's own color."""
    from capture.screen.ink import cells_box

    x0, y0, x1, y1 = cells_box(cap, cell, inset=INSET)
    pixels = [cap.image.pixel(x, y) for y in range(y0, y1) for x in range(x0, x1)]
    paper = Counter(pixels).most_common(1)[0][0]
    ink = sorted(
        (p for p in pixels if _distance(p, paper) > 60), key=lambda p: -_distance(p, paper)
    )
    if not ink:
        return {"ink_color": "none", "ink_pixels": "0"}
    core = ink[: max(1, len(ink) // 10)]
    mean = tuple(round(sum(p[i] for p in core) / len(core)) for i in range(3))
    return {"ink_color": _hex(mean), "ink_pixels": str(len(ink))}  # type: ignore[arg-type]


def measure(cap: Any, case: Case) -> dict[str, str]:
    """One case's measurements at one zoom."""
    from capture.tables import _strip, runs

    found = {"interior": interior(cap, case.cell), "gridlines": gridlines(cap, case.cell)}
    if case.kind == "font":
        found |= text_ink(cap, case.cell)
    if case.kind == "border":
        for edge in EDGES:
            found[f"edge_{edge}"] = runs(_strip(cap, case.cell, edge))
    return found


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``auto_colors.tsv``."""
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
                reads = {key: cap.value(case.cell, read) or "" for key, read in READS.items()}
                records.append(
                    {
                        "group": group,
                        "name": case.name,
                        "kind": case.kind,
                        "cell": case.cell,
                        "zoom": str(cap.zoom),
                        **measure(cap, case),
                        **reads,
                    }
                )
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: auto_colors\n" + "\t".join(COLUMNS) + "\n")
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
