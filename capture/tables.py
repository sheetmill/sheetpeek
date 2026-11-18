"""Capture how Excel for Mac draws table styles, from the screen and its display format (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.tables run unknown
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.screen run tables
    uv run --group capture python -m capture.screen read tables     # the saved captures again
    uv run --group capture python -m capture.tables build PATH      # every group's workbook

Each group is its own workbook and window, at zoom 100; the second marker is placed below and
right of the group's tables. Tables carry no filter (no ``autoFilter``), so no button covers a
header; their column names equal their header cells and a totals row holds only its label
(``totalsRowLabel``), as Excel requires. Groups, in the order ``run`` takes them:

- ``unknown``: the probe. Table ``Nowhere`` names a style that is neither a preset nor the
  workbook's own, beside table ``Control`` in TableStyleMedium2. Run it alone first: if Excel
  refuses the file (-50 with alerts off) the run stops there, which is the answer.
- ``presets_1`` to ``presets_8``: Excel's 60 presets (Light1 to 21, Medium1 to 28, Dark1 to 11),
  eight to a group, two across and four down, each a 5 by 6 table (header, four data rows,
  totals) with all four options on, under openpyxl's default theme (Office 2007).
- ``theme``: eight presets, one from each family and every accent, under a second theme whose
  ``dk1``, ``lt1``, ``dk2``, ``lt2`` and accents are all changed.
- ``order``: table ``Order``, a workbook style whose 13 elements each draw their own fill, font
  color and outer sides, lighter along the order; table ``Weights``, thin later lines over a
  thick earlier one (the row stripe's top, the header's bottom, the totals' hair top).
- ``bands``: row stripe sizes (1,1), (2,1), (1,2), (3,2), a first stripe of 3 with no second
  element; column stripe sizes (2,1) and (1,3); row stripes in a table without a header row.
- ``hidden``: sheet row 7 hidden inside a table's row stripes, and sheet column E inside another
  table's column stripes (each hides it for the whole sheet, so the tables share neither).
- ``flags``: TableStyleMedium2 with no option, each option alone, and without a header row.
- ``stacking``: TableStyleMedium2 (white bold header on accent 1): in the header row and in a
  data row, a font with no color, ``auto``, ``theme="1"``, ``theme="1"`` at tint 0.5, black
  ``rgb``, ``indexed="8"``, ``indexed="64"`` and red ``rgb``; own fills (red, white, ``gray125``);
  Arial; own bold and italic; own underline and strikethrough; own thin red sides over the
  table's inner line and over the totals' double top. Table ``Lines`` in a workbook style that
  draws underline and strikethrough from its whole-table element.
- ``normal_<variant>``: the same eight font colors as the Normal style's font, each its own
  workbook, a TableStyleMedium2 table whose cells set no font.
- ``merge``: a merge inside a table's data rows. Excel may refuse or repair the file; that is
  the result, and the run stops there.

Measured for every table cell: Excel's display format (fill, font color, bold, italic,
underline, strikethrough, font name; exact, from ``display format``), and across every edge of
the cell, at the segment's middle, the colors of the device pixels from 3 screen points before
the boundary to 3 after it, averaged over 5 pixels along the segment, as hex, so that each edge
can be classified afterwards (none, gridline, thin, medium, thick, double; their colors).
``read`` writes ``tests/captures/tables.tsv``, one row per group, cell and zoom.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Border, Color, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.table import TableStyle, TableStyleElement
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

PRESETS = (
    *(f"TableStyleLight{n}" for n in range(1, 22)),
    *(f"TableStyleMedium{n}" for n in range(1, 29)),
    *(f"TableStyleDark{n}" for n in range(1, 12)),
)
THEME_SAMPLE = (
    "TableStyleLight1", "TableStyleLight9", "TableStyleLight17", "TableStyleMedium3",
    "TableStyleMedium11", "TableStyleMedium20", "TableStyleDark5", "TableStyleDark11",
)  # fmt: skip
"""One preset of each kind, accents 1 to 6 and dark 1 among them."""
ELEMENTS = (
    "wholeTable", "firstColumnStripe", "secondColumnStripe", "firstRowStripe", "secondRowStripe",
    "lastColumn", "firstColumn", "headerRow", "totalRow", "firstHeaderCell", "lastHeaderCell",
    "firstTotalCell", "lastTotalCell",
)  # fmt: skip
FONT_COLORS = {
    "none": None,
    "auto": Color(auto=True),
    "theme1": Color(theme=1),
    "theme1_tint": Color(theme=1, tint=0.5),
    "black": Color(rgb="FF000000"),
    "indexed8": Color(indexed=8),
    "indexed64": Color(indexed=64),
    "red": Color(rgb="FFFF0000"),
}
SECOND_THEME = {
    "dk1": "1F3864", "lt1": "FFF2CC", "dk2": "7030A0", "lt2": "E2EFDA",
    "accent1": "C00000", "accent2": "00B050", "accent3": "7030A0",
    "accent4": "FFC000", "accent5": "00B0F0", "accent6": "FF00FF",
}  # fmt: skip
GROUPS = (
    "unknown",
    *(f"presets_{k}" for k in range(1, 9)),
    "theme",
    "order",
    "bands",
    "hidden",
    "flags",
    "stacking",
    *(f"normal_{name}" for name in FONT_COLORS),
    "merge",
)
REFUSALS = ("unknown", "merge")
"""Groups whose file Excel may refuse: run on their own, a refusal ends the run."""
ZOOM = 100
STRIP = 3
"""Screen points measured on each side of an edge."""
ALONG = 5
"""Device pixels averaged along an edge."""
ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "tables"
OUTPUT = ROOT / "tests" / "captures" / "tables.tsv"
READS = {
    "fill": "display.fill",
    "pattern": "display.pattern",
    "font_color": "display.font_color",
    "bold": "display.bold",
    "italic": "display.italic",
    "underline": "display.underline",
    "strike": "display.strike",
    "font_name": "display.font_name",
}
EDGES = ("top", "bottom", "left", "right")
COLUMNS = ("group", "table", "cell", "role", "zoom", *READS, *(f"edge_{e}" for e in EDGES))


@dataclass(frozen=True)
class Placed:
    """One table cell to measure: its group, table, address and what it is there for."""

    group: str
    table: str
    cell: str
    role: str


@dataclass
class Group:
    """A group's tables, as laid out, and the second marker below and right of them."""

    cells: list[Placed]
    marker: str = "B2"


def _letter(column: int) -> str:
    return openpyxl.utils.get_column_letter(column)


def _add(
    ws: Any,
    found: list[Placed],
    group: str,
    name: str,
    top: int,
    left: int,
    style: str | None,
    *,
    columns: int = 5,
    data: int = 4,
    header: bool = True,
    totals: bool = True,
    flags: tuple[bool, bool, bool, bool] = (True, True, True, True),
    roles: dict[str, str] | None = None,
) -> Table:
    """A table at (top, left): headers equal to its column names, numbers, a labelled totals row.

    ``flags`` are first column, last column, row stripes, column stripes.
    """
    names = [f"{name[:3]}{k + 1}" for k in range(columns)]
    row = top
    if header:
        for k, label in enumerate(names):
            ws.cell(row, left + k, label)
        row += 1
    for r in range(data):
        for k in range(columns):
            ws.cell(row + r, left + k, (r + 1) * 10 + k)
    last = row + data - 1 + (1 if totals else 0)
    if totals:
        ws.cell(last, left, "Total")
    ref = f"{_letter(left)}{top}:{_letter(left + columns - 1)}{last}"
    table = Table(displayName=name, ref=ref, headerRowCount=1 if header else 0)
    table.tableColumns = [TableColumn(id=k + 1, name=label) for k, label in enumerate(names)]
    if totals:
        table.totalsRowCount = 1
        table.tableColumns[0].totalsRowLabel = "Total"
    if style is not None:
        first, last_column, rows, cols = flags
        table.tableStyleInfo = TableStyleInfo(
            name=style,
            showFirstColumn=first,
            showLastColumn=last_column,
            showRowStripes=rows,
            showColumnStripes=cols,
        )
    ws.add_table(table)
    for r in range(top, last + 1):
        for c in range(left, left + columns):
            cell = f"{_letter(c)}{r}"
            found.append(Placed(group, name, cell, (roles or {}).get(cell, "")))
    return table


def _style(wb: Any, name: str, elements: list[tuple[str, DifferentialStyle, int]]) -> None:
    """Add a workbook table style of (element type, format, size) elements."""
    found = [
        TableStyleElement(type=kind, dxfId=wb._differential_styles.add(dxf), size=size)
        for kind, dxf, size in elements
    ]
    wb._table_styles.tableStyle.append(TableStyle(name=name, pivot=False, tableStyleElement=found))


def _fill(rgb: str) -> PatternFill:
    """A differential format's solid fill: drawn in its ``bgColor``."""
    return PatternFill(patternType="solid", bgColor="FF" + rgb)


def _side(style: str, rgb: str) -> Side:
    return Side(style=style, color="FF" + rgb)


def _presets(ws: Any, group: str, names: tuple[str, ...]) -> Group:
    found: list[Placed] = []
    for k, style in enumerate(names):
        top, left = 4 + 8 * (k // 2), 3 + 6 * (k % 2)
        _add(ws, found, group, style.replace("TableStyle", "T"), top, left, style)
    return Group(found, "O36")


def _order(ws: Any) -> Group:
    wb = ws.parent
    found: list[Placed] = []
    # Each element a fill and font color of its own, and outer sides that get lighter along the
    # order, so that a later element's line is told from an earlier one's by its gray.
    elements = []
    for k, kind in enumerate(ELEMENTS):
        gray = f"{16 * k:02X}" * 3
        fill = f"{40 + 15 * k:02X}{200 - 10 * k:02X}{(70 * k) % 256:02X}"
        font = f"{(255 - 19 * k) % 256:02X}{(30 * k) % 256:02X}{(90 + 12 * k) % 256:02X}"
        side = _side("thin", gray)
        dxf = DifferentialStyle(
            font=Font(color="FF" + font),
            fill=_fill(fill),
            border=Border(left=side, right=side, top=side, bottom=side),
        )
        elements.append((kind, dxf, 1))
    _style(wb, "Order", elements)
    _add(ws, found, "order", "Order", 4, 3, "Order", columns=6, data=4)
    thick, red, blue = _side("thick", "000000"), _side("thin", "FF0000"), _side("thin", "0000FF")
    _style(
        wb,
        "Weights",
        [
            ("wholeTable", DifferentialStyle(border=Border(horizontal=thick, vertical=thick)), 1),
            ("firstRowStripe", DifferentialStyle(border=Border(top=red)), 1),
            ("headerRow", DifferentialStyle(border=Border(bottom=blue)), 1),
            ("totalRow", DifferentialStyle(border=Border(top=_side("hair", "0000FF"))), 1),
        ],
    )
    _add(ws, found, "order", "Weights", 4, 11, "Weights", columns=4, data=4)
    return Group(found, "P13")


def _bands(ws: Any) -> Group:
    wb = ws.parent
    found: list[Placed] = []
    a, b = _fill("9BC2E6"), _fill("F4B084")
    # Row stripes: (first size, second size or None for no second element).
    for k, (first, second) in enumerate(((1, 1), (2, 1), (1, 2), (3, 2), (3, None))):
        elements = [("firstRowStripe", DifferentialStyle(fill=a), first)]
        if second is not None:
            elements.append(("secondRowStripe", DifferentialStyle(fill=b), second))
        name = f"Rows{first}{second or 0}"
        _style(wb, name, elements)
        _add(ws, found, "bands", name, 4, 3 + 3 * k, name, columns=2, data=9, totals=False,
             flags=(False, False, True, False))  # fmt: skip
    for k, (first, second) in enumerate(((2, 1), (1, 3))):
        name = f"Cols{first}{second}"
        _style(wb, name, [("firstColumnStripe", DifferentialStyle(fill=a), first),
                          ("secondColumnStripe", DifferentialStyle(fill=b), second)])  # fmt: skip
        _add(ws, found, "bands", name, 16, 3 + 8 * k, name, columns=7, data=3, totals=False,
             flags=(False, False, False, True))  # fmt: skip
    _add(ws, found, "bands", "NoHeader", 22, 3, "Rows11", columns=3, data=5, header=False,
         totals=False, flags=(False, False, True, False))  # fmt: skip
    return Group(found, "S28")


def _hidden(ws: Any) -> Group:
    """A hidden row inside row stripes and a hidden column inside column stripes.

    A hidden row or column is hidden for the whole sheet, so each table keeps rows and columns
    no other table uses; the hidden cells themselves are not measured.
    """
    wb = ws.parent
    found: list[Placed] = []
    rows, columns = DifferentialStyle(fill=_fill("9BC2E6")), DifferentialStyle(fill=_fill("F4B084"))
    _style(wb, "Hidden", [("firstRowStripe", rows, 1), ("firstColumnStripe", columns, 1)])
    _add(ws, found, "hidden", "HiddenRow", 4, 10, "Hidden", columns=3, data=6, totals=False,
         flags=(False, False, True, False))  # fmt: skip
    ws.row_dimensions[7].hidden = True
    _add(ws, found, "hidden", "HiddenCol", 12, 3, "Hidden", columns=6, data=3, totals=False,
         flags=(False, False, False, True))  # fmt: skip
    ws.column_dimensions["E"].hidden = True
    shown = [
        p for p in found if not (p.cell[0] == "E" or (p.table == "HiddenRow" and p.cell[1:] == "7"))
    ]
    return Group(shown, "N17")


def _flags(ws: Any) -> Group:
    found: list[Placed] = []
    cases = {
        "None": (False, False, False, False),
        "First": (True, False, False, False),
        "Last": (False, True, False, False),
        "Rows": (False, False, True, False),
        "Cols": (False, False, False, True),
    }
    for k, (name, flags) in enumerate(cases.items()):
        _add(ws, found, "flags", f"Flag{name}", 4 + 8 * (k // 3), 3 + 6 * (k % 3),
             "TableStyleMedium2", flags=flags)  # fmt: skip
    _add(ws, found, "flags", "FlagNoHeader", 12, 15, "TableStyleMedium2", header=False)
    return Group(found, "U21")


def _stacking(ws: Any) -> Group:
    wb = ws.parent
    found: list[Placed] = []
    columns = len(FONT_COLORS)
    roles: dict[str, str] = {}
    for k, (name, color) in enumerate(FONT_COLORS.items()):
        column = _letter(3 + k)
        for row in (4, 6):  # the header, and a data row with no row stripe
            cell = ws[f"{column}{row}"]
            if color is not None:
                cell.font = Font(name="Calibri", size=11, family=2, scheme="minor", color=color)
            else:
                cell.font = Font(name="Calibri", size=11, family=2, scheme="minor")
            roles[f"{column}{row}"] = f"font color {name}"
    table = _add(ws, found, "stacking", "Colors", 4, 3, "TableStyleMedium2", columns=columns,
                 data=4, roles=roles, flags=(False, False, True, False))  # fmt: skip
    del table
    own = {
        "C13": ("fill red", {"fill": PatternFill("solid", fgColor="FFFF0000")}),
        "D13": ("fill white", {"fill": PatternFill("solid", fgColor="FFFFFFFF")}),
        "E13": ("fill gray125", {"fill": PatternFill("gray125")}),
        "F13": ("font Arial", {"font": Font(name="Arial", size=11)}),
        "C14": ("own bold", {"font": Font(name="Calibri", size=11, scheme="minor", bold=True)}),
        "D14": ("own italic", {"font": Font(name="Calibri", size=11, scheme="minor", italic=True)}),
        "E14": ("own underline", {"font": Font(name="Calibri", size=11, scheme="minor",
                                               underline="double")}),
        "F14": ("own strike", {"font": Font(name="Calibri", size=11, scheme="minor",
                                            strike=True)}),
        "D15": ("own bottom over inner line", {"border": Border(bottom=_side("thin", "FF0000"))}),
        "E17": ("own top over totals double", {"border": Border(top=_side("thin", "FF0000"))}),
        "F16": ("own bottom against totals double",
                {"border": Border(bottom=_side("thin", "FF0000"))}),
    }  # fmt: skip
    for cell, (_, attributes) in own.items():
        for key, value in attributes.items():
            setattr(ws[cell], key, value)
    _add(ws, found, "stacking", "Own", 12, 3, "TableStyleMedium2", columns=4, data=4,
         roles={cell: role for cell, (role, _) in own.items()},
         flags=(False, False, False, False))  # fmt: skip
    _style(
        wb,
        "Lines",
        [("wholeTable", DifferentialStyle(font=Font(underline="single", strike=True)), 1)],
    )
    ws["J14"].font = Font(name="Calibri", size=11, scheme="minor", underline="double")
    _add(ws, found, "stacking", "Lines", 12, 9, "Lines", columns=3, data=3, totals=False,
         roles={"J14": "own double under the table's single"})  # fmt: skip
    return Group(found, "M19")


def _normal(ws: Any, name: str) -> Group:
    found: list[Placed] = []
    _add(ws, found, f"normal_{name}", "Normal", 4, 3, "TableStyleMedium2", columns=3, data=2)
    return Group(found, "H10")


def normal_color(name: str) -> Any:
    """A postprocess that gives the Normal style's font (the first ``<font>``) that color."""
    import re

    from capture.screen.run import patch_part

    color = FONT_COLORS[name]
    if color is None:
        element = ""
    else:
        value = "1" if color.type == "auto" else color.value
        tint = f' tint="{color.tint!r}"' if color.tint else ""
        element = f'<color {color.type}="{value}"{tint}/>'

    def recolor(text: str) -> str:
        start = text.index("<font>")
        end = text.index("</font>", start)
        font = re.sub(r"<color[^>]*/>", "", text[start:end])
        return text[:start] + font.replace("<font>", "<font>" + element, 1) + text[end:]

    return lambda path: patch_part(path, {"xl/styles.xml": recolor})


def _unknown(ws: Any) -> Group:
    found: list[Placed] = []
    _add(ws, found, "unknown", "Nowhere", 4, 3, "Nowhere")
    _add(ws, found, "unknown", "Control", 4, 9, "TableStyleMedium2")
    return Group(found, "O12")


def _merge(ws: Any) -> Group:
    found: list[Placed] = []
    _add(ws, found, "merge", "Merged", 4, 3, "TableStyleMedium2")
    ws.merge_cells("D6:E6")
    return Group(found, "I12")


def fill(ws: Any, group: str) -> Group:
    """Lay one group out on a worksheet."""
    if group.startswith("presets_"):
        k = int(group.split("_")[1]) - 1
        return _presets(ws, group, PRESETS[8 * k : 8 * k + 8])
    if group == "theme":
        return _presets(ws, group, THEME_SAMPLE)
    if group.startswith("normal_"):
        return _normal(ws, group.split("_", 1)[1])
    return {
        "unknown": _unknown,
        "order": _order,
        "bands": _bands,
        "hidden": _hidden,
        "flags": _flags,
        "stacking": _stacking,
        "merge": _merge,
    }[group](ws)


def second_theme(path: Path) -> None:
    """Change the saved workbook's theme colors to :data:`SECOND_THEME`."""
    import re

    from capture.screen.run import patch_part

    def recolor(text: str) -> str:
        for name, rgb in SECOND_THEME.items():
            text, count = re.subn(
                rf"(<a:{name}>\s*<a:(?:srgbClr val|sysClr val=\"\w+\" lastClr)=\")[0-9A-Fa-f]{{6}}",
                rf"\g<1>{rgb}",
                text,
            )
            if count != 1:
                raise SystemExit(f"theme color {name} not found once in the theme")
        return text

    patch_part(path, {"xl/theme/theme1.xml": recolor})


def build(path: Path) -> list[Placed]:
    """Write every group to a workbook of its own in the folder ``path``; every cell."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        found.extend(fill(wb.active, group).cells)
        target = path / f"tables_{group}.xlsx"
        wb.save(target)
        post = postprocess(group)
        if post is not None:
            post(target)
    return found


def postprocess(group: str) -> Any:
    """What a group's saved workbook needs that openpyxl does not write, or ``None``."""
    if group == "theme":
        return second_theme
    if group.startswith("normal_"):
        return normal_color(group.split("_", 1)[1])
    return None


def run(groups: tuple[str, ...] | None = None) -> None:
    """Capture each group's window, then write the TSV (needs Excel).

    Without groups, every group but :data:`REFUSALS`, which run on their own first.
    """
    from capture.screen.run import capture

    if groups is None:
        groups = tuple(g for g in GROUPS if g not in REFUSALS)
    for group in groups:
        laid = fill(openpyxl.Workbook().active, group)
        capture(
            lambda ws, group=group: fill(ws, group),
            cells=[p.cell for p in laid.cells],
            out=BUILD_DIR / group,
            markers=("B2", laid.marker),
            zooms=(ZOOM,),
            name=f"tables_{group}",
            postprocess=postprocess(group),
        )
    write()


def _strip(cap: Any, cell: str, edge: str) -> str:
    """The colors across one edge at its middle, from STRIP points before it to STRIP after."""
    per = cap.scale * 100 / cap.zoom  # image pixels a screen point
    left, right = cap.edge(cell, "left"), cap.edge(cell, "right")
    top, bottom = cap.edge(cell, "top"), cap.edge(cell, "bottom")
    reach = round(STRIP * per)
    found = []
    if edge in ("top", "bottom"):
        y = round(top if edge == "top" else bottom)
        x = round((left + right) / 2)
        for dy in range(-reach, reach):
            pixels = [
                cap.image.pixel(x + dx, y + dy) for dx in range(-(ALONG // 2), ALONG // 2 + 1)
            ]
            found.append(_mean(pixels))
    else:
        x = round(left if edge == "left" else right)
        y = round((top + bottom) / 2)
        for dx in range(-reach, reach):
            pixels = [
                cap.image.pixel(x + dx, y + dy) for dy in range(-(ALONG // 2), ALONG // 2 + 1)
            ]
            found.append(_mean(pixels))
    return " ".join(found)


def _mean(pixels: list[tuple[int, int, int]]) -> str:
    return "".join(f"{round(sum(p[i] for p in pixels) / len(pixels)):02X}" for i in range(3))


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``tables.tsv``."""
    from capture.screen.run import header, load

    records = []
    first: dict[str, str] = {}
    for group in GROUPS:
        out = BUILD_DIR / group
        if not (out / "formats.tsv").exists():
            print(f"{group}: not captured yet, left out", file=sys.stderr)
            continue
        first = first or header(out)
        placed = fill(openpyxl.Workbook().active, group).cells
        for cap in load(out):
            for p in placed:
                record = {
                    "group": p.group,
                    "table": p.table,
                    "cell": p.cell,
                    "role": p.role,
                    "zoom": str(cap.zoom),
                }
                record |= {key: cap.value(p.cell, read) or "" for key, read in READS.items()}
                record |= {f"edge_{e}": _strip(cap, p.cell, e) for e in EDGES}
                records.append(record)
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after", "zooms"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: tables\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cells in {argv[1]}")
        return 0
    if argv[:1] == ["run"] and argv[1:] and all(group in GROUPS for group in argv[1:]):
        run(tuple(argv[1:]))
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
