"""Capture more of how Excel for Mac draws filter and sort buttons, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.filters_more run beside
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.filters_more run
    uv run --group capture python -m capture.filters_more read          # the saved captures again
    uv run --group capture python -m capture.filters_more build PATH    # every group's workbook

ARCHITECTURE.md §4.19, the rules the suite ``filters`` left unverified: the criteria are not
applied on opening (the rows the file saved as hidden are what shows); a button on each column
of the range's first row, merged across or down and empty header cells included, none on a
hidden column or header row; a left-to-right sort (``columnSort``) draws no arrow, on the sheet
and in a table; the sheet's ``sortState`` beside its ``autoFilter`` (not inside it) counts.

Each group is its own workbook and window, at zoom 100 and 200 as ``filters``, in the dark
appearance as the other suites; case columns are 12 characters wide, rows 15 pt. Groups:

- ``main``: the sheet's filter C4:I9. C4:D4 is one header merged across, E4:E5 one merged down
  into the data, F4 is empty, column G is hidden, H filters on the values 1, 3 and 5 (H5:H9
  hold 1 to 5, so the criteria would hide rows 6 and 8) while the file hides row 7 only (and
  sets ``filterMode``, as Excel does when a filter hides rows), and the sheet's ``sortState``,
  inside its ``autoFilter``, is left to right (``columnSort="1"``) with its condition on I.
  Table ``ColSort`` (L4:N6) has a left-to-right ``sortState`` with its condition on M; table
  ``Hidden`` (L12:N14) has its header row 12 hidden.
- ``beside``: the sheet's filter C4:E6 with no ``sortState`` inside its ``autoFilter`` and one
  beside it, the worksheet's own child after it, sorting D descending (written into
  ``xl/worksheets/sheet1.xml`` before Excel opens the file, ``postprocess``). Excel may refuse
  or repair the file (-50 with alerts off): it runs on its own, first, and a refusal is kept
  (``refused.txt``) and written as the group's one row, its ``case`` ``refused:`` and Excel's
  message.

Cell boxes come from reference cells in row 1 and column A, so no merged cell's geometry is
read. Measured, as ``filters`` measures (``capture.filters._measure``), for each case cell: the
button's box, its gaps to the cell's right and bottom boundaries and its top below the row's
top, and the glyph inside it as a bitmap id, equal ids meaning identical glyphs, so they compare
with ``tests/captures/filters.tsv``'s; E5 (under the header merged down) and L13, M13 (below
the hidden header row) are measured for a button sheetpeek does not draw. Each case row's height
and column G's width, as Excel reports them (``height`` and ``width`` of the reference cell; 0 when
hidden), are written once per group in ``size``: they show which rows the criteria left hidden.
``read`` writes ``tests/captures/filters_more.tsv``.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.worksheet import filters as xl_filters
from openpyxl.worksheet.table import Table, TableColumn

from capture import filters


@dataclass(frozen=True)
class Case:
    group: str
    cell: str
    name: str


CASES = (
    Case("main", "C4", "merged across, first column"),
    Case("main", "D4", "merged across, second column"),
    Case("main", "E4", "merged down, header row"),
    Case("main", "E5", "merged down, data row"),
    Case("main", "F4", "empty header"),
    Case("main", "H4", "values filter, criteria not applied"),
    Case("main", "I4", "sheet columnSort condition"),
    Case("main", "L4", "table columnSort, plain"),
    Case("main", "M4", "table columnSort condition"),
    Case("main", "N4", "table columnSort, plain"),
    Case("main", "L13", "below a hidden header row"),
    Case("main", "M13", "below a hidden header row"),
    Case("beside", "C4", "plain, sortState beside"),
    Case("beside", "D4", "sorted descending, sortState beside"),
    Case("beside", "E4", "plain, sortState beside"),
)
EXPECTED: dict[tuple[str, str], str | None] = {
    ("main", "C4"): "plain", ("main", "D4"): "plain", ("main", "E4"): "plain",
    ("main", "E5"): None, ("main", "F4"): "plain", ("main", "H4"): "filtered",
    ("main", "I4"): "plain", ("main", "L4"): "plain", ("main", "M4"): "plain",
    ("main", "N4"): "plain", ("main", "L13"): None, ("main", "M13"): None,
    ("beside", "C4"): "plain", ("beside", "D4"): "descending", ("beside", "E4"): "plain",
}  # fmt: skip
"""sheetpeek's glyph for each case cell; ``None`` draws no button."""
SIZES: dict[str, dict[str, float]] = {
    "main": {
        "A4": 15, "A5": 15, "A6": 15, "A7": 0, "A8": 15, "A9": 15,
        "A12": 0, "A13": 15, "A14": 15, "G1": 0,
    },
    "beside": {"A4": 15, "A5": 15, "A6": 15},
}  # fmt: skip
"""sheetpeek's row heights (``A<row>``) and column G's width (``G1``) in points, 0 hidden."""
GROUPS = ("main", "beside")
REFUSALS = ("beside",)
"""Groups whose file Excel may refuse: run on their own, a refusal is recorded as the answer."""
MARKERS = {"main": ("B2", "P16"), "beside": ("B2", "H9")}
ZOOMS = (100, 200)
WIDTH = 12
ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "filters_more"
OUTPUT = ROOT / "tests" / "captures" / "filters_more.tsv"
COLUMNS = ("group", "zoom", "cell", "case", "size", *filters.COLUMNS[3:])
BESIDE = '<sortState ref="C5:E6"><sortCondition descending="1" ref="D5:D6"/></sortState>'


def cases(group: str) -> list[Case]:
    return [case for case in CASES if case.group == group]


def _columns(names: list[str]) -> list[TableColumn]:
    """A table's columns, named as their header cells: Excel refuses a table whose names differ."""
    return [TableColumn(id=k + 1, name=name) for k, name in enumerate(names)]


def _table(ws: Any, name: str, top: int, left: int, headers: list[str]) -> Table:
    for k, header in enumerate(headers):
        ws.cell(top, left + k, header)
        for r in (1, 2):
            ws.cell(top + r, left + k, r * (k + 1))
    last = openpyxl.utils.get_column_letter(left + len(headers) - 1)
    ref = f"{openpyxl.utils.get_column_letter(left)}{top}:{last}{top + 2}"
    table = Table(displayName=name, ref=ref)
    table.tableColumns = _columns(headers)
    table.autoFilter = xl_filters.AutoFilter(ref=ref)
    return table


def _main(ws: Any) -> None:
    for column in "CDEFGHILMN":
        ws.column_dimensions[column].width = WIDTH
    for k, header in enumerate(["across", None, "down", None, "hidden", "values", "colsort"]):
        if header is not None:
            ws.cell(4, 3 + k, header)
        for row in range(5, 10):
            if (row, k) != (5, 2):  # E5 lies under the header merged down
                ws.cell(row, 3 + k, row - 4 if k == 5 else (row - 4) * (k + 1))
    ws.merge_cells("C4:D4")
    ws.merge_cells("E4:E5")
    ws.column_dimensions["G"].hidden = True
    ws.row_dimensions[7].hidden = True
    ws.sheet_properties.filterMode = True
    ws.auto_filter.ref = "C4:I9"
    ws.auto_filter.filterColumn.append(
        xl_filters.FilterColumn(colId=5, filters=xl_filters.Filters(filter=["1", "3", "5"]))
    )
    ws.auto_filter.sortState = xl_filters.SortState(
        ref="C5:I9", columnSort=True, sortCondition=[xl_filters.SortCondition(ref="I5:I9")]
    )
    colsort = _table(ws, "ColSort", 4, 12, ["left", "rows", "right"])
    colsort.sortState = xl_filters.SortState(
        ref="L5:N6", columnSort=True, sortCondition=[xl_filters.SortCondition(ref="M5:M6")]
    )
    ws.add_table(colsort)
    ws.add_table(_table(ws, "Hidden", 12, 12, ["one", "two", "three"]))
    ws.row_dimensions[12].hidden = True


def _beside(ws: Any) -> None:
    for column in "CDE":
        ws.column_dimensions[column].width = WIDTH
    for k, header in enumerate(["plain", "down", "plain2"]):
        ws.cell(4, 3 + k, header)
        for r in (1, 2):
            ws.cell(4 + r, 3 + k, r * (k + 1))
    ws.auto_filter.ref = "C4:E6"


def fill(ws: Any, group: str) -> list[Case]:
    """Lay one group out on a worksheet; its cases."""
    {"main": _main, "beside": _beside}[group](ws)
    return cases(group)


def sort_beside(text: str) -> str:
    """The worksheet XML with :data:`BESIDE` as the ``autoFilter``'s next sibling."""
    found, count = re.subn(r"(<autoFilter\b[^>]*?/>)", rf"\g<1>{BESIDE}", text)
    if count != 1:
        raise SystemExit("the sheet's autoFilter was not found once, or holds children")
    return found


def postprocess(group: str) -> Any:
    """What a group's saved workbook needs that openpyxl does not write, or ``None``."""
    if group != "beside":
        return None
    from capture.screen.run import patch_part

    return lambda path: patch_part(path, {"xl/worksheets/sheet1.xml": sort_beside})


def build(path: Path) -> list[Case]:
    """Write every group to a workbook of its own in the folder ``path``; every case."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        found.extend(fill(wb.active, group))
        target = path / f"filters_more_{group}.xlsx"
        wb.save(target)
        post = postprocess(group)
        if post is not None:
            post(target)
    return found


def _references(group: str) -> list[str]:
    """The cells whose geometry gives every case cell's box: row 1 and column A."""
    found: dict[str, None] = dict.fromkeys(SIZES[group])
    for case in cases(group):
        letters, row = coordinate_from_string(case.cell)
        found[f"{letters}1"] = None
        found[f"A{row}"] = None
    return list(found)


def run(groups: tuple[str, ...] | None = None) -> None:
    """Capture each group's window at both zooms, then write the TSV (needs Excel).

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
                cells=_references(group),
                out=BUILD_DIR / group,
                markers=MARKERS[group],
                zooms=ZOOMS,
                name=f"filters_more_{group}",
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


def place(cap: Any, cell: str) -> None:
    """Give ``cell`` its box from its column's row-1 cell and its row's column-A cell."""
    letters, row = coordinate_from_string(cell)
    column, line = cap.geometry[f"{letters}1"], cap.geometry[f"A{row}"]
    cap.geometry[cell] = {
        "left": column["left"],
        "width": column["width"],
        "top": line["top"],
        "height": line["height"],
    }


def measure(cap: Any, group: str) -> list[dict[str, str]]:
    """Every case cell's record at one zoom (``capture.filters``' measuring)."""
    found = []
    for case in cases(group):
        place(cap, case.cell)
        found.append({"group": group} | filters._measure(cap, case.cell, case.name))
    return found


def sizes(cap: Any, group: str) -> list[dict[str, str]]:
    """Each reference row's height and column G's width as Excel reports them, in points."""
    found = []
    for cell in SIZES[group]:
        key = "width" if cell.endswith("1") and not cell.startswith("A") else "height"
        value = cap.geometry.get(cell, {}).get(key)
        found.append(
            {
                "group": group,
                "cell": cell,
                "case": f"{'column' if key == 'width' else 'row'} {key}",
                "size": "" if value is None else f"{value:g}",
            }
        )
    return found


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``filters_more.tsv``."""
    from capture.screen.run import header, load

    records: list[dict[str, str]] = []
    first: dict[str, str] = {}
    for group in GROUPS:
        out = BUILD_DIR / group
        refused = out / "refused.txt"
        if refused.exists():
            reason = refused.read_text(encoding="utf-8").strip()
            records.append({"group": group, "case": f"refused: {reason}"})
            continue
        if not (out / "formats.tsv").exists():
            print(f"{group}: not captured yet, left out", file=sys.stderr)
            continue
        first = first or header(out)
        captures = load(out)
        records.extend(sizes(captures[0], group))
        for cap in captures:
            records.extend(measure(cap, group))
    with open(OUTPUT, "w", encoding="utf-8") as out_file:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out_file.write(f"# {key}: {value}\n")
        out_file.write("# suite: filters_more\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out_file.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cells in {argv[1]}")
        return 0
    if argv == ["run"]:
        run()
        return 0
    if argv[:1] == ["run"] and all(group in GROUPS for group in argv[1:]):
        run(tuple(argv[1:]))
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
