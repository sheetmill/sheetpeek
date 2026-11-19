"""Capture whether Excel for Mac matches sheet extension URIs ignoring case (route C, and saves).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.extensions run
    uv run --group capture python -m capture.extensions read        # the saved captures again
    uv run --group capture python -m capture.extensions build PATH  # every group's workbook

ARCHITECTURE.md §3.7, "A file is reported for what it uses": the known sheet extension URIs
other than conditional formatting's (sparklines, validation, protected ranges, ignored errors)
are matched ignoring case until captured (unverified), and ``_detect._URIS_ANY_CASE`` follows
it. Conditional formatting's URI is already captured as compared exactly; this suite asks the
same of the others, each under its exact URI and under the URI in lower case (every one of them
is printed in upper case, so lower case is the re-cased form).

Two workbooks' sheets ``Ext`` (first and active) and ``Lists`` hold the same cells; the
extensions are hand-written ``x14`` XML that ``postprocess`` puts in the sheet's ``extLst``
(openpyxl writes none). On ``Ext``:

- C4:H4 hold 1 to 6 and J4 (20 characters wide, its row 30 pt) a column sparkline of them,
  series blue ``1F4E79``;
- C7, E7 and G7 hold the text ``123`` (a number stored as text) and C9, E9 and G9 ``=1/0``
  (an error): Excel's background error checking marks each with a green triangle in its top
  left corner. C7 and C9 are left marked (the control), E7 and E9 are exempted by the sheet's
  own ``ignoredErrors`` (the control of the exemption), G7 and G9 by the ``x14`` ignored
  errors extension;
- C11 has a list validation from ``Lists!$A$1:$A$3`` (the ``x14`` validation extension) and
  C13 lies in the protected range ``ExtRange`` (the ``x14`` protected ranges extension).

Groups, each its own workbook opened on its own (:data:`GROUPS`): ``exact`` holds the four
extensions under their exact URIs, the control; ``<kind>_lower`` holds only that kind's
extension, under its URI in lower case. A lower-case group Excel refuses is recorded
(``refused.txt``) as the answer and the run goes on; a refusal of ``exact`` stops the run, as
its XML would then be wrong.

Two steps for each group:

- ``screen`` (``exact``, ``sparkline_lower``, ``ignored_lower``): one window at zoom 100,
  markers B2 and L15. Measured in J4, 3 image pixels in from its sides: the ink, pixels more
  than 100 from the cell's most common color (``ink_pixels``, their mean color ``ink_color``,
  their box's width and height in points); in each of the six error cells, the top left
  square of 6 points, 1 image pixel in from the cell's left and top: the green pixels (green
  over 40 above red and blue, ``triangle_pixels``, their mean ``triangle_color``). Nothing else
  on the sheet is green.
- ``save``: no window. Excel saves the open workbook as a copy (``ExcelSession.save_copy``)
  and the copy's ``Ext`` sheet is kept (``saved_sheet.xml``). For each kind the group holds, it
  is found where Excel wrote it: ``main`` (the sheet's own element), ``exact`` (an extension
  under the exact URI), ``other:<uri>`` (kept under another URI, as an extension Excel did not
  read), or nowhere (dropped). This is the only evidence for validation and protected ranges,
  which draw nothing while the selection is parked; a kind written back as ``main`` or
  ``exact`` was read.

Not covered: web extensions (``x15:webExtensions``) need an add-in's parts, and slicers and
timelines need pivot or table caches; their URIs stay unverified.

``read`` writes ``tests/captures/extensions.tsv``: one row per group and measured cell
(``kind`` ``screen``), one per group and extension kind (``kind`` ``save``), and one per
refused group (``kind`` ``refused``).
"""

from __future__ import annotations

import os
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "extensions"
OUTPUT = ROOT / "tests" / "captures" / "extensions.tsv"
ZOOM = 100
MARKERS = ("B2", "L15")
SHEET = "Ext"
LISTS = "Lists"
URIS = {
    "sparkline": "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}",
    "validation": "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}",
    "protected": "{FC87AEE6-9EDD-4A0A-B7FB-166176984837}",
    "ignored": "{01252117-D84E-4E92-8308-4BE1C098FCBB}",
}
"""Each extension kind's URI as [MS-XLSX] prints it (all upper case)."""
GROUPS: dict[str, dict[str, str]] = {
    "exact": dict(URIS),
    **{f"{kind}_lower": {kind: uri.lower()} for kind, uri in URIS.items()},
}
"""Each group's extensions, kind → the URI written."""
SCREEN = ("exact", "sparkline_lower", "ignored_lower")
"""Groups whose window is captured: the others draw nothing a window shows."""
SPARKLINE = "J4"
NUMBER_TEXT = ("C7", "E7", "G7")
ERRORS = ("C9", "E9", "G9")
CASES = {
    SPARKLINE: "column sparkline of C4:H4",
    "C7": "number as text, not exempted",
    "E7": "number as text, exempted by ignoredErrors",
    "G7": "number as text, exempted by the extension",
    "C9": "=1/0, not exempted",
    "E9": "=1/0, exempted by ignoredErrors",
    "G9": "=1/0, exempted by the extension",
}
SERIES = "1F4E79"
OTHER = "D00000"
"""The sparkline's series color and the color of its other parts (none drawn by default)."""
_OTHER_PARTS = (
    "colorNegative", "colorAxis", "colorMarkers", "colorFirst", "colorLast", "colorHigh",
    "colorLow",
)  # fmt: skip
VALIDATION = "C11"
PROTECTED = "C13"
RANGE_NAME = "ExtRange"
INSET = 3
INK = 100
TRIANGLE = 6
"""The side, in points, of the corner square searched for the error indicator."""
GREEN = 40
COLUMNS = (
    "kind", "group", "cell", "case", "zoom", "ink_pixels", "ink_color", "ink_w", "ink_h",
    "triangle_pixels", "triangle_color", "extension", "uri", "saved_where", "saved_uris",
)  # fmt: skip
_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_XM = "http://schemas.microsoft.com/office/excel/2006/main"


def expected(group: str) -> dict[str, bool]:
    """Whether sheetpeek takes each of a group's extensions as the kind it is (by its URI).

    Taken from sheetpeek's own table, not a guess about Excel: the URIs in ``_URIS_EXACT``
    compared exactly, those in ``_URIS_ANY_CASE`` ignoring case.
    """
    from sheetpeek import _detect

    return {
        kind: uri in _detect._URIS_EXACT or uri.upper() in _detect._URIS_ANY_CASE
        for kind, uri in GROUPS[group].items()
    }


def fill(ws: Any) -> list[tuple[str, str]]:
    """Lay the cells out on ``Ext`` (and its list on ``Lists``); each measured cell with its case.

    The extensions are written by :func:`postprocess`, as openpyxl writes no ``extLst``.
    """
    ws.title = SHEET
    lists = ws.parent.create_sheet(LISTS)
    for row, item in enumerate(("one", "two", "three"), start=1):
        lists.cell(row, 1, item)
    for k in range(6):
        ws.cell(4, 3 + k, k + 1)
    ws.column_dimensions["J"].width = 20
    ws.row_dimensions[4].height = 30
    for cell in NUMBER_TEXT:
        ws[cell] = "123"
    for cell in ERRORS:
        ws[cell] = "=1/0"
    ws[VALIDATION] = "one"
    ws[PROTECTED] = "protected"
    return list(CASES.items())


def _ext(uri: str, body: str) -> str:
    return f'<ext uri="{uri}" xmlns:x14="{_X14}">{body}</ext>'


def extensions(group: str) -> str:
    """The group's ``extLst``, its extensions in the order Excel writes them."""
    bodies = {
        "validation": (
            f'<x14:dataValidations count="1" xmlns:xm="{_XM}"><x14:dataValidation type="list" '
            'allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>'
            f"{LISTS}!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>{VALIDATION}</xm:sqref>"
            "</x14:dataValidation></x14:dataValidations>"
        ),
        "sparkline": (
            f'<x14:sparklineGroups xmlns:xm="{_XM}"><x14:sparklineGroup type="column" '
            f'displayEmptyCellsAs="gap"><x14:colorSeries rgb="FF{SERIES}"/>'
            + "".join(f'<x14:{name} rgb="FF{OTHER}"/>' for name in _OTHER_PARTS)
            + f"<x14:sparklines><x14:sparkline><xm:f>{SHEET}!C4:H4</xm:f><xm:sqref>{SPARKLINE}"
            "</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>"
            "</x14:sparklineGroups>"
        ),
        "protected": (
            f'<x14:protectedRanges xmlns:xm="{_XM}"><x14:protectedRange name="{RANGE_NAME}">'
            f"<xm:sqref>{PROTECTED}</xm:sqref></x14:protectedRange></x14:protectedRanges>"
        ),
        "ignored": (
            f'<x14:ignoredErrors xmlns:xm="{_XM}"><x14:ignoredError numberStoredAsText="1">'
            f"<xm:sqref>{NUMBER_TEXT[2]}</xm:sqref></x14:ignoredError>"
            f'<x14:ignoredError evalError="1"><xm:sqref>{ERRORS[2]}</xm:sqref>'
            "</x14:ignoredError></x14:ignoredErrors>"
        ),
    }
    chosen = GROUPS[group]
    return (
        "<extLst>"
        + "".join(_ext(chosen[kind], bodies[kind]) for kind in bodies if kind in chosen)
        + "</extLst>"
    )


IGNORED = (
    f'<ignoredErrors><ignoredError sqref="{NUMBER_TEXT[1]}" numberStoredAsText="1"/>'
    f'<ignoredError sqref="{ERRORS[1]}" evalError="1"/></ignoredErrors>'
)
"""The sheet's own exemptions of E7 and E9, the control of the extension's."""


def sheet(text: str, group: str) -> str:
    """The ``Ext`` sheet's XML with its ``ignoredErrors`` and the group's ``extLst``.

    Both go at the end: openpyxl writes nothing after ``pageMargins`` for this sheet, and
    ``ignoredErrors`` and ``extLst`` follow it in the schema's order.
    """
    tail = "</worksheet>"
    if text.count(tail) != 1 or not text.rstrip().endswith(tail) or "<extLst" in text:
        raise ValueError("the sheet does not end as openpyxl writes it")
    before = text.rstrip()[: -len(tail)]
    if not before.rstrip().endswith("/>") or "pageMargins" not in before.rsplit("<", 1)[1]:
        raise ValueError("the sheet's last element is not pageMargins")
    return before + IGNORED + extensions(group) + tail


def postprocess(group: str) -> Any:
    """What a group's saved workbook needs: the ``Ext`` sheet's exemptions and extensions."""
    from capture.screen.run import patch_part

    def edit(path: Path) -> None:
        patch_part(path, {"xl/worksheets/sheet1.xml": lambda text: sheet(text, group)})

    return edit


def build(path: Path) -> list[tuple[str, str, str]]:
    """Write every group to a workbook of its own in the folder ``path``; each group's cases."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        placed = fill(wb.active)
        target = path / f"extensions_{group}.xlsx"
        wb.save(target)
        postprocess(group)(target)
        found.extend((group, cell, case) for cell, case in placed)
    return found


# -- running (needs Excel) -----------------------------------------------------------------------


def run(groups: tuple[str, ...] | None = None) -> None:
    """Capture and save each group on its own, then write the TSV (needs Excel)."""
    from capture.screen.run import capture
    from capture.screen.session import ExcelError

    for group in groups or tuple(GROUPS):
        try:
            if group in SCREEN:
                capture(
                    fill,
                    cells=list(CASES),
                    out=BUILD_DIR / group,
                    markers=MARKERS,
                    zooms=(ZOOM,),
                    name=f"extensions_{group}",
                    postprocess=postprocess(group),
                )
            save(group)
        except ExcelError as error:
            if group == "exact":
                raise
            record_refusal(group, str(error))
            print(f"{group}: Excel refused the file", file=sys.stderr)
    write()


def save(group: str) -> None:
    """Open the group's workbook alone, have Excel save a copy, keep the copy's ``Ext`` sheet."""
    from capture.screen.session import CONTAINER, ExcelSession

    out = BUILD_DIR / group
    out.mkdir(parents=True, exist_ok=True)
    folder = CONTAINER / "sheetpeek-capture"
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / f"extensions_{group}_{os.getpid()}.xlsx"
    copy = folder / f"extensions_{group}_saved_{os.getpid()}.xlsx"
    session = ExcelSession(path, evidence=out / "excel_timeouts")
    try:
        wb = openpyxl.Workbook()
        fill(wb.active)
        wb.save(path)
        postprocess(group)(path)
        with session:
            session.save_copy(copy)
            (out / "saved_sheet.xml").write_bytes(first_sheet(copy))
    finally:
        path.unlink(missing_ok=True)
        copy.unlink(missing_ok=True)


def first_sheet(path: Path) -> bytes:
    """The first sheet's XML in a saved package, found through the workbook's relationships."""
    rel = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    with zipfile.ZipFile(path) as package:
        book = ET.fromstring(package.read("xl/workbook.xml"))
        first = next(book.iter(f"{{{_MAIN}}}sheet")).get(rel)
        rels = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
        target = next(r.get("Target", "") for r in rels if r.get("Id") == first)
        name = target.lstrip("/") if target.startswith("/") else f"xl/{target}"
        return package.read(name)


def record_refusal(group: str, message: str) -> None:
    """Keep Excel's refusal of a group's file in ``refused.txt``, without the script's path."""
    out = BUILD_DIR / group
    out.mkdir(parents=True, exist_ok=True)
    reason = message.rpartition("execution error: ")[2].strip() or message
    (out / "refused.txt").write_text(reason + "\n", encoding="utf-8")


# -- measuring -----------------------------------------------------------------------------------


def _distance(a: tuple[int, int, int], b: tuple[int, int, int]) -> float:
    return float(sum((x - y) ** 2 for x, y in zip(a, b, strict=True)) ** 0.5)


def _hex(pixels: list[tuple[int, int, int]]) -> str:
    return "".join(f"{round(sum(p[i] for p in pixels) / len(pixels)):02X}" for i in range(3))


def _green(rgb: tuple[int, int, int]) -> bool:
    r, g, b = rgb
    return g - max(r, b) > GREEN


def measure(cap: Any) -> list[dict[str, str]]:
    """The sparkline cell's ink and each error cell's indicator at one zoom."""
    per = cap.scale * 100 / cap.zoom  # image pixels a screen point
    found = []
    x0, x1 = round(cap.edge(SPARKLINE, "left")) + INSET, round(cap.edge(SPARKLINE, "right")) - INSET
    y0, y1 = round(cap.edge(SPARKLINE, "top")) + INSET, round(cap.edge(SPARKLINE, "bottom")) - INSET
    pixels = [(x, y, cap.image.pixel(x, y)) for y in range(y0, y1) for x in range(x0, x1)]
    background = Counter(p for _, _, p in pixels).most_common(1)[0][0]
    ink = [(x, y, p) for x, y, p in pixels if _distance(p, background) > INK]
    record = {"cell": SPARKLINE, "case": CASES[SPARKLINE], "zoom": str(cap.zoom)}
    record["ink_pixels"] = str(len(ink))
    if ink:
        width = max(x for x, _, _ in ink) - min(x for x, _, _ in ink) + 1
        height = max(y for _, y, _ in ink) - min(y for _, y, _ in ink) + 1
        record |= {
            "ink_color": _hex([p for _, _, p in ink]),
            "ink_w": f"{round(width / per, 2):g}",
            "ink_h": f"{round(height / per, 2):g}",
        }
    else:
        record["ink_color"] = "none"
    found.append(record)
    side = round(TRIANGLE * per)
    for cell in (*NUMBER_TEXT, *ERRORS):
        left, top = round(cap.edge(cell, "left")) + 1, round(cap.edge(cell, "top")) + 1
        green = [
            p
            for y in range(top, top + side)
            for x in range(left, left + side)
            if _green(p := cap.image.pixel(x, y))
        ]
        found.append(
            {
                "cell": cell,
                "case": CASES[cell],
                "zoom": str(cap.zoom),
                "triangle_pixels": str(len(green)),
                "triangle_color": _hex(green) if green else "none",
            }
        )
    return found


def _local(element: ET.Element) -> str:
    return element.tag.rpartition("}")[2]


def _holds(element: ET.Element, kind: str) -> bool:
    """Whether ``element`` (a sheet child or an extension) holds this suite's ``kind``."""
    for e in element.iter():
        name = _local(e)
        if kind == "sparkline" and name == "sparklineGroup":
            return True
        if kind == "validation" and name in ("f", "formula1") and f"{LISTS}!" in (e.text or ""):
            return True
        if kind == "protected" and name == "protectedRange" and e.get("name") == RANGE_NAME:
            return True
        if kind == "ignored" and name == "ignoredError":
            sqref = e.get("sqref") or "".join(c.text or "" for c in e if _local(c) == "sqref")
            if {NUMBER_TEXT[2], ERRORS[2]} & set(sqref.split()):
                return True
    return False


def saved(text: bytes, kind: str) -> tuple[str, str]:
    """Where Excel wrote ``kind`` in a saved sheet, and every extension URI it wrote, in order.

    ``main`` (the sheet's own element), ``exact`` (an extension under the exact URI),
    ``other:<uri>``, or ``""`` (nowhere).
    """
    root = ET.fromstring(text)
    uris: list[str] = []
    where = ""
    for child in root:
        if _local(child) == "sheetData":
            continue
        if _local(child) == "extLst":
            for ext in child:
                uri = ext.get("uri", "")
                uris.append(uri)
                if not where and _holds(ext, kind):
                    where = "exact" if uri == URIS[kind] else f"other:{uri}"
        elif not where and _holds(child, kind):
            where = "main"
    return where, " ".join(uris)


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``extensions.tsv``."""
    from capture.screen.run import header, load

    records: list[dict[str, str]] = []
    first: dict[str, str] = {}
    for group, chosen in GROUPS.items():
        out = BUILD_DIR / group
        refused = out / "refused.txt"
        if refused.exists():
            reason = refused.read_text(encoding="utf-8").strip()
            records.append({"kind": "refused", "group": group, "case": f"refused: {reason}"})
            continue
        if group in SCREEN:
            if (out / "formats.tsv").exists():
                first = first or header(out)
                for cap in load(out):
                    records.extend({"kind": "screen", "group": group} | r for r in measure(cap))
            else:
                print(f"{group}: screen not captured yet, left out", file=sys.stderr)
        if (out / "saved_sheet.xml").exists():
            text = (out / "saved_sheet.xml").read_bytes()
            for kind, uri in chosen.items():
                where, uris = saved(text, kind)
                records.append(
                    {
                        "kind": "save", "group": group, "extension": kind, "uri": uri,
                        "saved_where": where, "saved_uris": uris,
                    }
                )  # fmt: skip
        else:
            print(f"{group}: save not captured yet, left out", file=sys.stderr)
    with open(OUTPUT, "w", encoding="utf-8") as file:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                file.write(f"# {key}: {value}\n")
        file.write("# suite: extensions\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            file.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cases in {argv[1]}")
        return 0
    if argv[:1] == ["run"] and all(group in GROUPS for group in argv[1:]):
        run(tuple(argv[1:]) or None)
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
