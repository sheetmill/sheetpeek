"""Table styles against Excel: the screen suite ``tables`` (ARCHITECTURE.md §4.20).

``tests/captures/tables.tsv.gz`` holds, for every table cell of the suite's groups, Excel's display
format (read exactly) and the device pixels across its edges (Microsoft 365, 16.115). Each group's
workbook is built again from ``capture/tables.py``, laid out, and compared cell by cell.
"""

from __future__ import annotations

import collections
import gzip
from pathlib import Path

import pytest

import sheetpeek
from sheetpeek import CellRange
from tests.capture_report import REPORT

openpyxl = pytest.importorskip("openpyxl")

CAPTURE = Path(__file__).with_name("captures") / "tables.tsv.gz"
UNDERLINES = {
    "none": None,
    "single": "single",
    "double": "double",
    "single accounting": "singleAccounting",
    "double accounting": "doubleAccounting",
}
LINE_AT = {
    "thin": (6, 7),
    "medium": (4, 5, 6, 7),
    "thick": (4, 5, 6, 7, 8, 9),
    "double": (4, 5, 8, 9),
}
"""Where each line's device pixels lie in a strip (at 2 pixels a point, the boundary after 6)."""
TOLERANCE = 8
"""Per channel: the screen's color conversion moves saturated colors (00B050 shows as 06B050)."""
KNOWN = {
    ("theme", "font_color"): (
        12,
        "theme color 1 untinted is drawn as automatic black (§3.3); the second theme's dark 1 "
        "is 1F3864",
    ),
}
"""Known deviations: (group, property) to (count, reason)."""


def _records() -> list[dict[str, str]]:
    lines = [
        ln
        for ln in gzip.decompress(CAPTURE.read_bytes()).decode().splitlines()
        if not ln.startswith("#")
    ]
    names = lines[0].split("\t")
    return [dict(zip(names, line.split("\t"), strict=False)) for line in lines[1:]]


def _hex(rgb: str) -> str:
    return "".join(f"{int(part):02X}" for part in rgb.split(","))


def _near(a: str, b: str) -> bool:
    return all(abs(int(a[i : i + 2], 16) - int(b[i : i + 2], 16)) <= TOLERANCE for i in (0, 2, 4))


def _line_matches(side, pixels: list[str]) -> bool:
    if side is None:  # no line: the boundary shows a fill on either side, or the gridline
        return all(
            _near(pixels[k], pixels[5]) or _near(pixels[k], pixels[8]) or _near(pixels[k], "D4D4D4")
            for k in (6, 7)
        )
    color = side.color or "000000"
    if side.style == "hair":  # dotted, so lighter where it is sampled: ink of its hue
        return all(
            not _near(pixels[k], pixels[5]) and not _near(pixels[k], pixels[8]) for k in (6, 7)
        )
    return all(_near(pixels[k], color) for k in LINE_AT[side.style])


@pytest.fixture(scope="module")
def compared(tmp_path_factory):
    tables = pytest.importorskip("capture.tables")
    records = _records()
    records = [r for r in records if not r["role"].startswith("refused")]
    groups = list(dict.fromkeys(r["group"] for r in records))
    differences: dict[tuple[str, str], list[str]] = collections.defaultdict(list)
    count = collections.Counter()
    folder = tmp_path_factory.mktemp("tables")
    for group in groups:
        path = folder / f"{group}.xlsx"
        wb = openpyxl.Workbook()
        tables.fill(wb.active, group)
        wb.save(path)
        post = tables.postprocess(group)
        if post is not None:
            post(path)
        rendered = sheetpeek.read(path).layout("excel")
        borders = rendered.borders  # None where nothing has a line
        at = {
            (c.row, c.column): (i, p, c.display_style)
            for i, row in enumerate(rendered.rows)
            for p, c in enumerate(row.cells)
        }
        for record in (r for r in records if r["group"] == group):
            where = CellRange.parse(record["cell"])
            i, p, style = at[(where.min_row, where.min_col)]
            font = style.font
            found = {
                "font_color": (_hex(record["font_color"]), font.color or "000000"),
                "bold": (record["bold"] == "true", font.bold),
                "italic": (record["italic"] == "true", font.italic),
                "strike": (record["strike"] == "true", font.strike),
                "underline": (UNDERLINES[record["underline"].removeprefix("underline style ")],
                              font.underline),
                "font_name": (record["font_name"], font.name),
            }  # fmt: skip
            spec = style.fill_spec
            if spec is None or getattr(spec, "kind", None) == "solid":
                excel = None if record["pattern"] == "pattern none" else _hex(record["fill"])
                found["fill"] = (excel, spec.foreground if spec is not None else None)
            for key, (excel, ours) in found.items():
                count[key] += 1
                if excel != ours:
                    differences[(group, key)].append(f"{record['cell']}: {excel} != {ours}")
            if spec is not None and getattr(spec, "kind", None) != "solid":
                continue  # a pattern's dots lie across the edges
            sides = {
                "top": borders.horizontal[i][p] if borders else None,
                "left": borders.vertical[i][p] if borders else None,
                "bottom": borders.horizontal[i + 1][p] if borders else None,
                "right": borders.vertical[i][p + 1] if borders else None,
            }
            for edge, side in sides.items():
                if not record[f"edge_{edge}"]:
                    continue
                count["edge"] += 1
                if not _line_matches(side, tables.strip(record[f"edge_{edge}"])):
                    differences[(group, "edge")].append(
                        f"{record['cell']} {edge}: {side} vs {record[f'edge_{edge}']}"
                    )
    return count, differences


def test_every_table_cell_is_displayed_as_excel_displays_it(compared):
    count, differences = compared
    unexpected = {k: v for k, v in differences.items() if k not in KNOWN}
    assert not unexpected, {k: v[:5] for k, v in unexpected.items()}
    for key, (expected, _reason) in KNOWN.items():
        assert len(differences.get(key, [])) == expected, key
    REPORT["tables"] = (
        f"tables: {count['edge']} edges and {count['fill']} fills compared with Excel, "
        f"{sum(len(v) for v in differences.values())} known deviations: "
        + "; ".join(f"{group} {key} {n} ({reason})" for (group, key), (n, reason) in KNOWN.items())
    )


def test_every_group_is_captured():
    tables = pytest.importorskip("capture.tables")
    assert {r["group"] for r in _records()} == set(tables.GROUPS)


def test_a_merge_inside_a_table_is_shown_where_excel_refuses_the_file(tmp_path):
    # Known deviation: Excel refuses the file with alerts off (it offers a repair); files it
    # writes cannot hold such a merge. sheetpeek shows it, the merge in its anchor's own fill.
    tables = pytest.importorskip("capture.tables")
    (row,) = [r for r in _records() if r["group"] == "merge"]
    assert row["role"].startswith("refused: ")
    path = tmp_path / "merge.xlsx"
    wb = openpyxl.Workbook()
    tables.fill(wb.active, "merge")
    wb.save(path)
    rendered = sheetpeek.read(path).layout("excel")
    merged = [c for r in rendered.rows for c in r.cells if c.merge is not None]
    assert merged
    assert {c.fill for c in merged} == {None}  # the anchor's own fill
    REPORT["tables merge"] = "tables: a merge inside a table, which Excel refuses, is shown (1)"
