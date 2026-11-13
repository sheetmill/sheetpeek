"""Rotated and stacked text (ARCHITECTURE.md §4.17).

Expected placements come from Excel's screen (the ``rotation`` suite's first cases, Calibri 11,
zoom 100, default 53 pt columns): the ink of "Rotated text here" from the cell's left edge, and
the fitted row heights. The ink is modeled as the glyphs' box inside the line: 0.8 pt in from
each end, from the baseline to 7.5 pt above it.
"""

from __future__ import annotations

import itertools
import math
import re

import openpyxl
import pytest
from openpyxl.styles import Alignment, Border, Font, Side

import sheetpeek
from sheetpeek._style.alignment import _rotation
from tests.workbooks import save, shared_strings

TEXT = "Rotated text here"
# angle, horizontal alignment: (ink's left edge, ink's right edge past the cell's right edge,
# fitted row height), captured.
CAPTURED = {
    (15, None): (4.5, 27.5, 38),
    (45, None): (6, 11.5, 69),
    (90, None): (38.5, -7, 85),
    (135, None): (-12.5, -8, 69),
    (180, None): (6, -39.5, 85),
    (45, "center"): (-3, 2.5, 69),
    (45, "right"): (-12, -6.5, 69),
}


def _book(tmp_path, cells, name="rotation.xlsx"):
    """A sheet with ``cells``: (coordinate, value, alignment, extra style) each, rows listed."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for coordinate, value, alignment, extra in cells:
        ws[coordinate] = value
        ws[coordinate].alignment = alignment
        for attribute, setting in extra.items():
            setattr(ws[coordinate], attribute, setting)
    return save(wb, tmp_path / name)


def _row(rendered, index):
    return next(r for r in rendered.rows if r.index == index)


def _ink(run, column):
    """The run's modeled ink, as (left from the column's left edge, right past its right edge)."""
    angle = math.radians(run.angle)
    xs = [
        run.x + s * math.cos(angle) - t * math.sin(angle)
        for s in (0.8, run.width - 0.8)
        for t in (0.0, 7.5)
    ]
    return min(xs) - column.x_pt, max(xs) - column.x_pt - column.width_pt


# The top of each letter of the captured texts, in em: Calibri's glyph bounds (Calibri.ttf, as
# Excel for Mac ships it). The first and last letter's tops bound a line's turned ink.
_TOPS = {
    ",": 0.109, "R": 0.632, "a": 0.474, "d": 0.678, "e": 0.474, "f": 0.684, "h": 0.68,
    "i": 0.648, "l": 0.68, "m": 0.474, "n": 0.474, "o": 0.474, "r": 0.474, "s": 0.474,
    "t": 0.584, "w": 0.467, "x": 0.467,
}  # fmt: skip


def _glyph_ink(run, column):
    """``_ink`` with the first and last letters' own heights, where ``_ink`` takes 7.5 pt."""
    angle = math.radians(run.angle)
    size = run.font.size * run.scale
    xs = [
        run.x + s * math.cos(angle) - t * math.sin(angle)
        for s, letter in ((0.8, run.text[0]), (run.width - 0.8, run.text[-1]))
        for t in (0.0, _TOPS.get(letter, 0.68) * size)
    ]
    return min(xs) - column.x_pt, max(xs) - column.x_pt - column.width_pt


@pytest.fixture(scope="module")
def captured(tmp_path_factory):
    cells = []
    for row, ((rotation, horizontal), _) in enumerate(CAPTURED.items(), 2):
        cells.append(
            (f"C{row * 2}", TEXT, Alignment(textRotation=rotation, horizontal=horizontal), {})
        )
    return sheetpeek.read(_book(tmp_path_factory.mktemp("rotation"), cells)).layout("excel")


@pytest.mark.parametrize(
    ("stored", "rotation", "stacked"),
    [(0, 0, False), (15, 15, False), (90, 90, False), (91, -1, False), (135, -45, False),
     (180, -90, False), (255, 0, True)],
)  # fmt: skip
def test_text_rotation_is_read_as_an_angle(tmp_path, stored, rotation, stacked):
    path = _book(tmp_path, [("A1", "x", Alignment(textRotation=stored), {})])
    style = sheetpeek.read(path).cell(1, 1).style
    assert (style.rotation, style.stacked) == (rotation, stacked)


@pytest.mark.parametrize("value", [-1, 181, 254, 256, 1000])
def test_a_rotation_outside_the_standard_is_upright(value):
    # openpyxl refuses such a workbook (tests/test_limits.py); this guards the reader move.
    assert _rotation(value) == 0


@pytest.mark.parametrize("key", CAPTURED, ids=[f"{a}-{h}" for a, h in CAPTURED])
def test_angled_text_is_placed_as_captured(captured, key):
    row = (list(CAPTURED).index(key) + 2) * 2
    rendered = next(r for r in captured.rows if r.index == row)
    (run,) = rendered.runs
    column = next(c for c in captured.columns if c.index == 3)
    left, right, height = CAPTURED[key]
    ink = _ink(run, column)
    assert ink[0] == pytest.approx(left, abs=1.1)
    assert ink[1] == pytest.approx(right, abs=1.1)
    assert rendered.height_pt == pytest.approx(height, abs=1.0)  # 70 against 69 at ±45 degrees
    assert run.angle == (key[0] if key[0] <= 90 else 90 - key[0])


def test_angled_text_spills_by_its_footprint(captured):
    spill = {row: captured.cell(row, 3).spill for row in (6, 10, 14)}  # 45, -45, 45 centered
    assert spill[6][0] == 0  # right, for a positive angle
    assert spill[6][1] > 0
    assert spill[10][0] > 0  # left, for a negative one
    assert spill[10][1] == 0
    assert min(spill[14]) > 0  # both ways, centered


def test_a_neighbor_with_a_value_stops_the_spill(tmp_path):
    path = _book(
        tmp_path,
        [("B2", TEXT, Alignment(textRotation=45), {}), ("C2", "x", Alignment(), {})],
    )
    rendered = sheetpeek.read(path).layout("excel")
    run = next(r for r in _row(rendered, 2).runs if r.source == (2, 2))
    assert rendered.cell(2, 2).spill == (0, 0)
    column = next(c for c in rendered.columns if c.index == 2)
    assert run.clip[2] == column.x_pt + column.width_pt


def test_a_custom_height_row_clips_at_the_row_and_keeps_the_text_at_the_bottom(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["B2"] = TEXT
    wb.active["B2"].alignment = Alignment(textRotation=90)
    wb.active.row_dimensions[2].height = 15
    rendered = sheetpeek.read(save(wb, tmp_path / "custom.xlsx")).layout("excel")
    row = _row(rendered, 2)
    (run,) = row.runs
    assert row.height_pt == 15
    assert run.clip[1] == row.y_pt + 1
    assert run.clip[3] == row.y_pt + row.height_pt
    assert run.y == pytest.approx(row.y_pt + row.height_pt)  # the start at the bottom: "Ro"


def test_stacked_text_draws_a_character_a_line(tmp_path):
    path = _book(tmp_path, [("C2", "Stack", Alignment(textRotation=255), {})])
    rendered = sheetpeek.read(path).layout("excel")
    row = _row(rendered, 2)
    assert row.height_pt == 77  # five lines of 15 pt and 2 (captured)
    assert [r.text for r in row.runs] == list("Stack")
    assert {r.align for r in row.runs} == {"center"}
    baselines = [r.y for r in row.runs]
    assert [b - a for a, b in itertools.pairwise(baselines)] == [15.0] * 4
    column = next(c for c in rendered.columns if c.index == 3)
    for run in row.runs:  # centered on the cell (captured)
        center = run.x + run.width / 2 - column.x_pt
        assert center == pytest.approx(column.width_pt / 2, abs=1.0)


def test_wrapped_angled_text_breaks_at_line_feeds_and_stacked_text_is_reported(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    for coordinate, rotation in (("B2", 45), ("B4", 255), ("B6", 45)):
        ws[coordinate] = "two\nlines"
        ws[coordinate].alignment = Alignment(textRotation=rotation, wrap_text=True)
    ws.row_dimensions[2].height = 60
    rendered = sheetpeek.read(save(wb, tmp_path / "wrapped.xlsx")).layout("excel")
    assert [run.text for run in _row(rendered, 2).runs] == ["two", "lines"]
    # Wrapped angled text leaves its row as it is (captured): in 15 pt a line holds one character.
    assert _row(rendered, 6).height_pt == 15
    assert [run.text for run in _row(rendered, 6).runs] == list("twolines")
    (notice,) = [n for n in rendered.notices if n.code == "text.rotation_wrapped"]
    assert (notice.outcome, notice.count) == ("approximated", 1)  # the stacked cell


def test_only_merged_angled_cells_with_borders_are_reported(tmp_path):
    # An unmerged cell's slanted borders and the text between them are drawn (§4.17); a merge's
    # stay upright, reported.
    thin = Side("thin")
    boxed = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = openpyxl.Workbook()
    ws = wb.active
    for coordinate, value, rotation in (
        ("B2", "a", 45), ("B4", "b", 90), ("B6", "c", 255), ("B8", "d", 45), ("B10", "e", 45),
    ):  # fmt: skip
        ws[coordinate] = value
        ws[coordinate].alignment = Alignment(textRotation=rotation)
        if coordinate != "B8":
            ws[coordinate].border = boxed
    ws.merge_cells("B10:C10")
    rendered = sheetpeek.read(save(wb, tmp_path / "boxed.xlsx")).layout("excel")
    found = sheetpeek.notices(rendered, "html")
    (slanted,) = [n for n in found if n.code == "border.slanted"]
    assert (slanted.count, slanted.outcome) == (1, "approximated")
    assert (slanted.cells[0].min_row, slanted.cells[0].min_col) == (10, 2)
    assert not [n for n in found if n.code in ("text.rotation", "text.stacked")]  # drawn


def test_other_outputs_report_angled_and_stacked_text(tmp_path):
    path = _book(
        tmp_path,
        [
            ("B2", "a", Alignment(textRotation=45), {}),
            ("B4", "b", Alignment(textRotation=255), {}),
        ],
    )
    sheet = sheetpeek.read(path)
    found = {n.code: n.outcome for n in sheetpeek.notices(sheet.layout("characters"), "text")}
    assert found["text.rotation"] == found["text.stacked"] == "approximated"
    found = {n.code: n.outcome for n in sheetpeek.notices(sheet.layout("none"), "markdown")}
    assert found["text.rotation"] == found["text.stacked"] == "omitted"


def test_html_turns_the_run_and_its_rules_about_the_baseline_start(tmp_path):
    path = _book(
        tmp_path,
        [("B2", "under", Alignment(textRotation=30), {"font": Font(underline="single")})],
    )
    markup = sheetpeek.render(path, to="html", on_notice="ignore")
    turns = re.findall(r"transform:rotate\((-?[\d.]+)deg\);transform-origin:([^\"]+)\"", markup)
    assert [angle for angle, _ in turns] == ["-30", "-30"]  # the run and its underline
    assert turns[0][1].startswith("0px ")
    assert "sp-clip" in markup  # an angled run is always clipped by its window


def test_angled_rich_text_keeps_its_runs(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["B2"] = "placeholder"
    wb.active["B2"].alignment = Alignment(textRotation=60)
    path = save(wb, tmp_path / "runs.xlsx")
    calibri = '<sz val="11"/><rFont val="Calibri"/>'
    shared_strings(
        path,
        {
            "B2": f"<r><rPr><b/>{calibri}</rPr><t>bold</t></r>"
            f'<r><rPr>{calibri}</rPr><t xml:space="preserve"> plain</t></r>'
        },
    )
    rendered = sheetpeek.read(path).layout("excel")
    (run,) = _row(rendered, 2).runs
    assert [s.text for s in run.spans] == ["bold", " plain"]
    assert run.angle == 60
    markup = sheetpeek.render(rendered, to="html", on_notice="ignore")
    assert "transform:rotate(-60deg)" in markup


def test_the_capture_workbook_reads_back_as_built(tmp_path):
    from capture.rotation import build

    placed = build(tmp_path / "rotation.xlsx")
    book = sheetpeek.open(tmp_path / "rotation.xlsx")
    sheets = {name: book.sheet(name) for name in dict.fromkeys(name for name, _, _ in placed)}
    assert len(placed) > 150
    for name, coordinate, case in placed:
        row, column = int(coordinate[1:]), 5
        style = sheets[name].cell(row, column).style
        expected = 0 if case.rotation == 255 else _rotation(case.rotation)
        assert (style.rotation, style.stacked) == (expected, case.rotation == 255), case.name
        if case.borders == "lr":
            assert style.borders.left is not None
            assert style.borders.top is None


@pytest.mark.parametrize("rotation", [45, 90])
def test_an_angled_number_shows_in_full_never_hashes(tmp_path, rotation):
    # Captured (rotation suite, "number 45" and "number 90"): Excel shows 12345678901 in a
    # default column, where upright it would show #.
    path = _book(tmp_path, [("B2", 12345678901, Alignment(textRotation=rotation), {})])
    rendered = sheetpeek.read(path).layout("excel")
    (run,) = _row(rendered, 2).runs
    assert run.text == "12345678901"
    assert rendered.cell(2, 2).text == "12345678901"
    assert not rendered.cell(2, 2).overflow


def _slant_captures():
    """The captured slant cases with left and right borders: (name, angle, row height, the four
    border x's at 25 % and 75 % of the row)."""
    from pathlib import Path

    path = Path(__file__).parent / "captures" / "rotation.tsv"
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    header = next(line for line in lines if line[0] == "group")
    rows = [dict(zip(header, line, strict=False)) for line in lines if line[0] == "slant"]
    return [
        (
            row["name"],
            _rotation(int(row["rotation"])),
            float(row["row_height"]),
            [
                float(row[f"border_{side}_{share}"])
                for share in (25, 75)
                for side in ("left", "right")
            ],
        )
        for row in rows
        if row["border_left_25"] and "fill" not in row["name"]
    ]


@pytest.mark.parametrize(("name", "angle", "height", "captured"), _slant_captures())
def test_slanted_sides_lie_where_excel_draws_them(tmp_path, name, angle, height, captured):
    # Captured (rotation suite, slant): the dark run of each side at 25 % and 75 % of the row,
    # 1 / |sin| wide for a 1 pt line. The modeled line's center lies within 1.25 pt of the run's,
    # measured across the line (15 degrees is the farthest, 1.05 pt, over 5 to 85 and -15 to -75
    # degrees, row heights 30 to 120 and columns of 4 to 20 characters).
    thin = Side("thin", color="FF000000")
    words = name.split()
    letters = {"l": "left", "r": "right", "t": "top", "b": "bottom"}
    sides = {letters[letter]: thin for letter in words[1]}
    wb = openpyxl.Workbook()
    ws = wb.active
    if "width" in words:
        ws.column_dimensions["E"].width = float(words[words.index("width") + 1])
    ws["E4"] = "Rotated text here"
    ws["E4"].alignment = Alignment(textRotation=_stored(angle))
    ws["E4"].border = Border(**sides)
    ws.row_dimensions[4].height = height  # fitted heights are tested on their own
    rendered = sheetpeek.read(save(wb, tmp_path / "slant.xlsx")).layout("excel")
    borders = rendered.borders
    if angle == 90:
        assert borders.slanted == ()
        return
    (slant,) = borders.slanted
    assert slant.shift_pt == pytest.approx(height / math.tan(math.radians(angle)))
    sine = abs(math.sin(math.radians(angle)))
    width = next(c for c in rendered.columns if c.letter == "E").width_pt
    for k, share in enumerate((0.25, 0.75)):
        base = slant.shift_pt * (1 - share)
        # With the left side alone, the darkest run's end is that side's too.
        far = (base + width, captured[2 * k + 1], True) if "right" in sides else None
        for model, start, end_side in [(base, captured[2 * k], False), *([far] if far else [])]:
            center = start - 1 / (2 * sine) if end_side else start + 1 / (2 * sine)
            assert abs(model - center) * sine <= 1.25, (name, share, model, center)


def _stored(angle):
    return angle if angle >= 0 else 90 - angle


def test_slanted_sides_leave_the_upright_grid_but_a_neighbors_side_stays(tmp_path):
    thin, medium = Side("thin"), Side("medium")
    path = _book(
        tmp_path,
        [
            ("B2", "a", Alignment(textRotation=45), {"border": Border(left=thin, right=thin)}),
            ("C2", None, Alignment(), {"border": Border(left=medium)}),
        ],
    )
    rendered = sheetpeek.read(path).layout("excel")
    borders = rendered.borders
    (slant,) = borders.slanted
    assert slant.left is not None
    assert slant.left == slant.right
    assert (slant.top, slant.bottom) == (None, None)
    p = [c.letter for c in rendered.columns].index("C")
    i = [r.index for r in rendered.rows].index(2)
    assert borders.vertical[i][p].style == "medium"  # C2's own left, upright
    assert borders.vertical[i][p - 1] is None  # B2's left left the grid


def test_borders_stay_upright_at_90_stacked_and_in_a_merge(tmp_path):
    thin = Side("thin")
    boxed = {"border": Border(left=thin, right=thin, top=thin, bottom=thin)}
    wb = openpyxl.Workbook()
    ws = wb.active
    for coordinate, rotation in (("B2", 90), ("B4", 255), ("B6", 45)):
        ws[coordinate] = "a"
        ws[coordinate].alignment = Alignment(textRotation=rotation)
        ws[coordinate].border = boxed["border"]
    ws.merge_cells("B6:C6")
    rendered = sheetpeek.read(save(wb, tmp_path / "upright.xlsx")).layout("excel")
    assert rendered.borders.slanted == ()


def test_a_slanted_box_keeps_the_columns_it_reaches(tmp_path):
    thin = Side("thin")
    boxed = {"border": Border(left=thin, right=thin, top=thin, bottom=thin)}
    for rotation, letters in ((15, "CDEF"), (45, "CDE"), (135, "ABC")):
        path = _book(
            tmp_path,
            [("C3", "Rotated text here", Alignment(textRotation=rotation), boxed)],
            f"{rotation}.xlsx",
        )
        rendered = sheetpeek.read(path).layout("excel")
        assert "".join(c.letter for c in rendered.columns) == letters, rotation


def test_html_draws_the_slanted_box_as_four_lines(tmp_path):
    from urllib.parse import unquote

    thin = Side("thin", color="FF0000C0")
    boxed = {"border": Border(left=thin, right=thin, top=thin, bottom=thin)}
    path = _book(tmp_path, [("C3", "Rotated text here", Alignment(textRotation=45), boxed)])
    rendered = sheetpeek.read(path).layout("excel")
    markup = unquote(sheetpeek.render(rendered, to="html", on_notice="ignore"))
    lines = re.findall(
        r"<line x1='([\d.]+)' y1='([\d.]+)' x2='([\d.]+)' y2='([\d.]+)' stroke='#0000C0'", markup
    )
    assert len(lines) == 4
    # 70 pt row and shift, 53 pt column: 93 and 71 px at 4/3, the box from x 0 to 164 px.
    left, right, top, bottom = (tuple(float(v) for v in line) for line in lines)
    assert left == (0.5, 93.5, 93.5, 0.5)
    assert right == (71.5, 93.5, 164.5, 0.5)
    assert top[1] == top[3] == 0.5
    assert bottom[1] == bottom[3] == 93.5
    assert not [n for n in sheetpeek.notices(rendered, "html") if n.code == "border.slanted"]


def _slant_text_captures():
    """Captured slant cases with a side drawn: (name, row height, ink edges)."""
    from pathlib import Path

    path = Path(__file__).parent / "captures" / "rotation.tsv"
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    header = next(line for line in lines if line[0] == "group")
    rows = [dict(zip(header, line, strict=False)) for line in lines if line[0] == "slant"]
    return [
        (row["name"], float(row["row_height"]), float(row["ink_left"]), float(row["ink_right"]))
        for row in rows
        if "fill" not in row["name"] and not row["name"].startswith(("90 ", "-90 "))
    ]


SLANT_BOUND = 1.5
"""Points, across the line, within which each slanted case's ink edges lie (median 0.6, 90th
percentile 1.1 over the 46 cases' 88 edges), but for ``SLANT_OUTLIERS``."""
SLANT_OUTLIERS = {
    # Top alignment between slanted sides, one case: the height is placed as without borders,
    # where this case is within 0.6 pt; Excel puts the text 2.2 to 2.5 pt further across.
    "45 lrtb vertical top": 2.75,
    # The unbordered placement is already 1 pt off at -60 degrees; the right edge is 1.9.
    "-60 lrtb": 2.25,
    # The insets are fitted at 11 pt; at 20 pt the left edge lies 1.7 pt inside.
    "45 lrtb 20": 2.0,
}


@pytest.mark.parametrize(("name", "height", "left", "right"), _slant_text_captures())
def test_text_between_slanted_sides_sits_where_excel_draws_it(tmp_path, name, height, left, right):
    # Captured (rotation suite, slant, 46 cases from 5 to 85 and -15 to -75 degrees): the ink's
    # edges, measured across the line, against the model's (the first and last letters' own
    # heights). Not compared: the right edge at 20 pt (xlformat measures the text narrower than
    # Excel), and in a 30 pt row at 45 degrees, where the row clips the text.
    from capture import rotation

    (case,) = [c for c in rotation.cases() + rotation.more_cases() if c.name == name]
    wb = openpyxl.Workbook()
    ws = wb.active
    rotation.fill(ws, [case])
    ws.row_dimensions[4].height = height  # fitted heights are tested on their own
    rendered = sheetpeek.read(save(wb, tmp_path / "slant.xlsx")).layout("excel")
    column = next(c for c in rendered.columns if c.letter == "E")
    (run,) = [r for r in _row(rendered, 4).runs if r.source == (4, 5)]
    lo, hi = _glyph_ink(run, column)
    across = abs(math.sin(math.radians(run.angle)))
    bound = SLANT_OUTLIERS.get(name, SLANT_BOUND)
    assert abs(lo - left) * across <= bound, (name, lo, left)
    if case.size == 11.0 and height != 30.0:
        assert abs(hi - right) * across <= bound, (name, hi, right)


def _turned_captures():
    from pathlib import Path

    path = Path(__file__).parent / "captures" / "rotation.tsv"
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    header = next(line for line in lines if line[0] == "group")
    rows = [dict(zip(header, line, strict=False)) for line in lines if line[0] != "group"]
    picked = [
        row
        for row in rows
        if row.get("group") == "turned"
        or (row.get("group"), row.get("name")) in {("other", "wrapped"), ("align", "justify")}
        or (row.get("group"), row.get("name")) == ("align", "distributed")
    ]
    return [
        (r["name"], float(r["row_height"]), float(r["ink_left"]), float(r["ink_right"]))
        for r in picked
    ]


TURNED_BOUND = 1.5
"""Points within which each wrapped angled case's ink edges lie (median 0.4 over 24 edges)."""
TURNED_OUTLIERS = {
    # The second line's end at a negative angle, one case: 1.7 pt.
    "wrapped -45 height 60": 2.0,
}


@pytest.mark.parametrize(("name", "height", "left", "right"), _turned_captures())
def test_wrapped_angled_text_sits_where_excel_draws_it(tmp_path, name, height, left, right):
    # Captured (rotation suite, turned; align and other at 45 degrees in 15 pt rows): lines as
    # long as fit between the row's top and bottom, a line apart toward the descent side, each
    # placed in height on its own; the row is not fitted to them. Ink edges within TURNED_BOUND.
    # In a 15 pt row each line holds one character and the text runs past the measured window
    # (Excel also gives a space its own line there, which sheetpeek drops), so only the left
    # edge is compared.
    from capture import rotation

    (case,) = [c for c in rotation.cases() + rotation.more_cases() if c.name == name]
    wb = openpyxl.Workbook()
    rotation.fill(wb.active, [case])
    rendered = sheetpeek.read(save(wb, tmp_path / "turned.xlsx")).layout("excel")
    column = next(c for c in rendered.columns if c.letter == "E")
    row = _row(rendered, 4)
    assert row.height_pt == height
    edges = [_glyph_ink(run, column) for run in row.runs if run.source == (4, 5)]
    bound = TURNED_OUTLIERS.get(name, TURNED_BOUND)
    assert abs(min(lo for lo, _ in edges) - left) <= bound, name
    if height > 15:
        assert abs(max(hi for _, hi in edges) - right) <= bound, name
