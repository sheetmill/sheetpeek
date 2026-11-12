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


def test_a_wrapped_angled_cell_is_drawn_on_one_line_and_reported(tmp_path):
    path = _book(tmp_path, [("B2", "two\nlines", Alignment(textRotation=45, wrap_text=True), {})])
    rendered = sheetpeek.read(path).layout("excel")
    (run,) = _row(rendered, 2).runs
    assert run.text == "twolines"
    (notice,) = [n for n in rendered.notices if n.code == "text.rotation_wrapped"]
    assert notice.outcome == "approximated"


def test_angled_cells_with_borders_are_reported_except_at_90(tmp_path):
    thin = Side("thin")
    boxed = {"border": Border(left=thin, right=thin, top=thin, bottom=thin)}
    path = _book(
        tmp_path,
        [
            ("B2", "a", Alignment(textRotation=45), boxed),
            ("B4", "b", Alignment(textRotation=90), boxed),
            ("B6", "c", Alignment(textRotation=255), boxed),
            ("B8", "d", Alignment(textRotation=45), {}),
        ],
    )
    rendered = sheetpeek.read(path).layout("excel")
    found = sheetpeek.notices(rendered, "html")
    (slanted,) = [n for n in found if n.code == "border.slanted"]
    assert (slanted.count, slanted.outcome) == (1, "approximated")
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
    sheet = sheetpeek.read(tmp_path / "rotation.xlsx")
    assert len(placed) > 90
    for coordinate, case in placed:
        row, column = int(coordinate[1:]), 5
        style = sheet.cell(row, column).style
        expected = 0 if case.rotation == 255 else _rotation(case.rotation)
        assert (style.rotation, style.stacked) == (expected, case.rotation == 255), case.name
        if case.borders == "lr":
            assert style.borders.left is not None
            assert style.borders.top is None
