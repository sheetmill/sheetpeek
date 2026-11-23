"""The capture suite ``row_bottoms`` (which cells set a row's shared bottom baseline,
ARCHITECTURE.md §3.6), without Excel.

Its workbooks are built and read back with sheetpeek; sheetpeek's layout of each case takes the
candidate ``SHEETPEEK`` names, among the suite's ``PREDICTIONS``, written before the capture. The
comparison with ``tests/captures/row_bottoms.tsv`` is skipped until that capture exists.
"""

import openpyxl
import pytest
from capture import row_bottoms as suite
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

import sheetpeek
from tests.capture_report import REPORT
from tests.screen_canvas import SCALE, Canvas

BLACK = (0, 0, 0)


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    return suite.build(tmp_path_factory.mktemp("row_bottoms"))


def _at(cell):
    letters, row = coordinate_from_string(cell)
    return row, column_index_from_string(letters)


def laid_out(paths):
    """sheetpeek's row height and last baseline, from the row's (a merge's) top, by case."""
    found = {}
    for group, path in paths.items():
        rendered = sheetpeek.read(path).layout("excel")
        for case in suite.cases(group):
            r, c = _at(case.cell)
            row = next(x for x in rendered.rows if x.index == r)
            ys = sorted({t.y - row.y_pt for t in row.runs if t.source == (r, c)})
            found[case.name] = (row.height_pt, ys[-1])
    return found


def test_every_case_is_built_as_described(built):
    sheet = sheetpeek.read(built["cells"])
    for case in suite.cases("cells"):
        assert sheet.cell(*_at(case.cell)).value == suite.PROBE or "\n" in str(
            sheet.cell(*_at(case.cell)).value
        ), case.name
    assert sheet.cell(6, 5).value is None
    assert sheet.cell(6, 5).font.size == 24
    assert sheet.cell(13, 5).font.vert_align == "superscript"
    assert sheet.cell(14, 5).font.vert_align == "subscript"
    assert sheet.cell(16, 5).runs is not None
    assert {c.index for c in sheet.columns if c.hidden} == {7}
    assert {(m.min_row, m.min_col, m.max_row) for m in sheet.merges} >= {(11, 5, 12), (17, 3, 18)}


def test_every_case_has_its_candidates_written_before_the_capture():
    names = [case.name for case in suite.cases()]
    assert len(names) == len(set(names))
    assert set(suite.PREDICTIONS) == set(names) == set(suite.SHEETPEEK)
    for name, candidates in suite.PREDICTIONS.items():
        assert suite.SHEETPEEK[name] in candidates, name


def test_sheetpeek_takes_the_candidate_it_names(built):
    found = laid_out(built)
    for name, (height, last) in found.items():
        assert last == suite.PREDICTIONS[name][suite.SHEETPEEK[name]], name
        if name in suite.FITTED_HEIGHTS:
            assert height == suite.FITTED_HEIGHTS[name][suite.SHEETPEEK[name]], name


def test_every_case_has_a_cell_of_its_own_left_of_and_below_the_markers():
    for group in suite.GROUPS:
        found = suite.cases(group)
        cells = [c.cell for c in found]
        assert len(set(cells)) == len(cells), group
        first, second = suite.MARKERS[group]
        marker_column = _at(second)[1]
        for case in found:
            for cell in (case.cell, case.last or case.cell):
                row, column = _at(cell)
                assert column < marker_column, (group, cell)
                assert row > 3, (group, cell)
        assert first not in cells
        assert second not in cells


def test_no_case_draws_magenta():
    # Every case is text in the default color on white: no fill, and no font color but the
    # Normal style's theme text color (a rich-text cell's own font).
    for group in suite.GROUPS:
        ws = openpyxl.Workbook().active
        suite.fill(ws, group)
        for row in ws.iter_rows():
            for cell in row:
                color = cell.font.color
                if cell.value is not None or cell.has_style:
                    assert color is None or (color.type, color.theme) == ("theme", 1), cell
                    assert not (color is not None and color.type == "rgb"), cell
                assert cell.fill.fill_type is None, cell.coordinate


def test_a_merged_case_is_measured_from_the_merges_top():
    ws = openpyxl.Workbook().active
    case = next(c for c in suite.fill(ws, "cells") if c.last)
    canvas = Canvas(ws, [*suite.references("cells"), case.cell, case.last])
    x0, y0, _, _ = canvas.box(case.cell)
    _, _, _, y1 = canvas.box(case.last)
    canvas.paint(x0 + 6, y1 - 20, x0 + 40, y1 - 6, BLACK)
    measured = suite.measure(canvas.capture(), case)
    assert measured is not None
    assert float(measured["baselines"]) == (y1 - 6 - y0) / SCALE


def test_baselines_are_the_bottoms_of_the_line_bands():
    assert suite.baselines("16+10 32+10") == [26.0, 42.0]
    assert suite.baselines("") == []


# --- the comparison with the capture

PINNED: dict[str, str] = {}
"""Cases where sheetpeek's baseline differs from Excel's, each with its reason."""


def _captured():
    if not suite.OUTPUT.exists():
        pytest.skip("not captured yet: run python -m capture.row_bottoms run")
    lines = [line.split("\t") for line in suite.OUTPUT.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    return {r["name"]: r for r in (dict(zip(header, row, strict=True)) for row in body)}


@pytest.mark.excel
def test_each_case_matches_one_candidate():
    # The ink of "H" ends on a device pixel: half a point at zoom 100 on a 2x display.
    captured = _captured()
    report = {}
    for name, candidates in suite.PREDICTIONS.items():
        last = suite.baselines(captured[name]["line_tops"])[-1]
        report[name] = [k for k, v in candidates.items() if abs(v - last) <= 0.5]
        assert report[name], (name, last, candidates)
    REPORT["row_bottoms candidates"] = "row_bottoms: " + "; ".join(
        f"{name}: {'/'.join(which)}" for name, which in report.items()
    )


@pytest.mark.excel
def test_sheetpeek_places_each_probe_where_excel_does(built):
    captured = _captured()
    found = laid_out(built)
    differ = {
        name
        for name, (height, last) in found.items()
        if abs(suite.baselines(captured[name]["line_tops"])[-1] - last) > 0.5
        or (name in suite.FITTED_HEIGHTS and float(captured[name]["row_height"]) != height)
    }
    assert differ == set(PINNED), sorted(differ ^ set(PINNED))
