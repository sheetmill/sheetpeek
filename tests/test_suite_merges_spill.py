"""The screen suite ``merges_spill`` (borders of merges and where text spills), without Excel.

ARCHITECTURE.md §4.12 ("Merges", and spilled text removing the vertical borders it crosses). The
workbook is built and read back with sheetpeek as built, and sheetpeek's resolved borders are the
suite's ``EXPECTED``. The measuring runs on a synthetic capture with known lines painted. The
comparison with ``tests/captures/merges_spill.tsv`` is skipped until that capture exists.
"""

from pathlib import Path

import openpyxl
import pytest
from capture import merges_spill as suite
from capture.screen.window import MARKER
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from tests.screen_canvas import Canvas

CAPTURE = Path(__file__).with_name("captures") / "merges_spill.tsv"
BORDERS = Path(__file__).with_name("captures") / "borders.tsv"
DRAWN_COVERAGE = 0.98
"""Share of an edge's length a drawn side must cover: the whole of it, less one device pixel of
an edge about 100 pixels long (the corners are already left out), for a stray antialiased pixel
where text or a meeting line touches the edge."""
ABSENT_COVERAGE = 0.5
"""Below this share no side is drawn: black text crossing a boundary is not a border color, and
a side Excel drew only between the glyphs would still cover more than half the row."""
DIAGONAL_DRAWN = 0.75
"""Of the four points looked at, at least three: one may fall where antialiasing pales the line
past the nearest-color naming."""
MAGENTA_DISTANCE = 120
"""No color the suite draws may be this close to the marker (Euclidean, per channel 0-255)."""


# --- the workbook


def _built(tmp_path):
    path = tmp_path / "merges_spill.xlsx"
    suite.build(path)
    return path


def test_the_workbook_reads_back_as_built(tmp_path):
    path = _built(tmp_path)
    sheet = sheetpeek.read(path)
    names = {"l": "left", "r": "right", "t": "top", "b": "bottom"}
    for address, placed in suite.CELLS.items():
        letters, row = coordinate_from_string(address)
        cell = sheet.cell(row, column_index_from_string(letters))
        borders = cell.style.borders
        for edge, style, color in placed.sides:
            side = getattr(borders, names[edge])
            assert (side.style, side.color) == (style, suite.INK[color]), (address, edge)
        if placed.diagonal is not None:
            style, color = placed.diagonal
            assert borders.diagonal == sheetpeek.Side(style, suite.INK[color]), address
            assert (borders.diagonal_up, borders.diagonal_down) == (True, True), address
        if placed.value is not None:
            assert cell.value == placed.value, address
    wb = openpyxl.load_workbook(path)
    assert sorted(str(m) for m in wb.active.merged_cells.ranges) == sorted(suite.MERGES)


def test_expected_is_what_sheetpeek_draws(tmp_path):
    assert suite.sheetpeek_drawn(_built(tmp_path)) == suite.EXPECTED


def test_the_spill_cases_spill_as_their_docstrings_say(tmp_path):
    rendered = sheetpeek.read(_built(tmp_path)).layout("excel")
    crossed = {row.index: row.crossed[:6] for row in rendered.rows}
    # Boundaries C|D, D|E, E|F, F|G, G|H, H|I.
    assert crossed[14] == (True, True, True, False, False, False)
    assert crossed[16] == (True, True, False, False, False, False)
    assert crossed[18] == (False, True, True, False, False, False)
    assert crossed[20] == (False, True, True, False, False, False)
    assert crossed[22] == (True, False, False, False, False, False)


def test_every_record_names_a_case_and_its_cells_are_referenced():
    refs = set(suite.references())
    keys = [(cell, edge) for _, cell, edge, _ in suite.EDGES]
    keys += [(box, direction) for _, box, direction, _, _ in suite.DIAGONALS]
    assert len(keys) == len(set(keys)) == len(suite.EXPECTED)
    for case, cell, _, _ in suite.EDGES:
        assert case in suite.CASES
        letters, row = coordinate_from_string(cell)
        assert f"{letters}{suite.REF_ROW}" in refs
        assert f"{suite.REF_COLUMN}{row}" in refs
    assert {case for case, *_ in suite.EDGES} | {case for case, *_ in suite.DIAGONALS} == set(
        suite.CASES
    )


def test_no_case_draws_near_the_marker_color():
    marker = suite._rgb(MARKER)
    for name, hex6 in {**suite.INK, **suite.FILLS}.items():
        rgb = suite._rgb(hex6)
        distance = sum((a - b) ** 2 for a, b in zip(rgb, marker, strict=True)) ** 0.5
        assert distance > MAGENTA_DISTANCE, name


# --- the measuring, on a synthetic capture


def _canvas():
    wb = openpyxl.Workbook()
    suite.fill(wb.active)
    return Canvas(wb.active, suite.references())


def _px(canvas, cell):
    """A grid cell's box in image pixels on the canvas, from the reference cells."""
    letters, row = coordinate_from_string(cell)
    left, _, right, _ = canvas.box(f"{letters}{suite.REF_ROW}")
    _, top, _, bottom = canvas.box(f"{suite.REF_COLUMN}{row}")
    return left, top, right, bottom


def _record(canvas, cell, edge):
    (found,) = [e for e in suite.EDGES if e[1:3] == (cell, edge)]
    return suite.measure_edge(canvas.capture(), *found)


def test_a_side_is_measured_by_its_color_rows_and_pattern():
    canvas = _canvas()
    left, top, right, bottom = _px(canvas, "C4")
    # A thick red top: 3 points from 1 point before the boundary (offsets -1 to 1.5).
    canvas.paint(left, top - 2, right, top + 4, suite._rgb(suite.RED))
    record = _record(canvas, "C4", "top")
    assert record["color"] == "red"
    assert record["coverage"] == "1.000"
    assert record["rows"] == "-1 -0.5 0 0.5 1 1.5"
    assert record["pattern"] == "solid"
    # A dashed thin blue right side on the gridline: 3 points on, 1 off.
    for y in range(top, bottom, 8):
        canvas.paint(right, y, right + 2, y + 6, suite._rgb(suite.BLUE))
    record = _record(canvas, "C4", "right")
    assert (record["color"], record["rows"], record["pattern"]) == ("blue", "0 0.5", "on 3 off 1")


def test_a_double_side_keeps_its_two_rows_apart():
    canvas = _canvas()
    left, top, _, bottom = _px(canvas, "C5")
    green = suite._rgb(suite.GREEN)
    canvas.paint(left - 2, top, left, bottom, green)
    canvas.paint(left + 2, top, left + 4, bottom, green)
    record = _record(canvas, "C5", "left")
    assert (record["color"], record["rows"]) == ("green", "-1 -0.5 1 1.5")


def test_text_across_a_boundary_is_no_side_and_the_boundary_shows_what_lies_under_it():
    canvas = _canvas()
    _left, top, right, bottom = _px(canvas, "C14")
    # Black text crossing C|D in the middle of the row, nothing else.
    middle = (top + bottom) // 2
    canvas.paint(right - 8, middle - 6, right + 8, middle + 6, (0, 0, 0))
    record = _record(canvas, "C14", "right")
    assert record.get("color", "") == ""
    assert record["coverage"] == "0"
    assert record.get("rows", "") == ""
    assert record["shows"] == "white"
    # A side removed only where the text is would still show as a side, partly covered.
    canvas.paint(right, top, right + 2, middle - 6, suite._rgb(suite.BLUE))
    canvas.paint(right, middle + 6, right + 2, bottom, suite._rgb(suite.BLUE))
    record = _record(canvas, "C14", "right")
    assert record["color"] == "blue"
    assert float(record["coverage"]) < DRAWN_COVERAGE


def test_a_fill_and_the_gridline_are_named_on_the_boundary():
    canvas = _canvas()
    left, top, right, bottom = _px(canvas, "D16")
    canvas.paint(left, top, right + 2, bottom + 2, suite._rgb(suite.SKY))
    assert _record(canvas, "D16", "top")["shows"] == "sky"
    assert _record(canvas, "D16", "right")["shows"] == "sky"
    left, top, right, bottom = _px(canvas, "D20")
    canvas.paint(left, top, left + 2, bottom, (212, 212, 212))
    record = _record(canvas, "D20", "left")
    assert record["shows"] == "grid"
    assert record["coverage"] == "0"


def test_a_diagonal_is_found_along_its_own_box_only():
    canvas = _canvas()
    left, top, _, _ = _px(canvas, "H8")
    _, _, right, bottom = _px(canvas, "J9")
    red = suite._rgb(suite.RED)
    steps = 2 * (right - left)
    for k in range(steps):  # the merge's up diagonal, 2 pixels thick
        x = left + k * (right - left) / steps
        y = bottom - k * (bottom - top) / steps
        canvas.paint(round(x), round(y) - 1, round(x) + 1, round(y) + 1, red)
    cap = canvas.capture()
    found = {(d[1], d[2]): suite.measure_diagonal(cap, *d)["coverage"] for d in suite.DIAGONALS}
    assert found["H8:J9", "up"] == "1"
    assert found["H8:J9", "down"] == "0"
    assert found["H8:H8", "up"] == found["H8:H8", "down"] == "0"
    assert found["C11:D12", "up"] == "0"


def test_every_edge_and_diagonal_is_measured_on_a_blank_capture():
    records = suite.measure(_canvas().capture())
    assert len(records) == len(suite.EDGES) + len(suite.DIAGONALS)
    assert all(r["coverage"] == "0" for r in records)
    assert {r["shows"] for r in records if r["kind"] == "edge"} == {"white"}


# --- the comparison with Excel


def _table(path):
    lines = [ln for ln in path.read_text(encoding="utf-8").splitlines() if ln[:2] != "# "]
    names = lines[0].split("\t")
    return [dict(zip(names, line.split("\t"), strict=False)) for line in lines[1:]]


@pytest.fixture(scope="module")
def captured():
    if not CAPTURE.exists():
        pytest.skip("not captured yet: tests/captures/merges_spill.tsv")
    return _table(CAPTURE)


def _geometry():
    """The offsets the suite ``borders`` captured for each zoom, style and edge, as one string."""
    found: dict[tuple[str, str, str], list[str]] = {}
    for row in _table(BORDERS):
        found.setdefault((row["zoom"], row["style"], row["edge"]), []).append(row["offset"])
    return {key: " ".join(offsets) for key, offsets in found.items()}


def _color_name(hex6):
    return {v: k for k, v in suite.INK.items()}[hex6]


@pytest.mark.excel
def test_each_edge_draws_the_side_sheetpeek_resolves(captured):
    # Merges: the outline per segment, the anchor's sides on its own segments, no inner sides, a
    # neighbor competing per segment. Spill: crossed vertical sides gone over the whole row,
    # horizontal sides and the blocking cell's side kept.
    geometry = _geometry()
    wrong = []
    for record in (r for r in captured if r["kind"] == "edge"):
        expected = suite.EXPECTED[record["cell"], record["edge"]]
        key = f"{record['zoom']} {record['cell']} {record['edge']} ({record['statement']})"
        if expected == "none":
            if record["rows"] or float(record["coverage"]) >= ABSENT_COVERAGE:
                wrong.append(f"{key}: drawn {record['color']} {record['rows']}, sheetpeek none")
            continue
        style, color = expected.split()
        rows = geometry[record["zoom"], style, record["edge"]]
        found = (record["color"], record["rows"], record["pattern"])
        if found != (_color_name(color), rows, "solid") or (
            float(record["coverage"]) < DRAWN_COVERAGE
        ):
            wrong.append(f"{key}: {found} {record['coverage']}, sheetpeek {expected} rows {rows}")
    assert not wrong, "\n".join(wrong)


@pytest.mark.excel
def test_where_no_side_is_drawn_the_boundary_shows_what_sheetpeek_lays_out(captured, tmp_path):
    # A crossed vertical boundary and one inside a merge show no gridline (white, or a fill
    # where one covers it: the cell right or below paints over its neighbor's).
    rendered = sheetpeek.read(_built(tmp_path)).layout("excel")
    rows = {r.index: r for r in rendered.rows}
    fills = {v: k for k, v in suite.FILLS.items()}

    columns = [c.index for c in rendered.columns]

    def cell(column, row):
        """The laid-out cell, or ``None`` outside the used range (an empty cell there)."""
        if row not in rows:
            return None
        return next((c for c in rows[row].cells if c.column == column), None)

    def shows(address, edge):
        """The fill, white or gridline on the boundary; ``first`` is before it, ``second`` after."""
        letters, row = coordinate_from_string(address)
        column = column_index_from_string(letters)
        inside = False
        if edge in ("left", "right"):
            q = column if edge == "left" else column + 1
            first, second = cell(q - 1, row), cell(q, row)
            if q - 1 in columns and q in columns:
                inside = rows[row].crossed[columns.index(q) - 1]
        else:
            r = row if edge == "top" else row + 1
            first, second = cell(column, r - 1), cell(column, r)
        if first is not None and second is not None:
            inside = inside or (first.merge is not None and first.merge == second.merge)
        fill = (second and second.fill) or (first and first.fill)
        if fill is not None:
            return fills[fill]
        return "white" if inside else "grid"

    wrong = []
    for record in (r for r in captured if r["kind"] == "edge"):
        if suite.EXPECTED[record["cell"], record["edge"]] != "none":
            continue
        expected = shows(record["cell"], record["edge"])
        if record["shows"] != expected:
            wrong.append(
                f"{record['zoom']} {record['cell']} {record['edge']}: shows {record['shows']}, "
                f"sheetpeek {expected}"
            )
    assert not wrong, "\n".join(wrong)


NAMED_DIAGONALS = {
    ("100", "H8:H8"): "the merge's line passes 1.75 to 2.25 pt from the anchor box's own "
    "diagonal at its 35 and 45 % points, within the search reach at zoom 100",
    ("200", "H8:J9"): "Excel drew only the left half of the merge's diagonals, from its left "
    "corners to its center, at zoom 200 (full at 100); one window, cause not found",
}
"""Records not compared, each with its reason."""


@pytest.mark.excel
def test_diagonals_run_across_the_box_sheetpeek_draws_them_in(captured):
    # A merge's diagonals are its anchor's, across the whole merge; a covered cell's are not drawn.
    wrong = []
    for record in (r for r in captured if r["kind"] == "diagonal"):
        if (record["zoom"], record["cell"]) in NAMED_DIAGONALS:
            continue
        drawn = suite.EXPECTED[record["cell"], record["edge"]] == "drawn"
        coverage = float(record["coverage"])
        if (coverage >= DIAGONAL_DRAWN) != drawn or (not drawn and coverage > 0):
            wrong.append(f"{record['zoom']} {record['cell']} {record['edge']}: {coverage}")
    assert not wrong, "\n".join(wrong)
