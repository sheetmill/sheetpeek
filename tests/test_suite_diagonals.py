"""The screen suite ``diagonals`` (diagonal borders, ARCHITECTURE.md §4.12), without Excel.

Its workbooks are built and read back with sheetpeek as built, and what sheetpeek's HTML draws
for each case (the SVG line of its border layer) is the suite's ``EXPECTED``. Its measuring runs
on a synthetic capture with known lines painted. The comparison with
``tests/captures/diagonals.tsv`` is skipped until that capture exists.
"""

import math
import re
import statistics
from collections import Counter

import openpyxl
import pytest
from capture import diagonals as suite
from capture.screen.window import MARKER
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from tests.capture_report import REPORT
from tests.screen_canvas import Canvas

CAPTURE = suite.OUTPUT


def _rc(cell):
    letters, row = coordinate_from_string(cell)
    return row, column_index_from_string(letters)


# --- what sheetpeek draws


def _unescape(uri):
    """The SVG inside a ``url("data:image/svg+xml,...")``, as ``svg_uri`` escaped it."""
    svg = uri.split("data:image/svg+xml,", 1)[1].rsplit('")', 1)[0]
    for char, code in (("#", "%23"), ("<", "%3C"), (">", "%3E"), ('"', "%22"), ("%", "%25")):
        svg = svg.replace(code, char)
    return svg


def parse(declarations, width, height):
    """A diagonal's SVG as the data ``EXPECTED`` holds."""
    svg = _unescape(declarations)
    lines = re.findall(r"<line ([^/]*)/>", svg)
    found = {"box": (width, height), "up": False, "down": False, "double": False}
    for line in lines:
        attributes = dict(re.findall(r"([a-z0-9-]+)='([^']*)'", line))
        found["up" if attributes["y1"] != "0" else "down"] = True
        found["stroke"] = float(attributes["stroke-width"])
        found["dashes"] = attributes.get("stroke-dasharray", "")
        found["double"] = found["double"] or "transform" in attributes
    return found


def drawn(path):
    """What sheetpeek's HTML draws for each case of the workbook at ``path``, as ``EXPECTED``."""
    group = path.stem.removeprefix("diagonals_")
    page = sheetpeek.render(path, "html", on_notice=lambda notice: None)
    boxes = re.findall(r'<i class="(sp-b-[a-z]+)" style="([^"]*)"></i>', page)
    sheet = sheetpeek.read(path)
    ordered = sorted(suite.cases(group), key=lambda c: _rc(c.cell))
    with_lines = []
    for case in ordered:
        borders = sheet.cell(*_rc(case.cell)).style.borders
        if borders is not None and borders.diagonal is not None:
            with_lines.append(case)
    assert len(with_lines) == len(boxes), group  # the only lines on the sheet, row by row
    found = {case.cell: None for case in suite.cases(group)}
    for case, (name, style) in zip(with_lines, boxes, strict=True):
        body = re.search(rf"\.sheetpeek \.{name} \{{ ([^}}]*) \}}", page)[1]
        width = int(re.search(r"width:(\d+)px", style)[1])
        height = int(re.search(r"height:(\d+)px", style)[1])
        found[case.cell] = parse(body, width, height)
    return found


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    folder = tmp_path_factory.mktemp("diagonals")
    suite.build(folder)
    return folder


# --- the workbooks


def test_the_diagonal_workbooks_read_back_as_built(built):
    assert len(suite.CASES) == 2 * 13 * 3 + 9
    for group in suite.GROUPS:
        path = built / f"diagonals_{group}.xlsx"
        sheet = sheetpeek.read(path)
        for case in suite.cases(group):
            borders = sheet.cell(*_rc(case.cell)).style.borders
            if case.style is None or not (case.up or case.down):
                # a diagonal is read only with a style and a direction
                assert borders is None or borders.diagonal is None, case.name
                continue
            assert borders.diagonal.style == case.style, case.name
            assert (borders.diagonal_up, borders.diagonal_down) == (case.up, case.down)
            assert (borders.left, borders.right, borders.top, borders.bottom) == (None,) * 4
            if case.text:
                assert sheet.cell(*_rc(case.cell)).value == suite.TEXT
        assert drawn(path) == suite.EXPECTED[group], group


def test_every_style_is_drawn_in_each_direction_and_size():
    for group in suite.GROUPS:
        found = {(c.style, c.direction) for c in suite.cases(group)}
        for style in suite.STYLES:
            for direction in suite.DIRECTIONS:
                assert (style, direction) in found, (group, style, direction)


def test_every_case_has_a_cell_of_its_own_between_empty_ones():
    for group in suite.GROUPS:
        found = suite.cases(group)
        cells = {c.cell for c in found}
        assert len(cells) == len(found), group
        for case in found:
            row, column = _rc(case.cell)
            for neighbor in (
                f"{get_column_letter(column + 1)}{row}",
                f"{get_column_letter(column)}{row + 1}",
            ):
                assert neighbor not in cells, case.name
            assert case.cell not in suite.MARKERS
    last_row, last_column = _rc(suite.MARKERS[1])
    assert all(_rc(c.cell)[0] < last_row and _rc(c.cell)[1] < last_column for c in suite.CASES)


def test_no_case_draws_near_the_marker_color():
    marker = tuple(int(MARKER[i : i + 2], 16) for i in (0, 2, 4))
    for color in (suite.BLACK, suite.TEXT_COLOR):
        rgb = tuple(int(color[2:][i : i + 2], 16) for i in (0, 2, 4))
        assert max(abs(a - b) for a, b in zip(rgb, marker, strict=True)) > 96


def test_the_cells_read_give_every_case_its_geometry():
    for group in suite.GROUPS:
        ws = openpyxl.Workbook().active
        suite.fill(ws, group)
        every = Canvas(ws, [c.cell for c in suite.cases(group)]).capture()
        read = Canvas(ws, suite.read_cells(group)).capture()
        assert len(suite.read_cells(group)) < len(suite.cases(group)) // 2
        assert suite.complete(read, group).geometry == every.geometry, group


# --- measuring, on a synthetic capture


def _canvas(group="large"):
    ws = openpyxl.Workbook().active
    suite.fill(ws, group)
    return Canvas(ws, [c.cell for c in suite.cases(group)])


def _case(group, cell):
    return next(c for c in suite.cases(group) if c.cell == cell)


def _paint_line(canvas, cell, line, rows=(-1, 0), period=None, rgb=(0, 0, 0)):
    """Paint image rows ``rows`` (from the corner-to-corner line) in each column of the cell.

    ``period``, ``(on, off)`` in image columns, leaves gaps along x.
    """
    x0, y0, x1, y1 = canvas.box(cell)
    for x in range(x0, x1):
        if period is not None and (x - x0) % sum(period) >= period[0]:
            continue
        f = (x + 0.5 - x0) / (x1 - x0)
        y = y1 - f * (y1 - y0) if line == "up" else y0 + f * (y1 - y0)
        for dy in rows:
            canvas.paint(x, round(y) + dy, x + 1, round(y) + dy + 1, rgb)


def _records(canvas, group, cell):
    return [r for r in suite.measure(canvas.capture(), group) if r["cell"] == cell]


def test_a_thin_line_reads_corner_to_corner_one_point_wide():
    canvas = _canvas()
    _paint_line(canvas, "C4", "up")
    (record,) = _records(canvas, "large", "C4")
    assert (record["line"], record["ink"], record["pattern"]) == ("up", "000000", "solid")
    assert (record["start"], record["end"]) == ("0", "0")
    # two image rows a column: 1 pt down the column; the rows sit about centered on the line
    assert record["coverage"] == "1"
    assert abs(float(record["offset"])) <= 0.25
    assert record["bands"] == "1"
    assert set(suite.COLUMNS) >= set(record)


def test_a_dashed_line_reads_as_its_runs_along_x():
    canvas = _canvas()
    _paint_line(canvas, "I6", "down", period=(6, 2))  # dashed down: on 3 off 1 pt along x
    (record,) = _records(canvas, "large", "I6")
    assert record["pattern"] == "on 3 off 1"
    runs = record["runs"].split()
    assert len(runs) >= 40
    # the partial dash at the start is dropped: whole gaps and dashes alternate from there
    assert set(zip(runs[0::2], runs[1::2], strict=True)) == {("off", "1"), ("on", "3")}
    assert record["coverage"] == "1"


def test_a_double_line_reads_as_two_bands():
    canvas = _canvas()
    _paint_line(canvas, "I4", "up", rows=(-3, -2, 1, 2))  # double up
    (record,) = _records(canvas, "large", "I4")
    assert record["bands"] == "2"
    assert record["coverage"] == "2"
    assert abs(float(record["offset"])) <= 0.25


def test_a_thick_line_reads_its_width_and_offset():
    canvas = _canvas()
    _paint_line(canvas, "G4", "up", rows=(0, 1, 2, 3, 4, 5))  # 3 pt, centered 2.5 px below
    (record,) = _records(canvas, "large", "G4")
    assert record["coverage"] == "3"
    assert abs(float(record["offset"]) - 1.25) <= 0.25


def test_both_lines_of_a_cell_are_measured_apart():
    canvas = _canvas()
    _paint_line(canvas, "C8", "up")
    _paint_line(canvas, "C8", "down", rows=(-2, -1, 0, 1))  # the down line 2 pt
    records = {r["line"]: r for r in _records(canvas, "large", "C8")}
    assert set(records) == {"up", "down"}
    assert (records["up"]["coverage"], records["down"]["coverage"]) == ("1", "2")
    for record in records.values():
        assert (record["start"], record["end"], record.get("pattern")) == ("0", "0", None)


def test_a_small_cell_reads_its_line_too():
    canvas = _canvas("small")
    _paint_line(canvas, "E6", "down", rows=(-2, -1, 0, 1))  # medium down
    (record,) = _records(canvas, "small", "E6")
    assert (record["coverage"], record["pattern"], record["bands"]) == ("2", "solid", "1")
    assert record["height"] == "15"


def test_text_and_the_line_crossing_it_are_told_apart():
    canvas = _canvas()
    cell = "C16"  # thin both, text centered
    _paint_line(canvas, cell, "up")
    _paint_line(canvas, cell, "down")
    x0, y0, x1, y1 = canvas.box(cell)
    cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
    canvas.paint(cx - 40, cy - 10, cx - 30, cy + 10, (0, 0, 192))  # a glyph over the lines
    canvas.paint(cx + 10, cy - 10, cx + 30, cy + 10, (0, 0, 192))
    records = {r["line"]: r for r in _records(canvas, "large", cell)}
    for record in records.values():
        left, top, right, bottom = (float(v) for v in record["text_box"].split())
        assert (right - left, bottom - top) == (35, 10)
        counts = dict(zip(*[iter(record["crossing"].split())] * 2, strict=True))
        assert int(counts["blue"]) > 0
        assert int(counts["black"]) > 0
        assert record["start"] == "0"  # the line is still found outside the text
        assert record["ink"] == "000000"  # the text's blue is not the line's ink


def test_every_case_is_measured_on_a_blank_capture():
    for group in suite.GROUPS:
        cap = _canvas(group).capture()
        records = suite.measure(cap, group)
        assert len(records) == sum(len(c.lines) for c in suite.cases(group))
        for record in records:
            assert {"ink", "start", "coverage", "pattern"}.isdisjoint(record), record["name"]


def test_write_leaves_out_what_was_not_captured(tmp_path, monkeypatch):
    monkeypatch.setattr(suite, "BUILD_DIR", tmp_path)
    monkeypatch.setattr(suite, "OUTPUT", tmp_path / "out.tsv")
    suite.write()
    lines = (tmp_path / "out.tsv").read_text(encoding="utf-8").splitlines()
    assert lines == ["# suite: diagonals", "\t".join(suite.COLUMNS)]


# --- the comparison with Excel (suite diagonals)

WIDTH_TOLERANCE = (0.35, 0.25)
"""A line's measured thickness may differ from the expected by 0.35 screen points plus a
quarter of it; provisional until the capture's spread is known. The thickness is the darkness
summed down an image column times the cosine of the line's angle: an antialiased edge blended in
linear light reads lighter in sRGB levels, so a 1-point line may read up to a third thin, and a
column is half a device pixel from the line's true center on average."""
DASH_TOLERANCE = 1.0
"""Screen points a dash or gap along the line may differ by: the columns are counted in device
pixels (half a screen point) along x, a slanted dash's antialiased ends add or lose a column at
each end, and dividing by the cosine stretches the error by up to 8 %."""
END_TOLERANCE = 1.0
"""Screen points a line's ends may lie from where sheetpeek ends it, beyond half its stroke: the
first and last inked columns are found by a luma threshold on antialiased pixels."""
OFFSET_TOLERANCE = 0.5
"""Screen points the ink's center may lie from sheetpeek's line: a pixel's rounding either way."""
SOLID = ("thin", "medium", "thick", "double", "hair")


@pytest.fixture(scope="module")
def captured():
    if not CAPTURE.exists():
        pytest.skip("not captured yet: run python -m capture.diagonals run")
    lines = [ln for ln in CAPTURE.read_text(encoding="utf-8").splitlines() if ln[:2] != "# "]
    names = lines[0].split("\t")
    return [dict(zip(names, line.split("\t"), strict=False)) for line in lines[1:]]


def _screen(row, key):
    """A measured length in screen points (sheet points times the zoom)."""
    return float(row[key]) * int(row["zoom"]) / 100


def _cosine(row):
    w, h = float(row["width"]), float(row["height"])
    return w / math.hypot(w, h)


def _expected(row):
    return suite.EXPECTED[row["group"]][row["cell"]]


def _spread(label, errors):
    if errors:
        ordered = sorted(errors)
        REPORT[f"diagonals {label}"] = (
            f"diagonals {label}: {len(ordered)} lines, median {statistics.median(ordered):.2f}, "
            f"90th percentile {ordered[int(0.9 * (len(ordered) - 1))]:.2f}, "
            f"max {ordered[-1]:.2f} screen points"
        )


def _runs(row):
    """The measured runs from the first whole dash, ``on 3 off 1 ...`` as ``[3.0, 1.0, ...]``."""
    words = row["runs"].split()
    if words[:1] == ["off"]:
        words = words[2:]
    return [float(v) for v in words[1::2]]


def _dashed(row):
    """Whether a row's line has dashes: at least one whole dash and gap between its ends."""
    return len(_runs(row)) >= 2


def _along(row):
    """The measured runs (dash, gap, ... between the partial ends) along the line, in screen
    points."""
    scale = int(row["zoom"]) / 100 / _cosine(row)
    return [v * scale for v in _runs(row)]


def _cycle_matches(measured, dashes, tolerance):
    """Whether ``measured`` (on, off, ...) follows ``dashes`` from one of its dashes, each run
    within ``tolerance``."""
    if len(measured) < len(dashes):
        return False
    return any(
        all(abs(a - dashes[(k + i) % len(dashes)]) <= tolerance for i, a in enumerate(measured))
        for k in range(0, len(dashes), 2)
    )


@pytest.mark.excel
def test_a_diagonal_runs_corner_to_corner(captured):
    # sheetpeek's line runs from the cell's corner on the gridline before it to the far
    # gridline's pixel, 1 screen point past the far boundary; only the solid styles, since a
    # dashed line may end in a gap.
    bad, errors = [], []
    for row in captured:
        expected = _expected(row)
        if expected is None or row["style"] not in SOLID or row["start"] == "":
            continue
        reach = expected["stroke"] / 2 + (1 if expected["double"] else 0)
        for key, where in (("start", 0.0), ("end", 1.0)):
            error = abs(_screen(row, key) - where)
            errors.append(error)
            if error > END_TOLERANCE + reach:
                bad.append((row["group"], row["cell"], row["zoom"], row["line"], key, row[key]))
    _spread("ends", errors)
    assert errors
    assert not bad, bad


def _thickness(row):
    return _screen(row, "coverage") * _cosine(row)


@pytest.mark.excel
def test_each_style_draws_its_width(captured):
    # The solid styles and the dashed ones whose dashes are at least 3 points long; a double
    # line is two 1-point lines, 2 bands.
    bad, errors = [], []
    for row in captured:
        expected = _expected(row)
        if expected is None or row["coverage"] == "":
            continue
        if any(float(v) < 3 for v in expected["dashes"].split()):
            continue
        width = 2.0 if expected["double"] else expected["stroke"]
        error = abs(_thickness(row) - width)
        errors.append(error)
        absolute, relative = WIDTH_TOLERANCE
        if error > absolute + relative * width or row["bands"] != (
            "2" if expected["double"] else "1"
        ):
            bad.append(
                (row["group"], row["cell"], row["zoom"], row["line"], row["coverage"], row["bands"])
            )
    _spread("widths", errors)
    assert errors
    assert not bad, bad


@pytest.mark.excel
def test_a_hair_diagonal_is_drawn_as_thin(captured):
    # Excel against itself: in each group, zoom and direction, hair as thin (unverified).
    rows = {
        (r["group"], r["zoom"], r["direction"], r["line"], r["style"]): r
        for r in captured
        if r["name"] == f"{r['style']} {r['direction']}"  # not the text cases
    }
    checked = 0
    for (group, zoom, direction, line, style), hair in rows.items():
        thin = rows.get((group, zoom, direction, line, "thin"))
        if style != "hair" or thin is None:
            continue
        checked += 1
        assert hair["pattern"] == thin["pattern"], (group, zoom, direction)
        if hair["coverage"] and thin["coverage"]:
            absolute, relative = WIDTH_TOLERANCE
            assert abs(_thickness(hair) - _thickness(thin)) <= absolute + relative, (group, zoom)
    assert checked


@pytest.mark.excel
def test_each_style_draws_its_dashes(captured):
    bad, checked = [], 0
    for row in captured:
        expected = _expected(row)
        if expected is None or row["direction"] not in ("up", "down") or row["pattern"] == "":
            continue
        if row["name"] != f"{row['style']} {row['direction']}":
            continue  # the text cases
        checked += 1
        if not expected["dashes"]:
            if row["pattern"] != "solid":
                bad.append((row["group"], row["cell"], row["zoom"], row["pattern"]))
            continue
        dashes = [float(v) for v in expected["dashes"].split()]
        if not _dashed(row):
            bad.append((row["group"], row["cell"], row["zoom"], row["pattern"], dashes))
            continue
        if not _cycle_matches(_along(row), dashes, DASH_TOLERANCE):
            bad.append((row["group"], row["cell"], row["zoom"], row["pattern"], dashes))
    assert checked
    assert not bad, bad


def _period(measured):
    """The shortest leading runs that the rest of ``measured`` repeats within the tolerance."""
    for k in range(2, len(measured) // 2 + 1, 2):
        if _cycle_matches(measured, measured[:k], DASH_TOLERANCE):
            return measured[:k]
    return None


@pytest.mark.excel
def test_lines_are_fixed_in_screen_points_at_zoom_200(captured):
    # Excel against itself: a line's width and dashes along it are the same in screen points at
    # zoom 100 and 200, as sheetpeek draws them (CSS px at any page scale). Not the text cases,
    # whose text interrupts the line's columns differently at each zoom.
    at = {}
    for row in captured:
        if row["name"] != f"{row['style']} {row['direction']}":
            continue
        at.setdefault((row["group"], row["cell"], row["line"]), {})[row["zoom"]] = row
    checked, bad = 0, []
    for key, zooms in at.items():
        if "100" not in zooms or "200" not in zooms:
            continue
        low, high = zooms["100"], zooms["200"]
        if low["coverage"] and high["coverage"]:
            checked += 1
            absolute, relative = WIDTH_TOLERANCE
            if abs(_thickness(low) - _thickness(high)) > absolute + relative * _thickness(low):
                bad.append((key, "width", low["coverage"], high["coverage"]))
        if _dashed(low) and _dashed(high):
            period = _period(_along(low))
            if period is None or not _cycle_matches(_along(high), period, DASH_TOLERANCE):
                bad.append((key, "dashes", low["pattern"], high["pattern"]))
    assert checked
    assert not bad, bad


def _sheetpeek_offset(row):
    """Where sheetpeek's line crosses the cell's middle column, from the corner-to-corner line.

    Its SVG runs over the box of the cell and its far gridline, 1 screen point wider and taller,
    from corner to corner of that box (screen points, down positive).
    """
    zoom = int(row["zoom"]) / 100
    w, h = float(row["width"]) * zoom, float(row["height"]) * zoom
    x = w / 2
    if row["line"] == "up":
        return (h + 1) * (1 - x / (w + 1)) - h / 2
    return (h + 1) * x / (w + 1) - h / 2


@pytest.mark.excel
def test_the_line_lies_where_sheetpeek_draws_it(captured):
    bad, errors = [], []
    for row in captured:
        expected = _expected(row)
        if expected is None or row["offset"] == "":
            continue
        error = abs(_screen(row, "offset") - _sheetpeek_offset(row))
        errors.append(error)
        if error > OFFSET_TOLERANCE:
            bad.append((row["group"], row["cell"], row["zoom"], row["line"], row["offset"]))
    _spread("offsets", errors)
    assert errors
    assert not bad, bad


@pytest.mark.excel
def test_text_and_the_line_both_show(captured):
    # sheetpeek draws the border layer under the text: where the line's path crosses the text,
    # the text's blue shows on it, and the line shows beside the text.
    checked, shown = 0, Counter()
    for row in captured:
        if not row["name"].endswith(("text", "text centered")):
            continue
        checked += 1
        assert row["text_box"], row["name"]
        assert row["start"], row["name"]  # the line is drawn
        assert row["ink"], row["name"]
        counts = {k: int(v) for k, v in zip(*[iter(row["crossing"].split())] * 2, strict=True)}
        shown.update(counts)
        if counts["blue"] + counts["black"] >= 5:
            assert counts["blue"] > 0, (row["name"], row["line"], row["zoom"], row["crossing"])
    REPORT["diagonals crossing"] = (
        f"diagonals: on the line's path through the text, {checked} lines: "
        + ", ".join(f"{k} {v}" for k, v in sorted(shown.items()))
    )
    assert checked


@pytest.mark.excel
def test_without_a_style_or_a_direction_nothing_is_drawn(captured):
    checked = 0
    for row in captured:
        if row["direction"] != "none" and row["style"]:
            continue
        checked += 1
        assert suite.EXPECTED[row["group"]][row["cell"]] is None
        assert row["start"] == row["ink"] == "", (row["name"], row["line"])
    assert checked
