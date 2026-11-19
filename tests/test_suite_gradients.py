"""The screen suite ``gradients`` (gradient fills, ARCHITECTURE.md §4.12), without Excel.

Its workbooks are built and read back with sheetpeek as built, and what sheetpeek's HTML draws
for each case (the CSS ``linear-gradient`` or the path's SVG bands) is the suite's ``EXPECTED``.
Its measuring runs on a synthetic capture with known gradients painted. The comparison with
``tests/captures/gradients.tsv`` is skipped until that capture exists.
"""

import itertools
import math
import re
import statistics
import zipfile

import openpyxl
import pytest
from capture import gradients as suite
from capture.screen.window import MARKER
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from sheetpeek import Gradient
from sheetpeek.backends.html._patterns import gradient as gradient_css
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
    """A gradient's declarations as the data ``EXPECTED`` holds: its box, geometry and stops."""
    found = re.search(r"linear-gradient\(([-0-9.]+)deg,([^)]*)\)", declarations)
    if found:
        stops = tuple(
            (float(p) / 100, c) for c, p in re.findall(r"#([0-9A-F]{6}) ([-0-9.e]+)%", found[2])
        )
        return {"kind": "linear", "box": (width, height), "angle": float(found[1]), "stops": stops}
    svg = _unescape(declarations)
    polygons = re.findall(r"points='([^']*)'", svg)
    corners = [tuple(float(v) for v in p.split(",")) for p in polygons[0].split()[:2]]
    bottom = float(polygons[1].split()[0].split(",")[1])
    first = svg.split("</linearGradient>")[0]
    stops = tuple(
        (float(p), c) for p, c in re.findall(r"offset='([-0-9.e]+)' stop-color='#(\w{6})'", first)
    )
    inner = re.search(r"<rect [^>]*fill='#(\w{6})'", svg)
    return {
        "kind": "path",
        "box": (width, height),
        "rect": (corners[0][0], corners[0][1], corners[1][0], bottom),
        "stops": stops,
        "inner": inner[1] if inner else None,
    }


def drawn(path):
    """What sheetpeek's HTML draws for each case of the workbook at ``path``, as ``EXPECTED``."""
    group = path.stem.removeprefix("gradients_")
    page = sheetpeek.render(path, "html")
    boxes = re.findall(r'<i class="(sp-f-[a-z]+)" style="([^"]*)"></i>', page)
    sheet = sheetpeek.read(path)
    found = {}
    ordered = sorted(suite.cases(group), key=lambda c: _rc(c.cell))
    # a gradient is a box of the fill layer; a solid fill (no stops: black) is the cell's own
    filled = [c for c in ordered if isinstance(sheet.cell(*_rc(c.cell)).style.fill_spec, Gradient)]
    assert len(filled) == len(boxes), group  # the layer writes its boxes row by row
    for case in ordered:
        found[case.cell] = None
    for case, (name, style) in zip(filled, boxes, strict=True):
        body = re.search(rf"\.sheetpeek \.{name} \{{ ([^}}]*) \}}", page)[1]
        width = int(re.search(r"width:(\d+)px", style)[1])
        height = int(re.search(r"height:(\d+)px", style)[1])
        found[case.cell] = parse(body, width, height)
    return {cell: found[cell] for cell in (c.cell for c in suite.cases(group))}


def _hex(rgb):
    return "".join(f"{round(c):02X}" for c in rgb)


def _rgb(color):
    return tuple(int(color[i : i + 2], 16) for i in (0, 2, 4))


def _mix(a, b, part):
    """``part`` of the way from color ``a`` to ``b``, in sRGB."""
    return _hex(x + (y - x) * part for x, y in zip(_rgb(a), _rgb(b), strict=True))


def _along(stops, t):
    """A CSS or SVG gradient's color at ``t`` of its line: the ends hold, equal stops step."""
    if t <= stops[0][0]:
        return stops[0][1]
    for (p0, c0), (p1, c1) in itertools.pairwise(stops):
        if p0 <= t < p1:
            return _mix(c0, c1, (t - p0) / (p1 - p0))
    return stops[-1][1]


def paint(found, fx, fy):
    """The color sheetpeek draws at ``(fx, fy)`` of the cell, its box less the far gridlines."""
    w, h = found["box"]
    x, y = fx * (w - 1), fy * (h - 1)
    if found["kind"] == "linear":
        a = math.radians(found["angle"])
        length = abs(w * math.sin(a)) + abs(h * math.cos(a))
        t = ((x - w / 2) * math.sin(a) - (y - h / 2) * math.cos(a)) / length + 0.5
        return _along(found["stops"], t)
    left, top, right, bottom = found["rect"]
    # The band a point is in is the one it is farthest out in, as a share of the band's depth.
    depths = [(left - x) / left if left > 0 else -math.inf]
    depths.append((x - right) / (w - right) if right < w else -math.inf)
    depths.append((top - y) / top if top > 0 else -math.inf)
    depths.append((y - bottom) / (h - bottom) if bottom < h else -math.inf)
    t = max(depths)
    if t <= 0 and found["inner"] is not None:
        return found["inner"]
    return _along(found["stops"], max(t, 0.0))


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    folder = tmp_path_factory.mktemp("gradients")
    suite.build(folder)
    return folder


# --- the workbooks


def test_the_gradient_workbooks_read_back_as_built(built):
    for group in suite.GROUPS:
        path = built / f"gradients_{group}.xlsx"
        sheet = sheetpeek.read(path)
        for case in suite.cases(group):
            spec = sheet.cell(*_rc(case.cell)).style.fill_spec
            if not case.stops:
                # no stops: solid black, as Excel draws it (captured)
                assert spec == sheetpeek.Pattern("solid", "000000"), case.name
                continue
            assert isinstance(spec, Gradient), case.name
            assert spec.kind == case.kind, case.name
            assert sorted(spec.stops) == sorted(case.stops), case.name
            degree = float(case.degree or 0)
            expected = degree % 360 if math.isfinite(degree) else 0.0
            assert spec.degree == expected, case.name
            if case.rect is not None:
                assert (spec.left, spec.right, spec.top, spec.bottom) == case.rect, case.name
        assert drawn(path) == suite.EXPECTED[group], group


def test_each_gradient_is_written_as_the_case_gives_it(built):
    with zipfile.ZipFile(built / "gradients_odd.xlsx") as package:
        styles = package.read("xl/styles.xml").decode()
    two = (
        '<gradientFill><stop position="0"><color rgb="FFFF0000"/></stop>'
        '<stop position="0"><color rgb="FF0000FF"/></stop></gradientFill>'
    )
    assert two in styles  # in file order: red, then blue
    assert '<gradientFill degree="-45">' in styles
    assert str(suite.PLACEHOLDER) not in styles
    with zipfile.ZipFile(built / "gradients_not_finite.xlsx") as package:
        styles = package.read("xl/styles.xml").decode()
    for text in ("NaN", "INF", "-INF"):
        assert f'<gradientFill degree="{text}">' in styles
    with zipfile.ZipFile(built / "gradients_no_stops.xlsx") as package:
        styles = package.read("xl/styles.xml").decode()
    assert "<gradientFill></gradientFill>" in styles
    assert '<gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5">' in styles
    with zipfile.ZipFile(built / "gradients_main.xlsx") as package:
        styles = package.read("xl/styles.xml").decode()
    assert '<gradientFill type="path" left="0.2" right="0.8" top="0.3" bottom="0.7">' in styles
    assert '<stop position="0.25"><color rgb="FFFFFFFF"/></stop>' in styles


def test_every_case_has_a_cell_of_its_own_with_unfilled_neighbors():
    taken = {c.cell for c in suite.CASES}
    for group in suite.GROUPS:
        found = suite.cases(group)
        assert len({c.cell for c in found}) == len(found), group
        for case in found:
            row, column = _rc(case.cell)
            # the far gridlines are the case's own: nothing filled to the right or below
            assert f"{get_column_letter(column + 1)}{row}" not in taken
            assert f"{get_column_letter(column)}{row + 1}" not in taken
            assert case.cell not in suite.MARKERS
    last_row, last_column = _rc(suite.MARKERS[1])
    assert all(_rc(c.cell)[0] < last_row and _rc(c.cell)[1] < last_column for c in suite.CASES)


def test_no_case_draws_near_the_marker_color():
    marker = _rgb(MARKER)
    for case in suite.CASES:
        colors = [c for _, c in case.stops]
        # every color on the ramps between the stops, and from the last back to the first
        for a, b in zip(colors, colors[1:] + colors[:1], strict=True):
            for k in range(11):
                rgb = _rgb(_mix(a, b, k / 10))
                assert max(abs(x - y) for x, y in zip(rgb, marker, strict=True)) > 96, case.name


def test_paint_follows_sheetpeeks_gradient_along_the_axis(built):
    # The evaluator of sheetpeek's CSS against sheetpeek's model (Gradient.color_at) along the
    # axis of the 0 and 90 degree cases. sheetpeek's box is the cell and its far gridline, a
    # pixel more, so on the 57 px box of a 42 pt row the colors move by up to 255 / 57 levels
    # where one stretch spans the box, and by that over the stretch's share where it is shorter
    # (black at 0 to a first stop at 0.25: 4 times as many).
    sheet = sheetpeek.read(built / "gradients_main.xlsx")
    for case in suite.cases("main"):
        if case.kind != "linear" or case.degree not in (None, "90"):
            continue
        spec = sheet.cell(*_rc(case.cell)).style.fill_spec
        found = suite.EXPECTED["main"][case.cell]
        stops = itertools.pairwise(spec.drawn_stops)
        shortest = min(p1 - p0 for (p0, _), (p1, _) in stops if p1 > p0)
        for f in suite.fractions(suite.PROFILE):
            at = (f, 0.5) if case.degree is None else (0.5, f)
            drawn_rgb, model = _rgb(paint(found, *at)), _rgb(spec.color_at(f))
            off = max(abs(a - b) for a, b in zip(drawn_rgb, model, strict=True))
            assert off <= 5 / shortest, case.name


def test_paint_draws_a_path_from_its_rectangle_out():
    corner = suite.EXPECTED["main"]["Q4"]  # all four 0: the first stop at the top-left corner
    assert paint(corner, 0, 0) == "FFFFFF"
    assert _near(paint(corner, 1, 1), "0000FF", 2)  # the far gridline's pixel is past it
    assert paint(corner, 0.5, 0.25) == paint(corner, 0.25, 0.5) == paint(corner, 0.5, 0.5)
    center = suite.EXPECTED["main"]["S4"]
    # within the levels the box's extra pixel moves (the rectangle is at half the box)
    assert _near(paint(center, 0.5, 0.5), "FFFFFF", 4)
    for at in ((0, 0.5), (1, 0.5), (0.5, 0), (0.5, 1)):
        assert _near(paint(center, *at), "0000FF", 6), at
    box = suite.EXPECTED["main"]["U4"]
    assert box["inner"] == "FFFFFF"
    assert paint(box, 0.5, 0.5) == "FFFFFF"


# --- measuring, on a synthetic capture


def _canvas(group="main"):
    ws = openpyxl.Workbook().active
    suite.fill(ws, group)
    return Canvas(ws, [c.cell for c in suite.cases(group)])


def _case(cell, group="main"):
    return next(c for c in suite.cases(group) if c.cell == cell)


def _ramp(canvas, cell, axis):
    """A white-to-blue ramp over the cell and 1 pt (2 px) past its far edge on ``axis``."""
    x0, y0, x1, y1 = canvas.box(cell)
    span = (x1 - x0) if axis == "x" else (y1 - y0)
    for y in range(y0, y1 + (2 if axis == "y" else 0)):
        for x in range(x0, x1 + (2 if axis == "x" else 0)):
            t = min(1.0, ((x if axis == "x" else y) - (x0 if axis == "x" else y0) + 0.5) / span)
            canvas.paint(x, y, x + 1, y + 1, _rgb(_mix("FFFFFF", "0000FF", t)))


def _near(a, b, tolerance):
    return max(abs(x - y) for x, y in zip(_rgb(a), _rgb(b), strict=True)) <= tolerance


def test_a_horizontal_ramp_reads_as_its_colors_and_reaches_over_the_gridline():
    canvas = _canvas()
    _ramp(canvas, "Q6", "x")
    measured = suite.measure(canvas.capture(), _case("Q6"))
    assert (measured["width"], measured["height"]) == ("126", "42")
    rows = [row.split() for row in measured["grid"].split("/")]
    assert len(rows) == suite.GRID
    assert all(row == rows[0] for row in rows)
    for f, color in zip(suite.fractions(suite.GRID), rows[0], strict=True):
        assert _near(color, _mix("FFFFFF", "0000FF", f), 2), (f, color)
    profile = measured["profile_x"].split()
    assert len(profile) == suite.PROFILE
    assert [_rgb(c)[0] for c in profile] == sorted((_rgb(c)[0] for c in profile), reverse=True)
    assert len(set(measured["profile_y"].split())) == 1  # across the ramp: one color
    assert measured["fill_right"] == "1"  # 2 device pixels past the edge, one point
    assert measured["fill_bottom"] == "0"  # the ramp stops at the bottom edge


def test_a_vertical_ramp_reads_down_the_cell():
    canvas = _canvas()
    _ramp(canvas, "C4", "y")
    measured = suite.measure(canvas.capture(), _case("C4"))
    profile = measured["profile_y"].split()
    for f, color in zip(suite.fractions(suite.PROFILE), profile, strict=True):
        assert _near(color, _mix("FFFFFF", "0000FF", f), 2), (f, color)
    assert len(set(measured["profile_x"].split())) == 1
    assert measured["fill_bottom"] == "1"
    assert measured["fill_right"] == "0"


def test_a_fill_that_stops_at_its_edge_reaches_0():
    canvas = _canvas()
    x0, y0, x1, y1 = canvas.box("S6")
    canvas.paint(x0, y0, x1, y1, (0, 0, 255))
    measured = suite.measure(canvas.capture(), _case("S6"))
    assert (measured["fill_right"], measured["fill_bottom"]) == ("0", "0")
    assert set(measured["grid"].replace("/", " ").split()) == {"0000FF"}


def test_every_case_is_measured_on_a_blank_capture():
    for group in suite.GROUPS:
        cap = _canvas(group).capture()
        for case in suite.cases(group):
            measured = suite.measure(cap, case)
            assert measured is not None, case.name
            assert set(suite.COLUMNS) >= set(measured)
            assert set(measured["grid"].replace("/", " ").split()) == {"FFFFFF"}
            assert (measured["fill_right"], measured["fill_bottom"]) == ("", "")


def test_write_keeps_a_refusal_and_leaves_out_what_was_not_captured(tmp_path, monkeypatch):
    monkeypatch.setattr(suite, "BUILD_DIR", tmp_path)
    monkeypatch.setattr(suite, "OUTPUT", tmp_path / "out.tsv")
    suite.record_refusal("not_finite", "osascript: execution error: Excel refused (-50)")
    suite.write()
    lines = (tmp_path / "out.tsv").read_text(encoding="utf-8").splitlines()
    assert lines[0] == "# suite: gradients"
    assert lines[1].split("\t") == list(suite.COLUMNS)
    (row,) = [dict(zip(suite.COLUMNS, line.split("\t"), strict=True)) for line in lines[2:]]
    assert (row["group"], row["name"]) == ("not_finite", "refused: Excel refused (-50)")


# --- the comparison with Excel (suite gradients)

TOLERANCE = 16
"""Levels of 255 a sample may differ by in any channel, provisional until the capture's spread
is known. A sample is the mean of 3 by 3 device pixels; across a 44 pt cell at zoom 100 (88
device pixels) a white-to-blue ramp changes about 3 levels a pixel, so a sample placed a pixel
off (the cell's edges are rounded to device pixels) differs by 3 to 6, sheetpeek's box one pixel
wider (over the gridline) by about 3, and the conversion from the display's profile to sRGB and
any dithering by 2 or 3."""
MEDIAN = 6
"""The largest median difference of a case's samples: the systematic part of the error."""


@pytest.fixture(scope="module")
def captured():
    if not CAPTURE.exists():
        pytest.skip("not captured yet: run python -m capture.gradients run")
    lines = [ln for ln in CAPTURE.read_text(encoding="utf-8").splitlines() if ln[:2] != "# "]
    names = lines[0].split("\t")
    rows = [dict(zip(names, line.split("\t"), strict=False)) for line in lines[1:]]
    refused = {r["group"]: r["name"] for r in rows if r["name"].startswith("refused")}
    if refused:
        REPORT["gradients refused"] = "gradients: " + "; ".join(
            f"{group} {reason}" for group, reason in refused.items()
        )
    found = {(r["group"], r["cell"], int(r["zoom"])): r for r in rows if r["group"] not in refused}
    return found, refused


def _group(captured, group):
    found, refused = captured
    if group in refused:
        pytest.skip(f"Excel refused the file: {refused[group]}")
    rows = {key: row for key, row in found.items() if key[0] == group}
    if not rows:
        pytest.skip(f"group {group} not captured yet")
    return rows


def _excel_drawn(spec, row):
    """What sheetpeek draws for ``spec`` on a cell of Excel's measured size (CSS px at 4/3 pt)."""
    w = round(float(row["width"]) * 4 / 3) + 1
    h = round(float(row["height"]) * 4 / 3) + 1
    return parse(gradient_css(spec, w, h), w, h)


def _samples(row, which):
    """The measured samples as ``(fx, fy, color)``."""
    if which == "grid":
        places = suite.fractions(suite.GRID)
        return [
            (fx, fy, color)
            for fy, line in zip(places, row["grid"].split("/"), strict=True)
            for fx, color in zip(places, line.split(), strict=True)
        ]
    places = suite.fractions(suite.PROFILE)
    colors = row[which].split()
    if which == "profile_x":
        return [(f, 0.5, c) for f, c in zip(places, colors, strict=True)]
    return [(0.5, f, c) for f, c in zip(places, colors, strict=True)]


def _differences(found, row, which=("grid", "profile_x", "profile_y")):
    diffs = []
    for part in which:
        for fx, fy, color in _samples(row, part):
            expected = _rgb(paint(found, fx, fy))
            diffs.append(max(abs(a - b) for a, b in zip(_rgb(color), expected, strict=True)))
    return diffs


def _compare(rows, specs, label, which=("grid", "profile_x", "profile_y")):
    """Each case's samples against sheetpeek's drawing on Excel's geometry; reports the spread."""
    every, bad = [], []
    for (group, cell, zoom), row in sorted(rows.items()):
        spec = specs[group][cell]
        if spec is None:
            continue
        diffs = _differences(_excel_drawn(spec, row), row, which)
        every += diffs
        if max(diffs) > TOLERANCE or statistics.median(diffs) > MEDIAN:
            bad.append((cell, zoom, row["name"], statistics.median(diffs), max(diffs)))
    if every:
        ordered = sorted(every)
        REPORT[f"gradients {label}"] = (
            f"gradients {label}: {len(every)} samples, median {statistics.median(ordered):g}, "
            f"90th percentile {ordered[int(0.9 * (len(ordered) - 1))]:g}, max {ordered[-1]:g}"
        )
    assert not bad, bad


@pytest.fixture(scope="module")
def specs(built):
    found = {}
    for group in suite.GROUPS:
        sheet = sheetpeek.read(built / f"gradients_{group}.xlsx")
        found[group] = {
            c.cell: sheet.cell(*_rc(c.cell)).style.fill_spec for c in suite.cases(group)
        }
    return found


def _named(rows, test):
    return {key: row for key, row in rows.items() if test(row["name"])}


@pytest.mark.excel
def test_a_linear_gradient_takes_its_angle_in_the_unit_cell(captured, specs):
    # Seven angles in tall, square and wide cells, at zoom 100 and 200 (a gradient is laid out
    # in the cell, so the samples at each fraction are the same at both zooms); the corners
    # take the end stops, colors interpolate in sRGB.
    rows = _named(_group(captured, "main"), lambda n: n.startswith("linear"))
    _compare(rows, specs, "angles")


@pytest.mark.excel
def test_a_path_gradient_draws_bands_out_from_its_rectangle(captured, specs):
    rows = _named(_group(captured, "main"), lambda n: n.startswith("path"))
    _compare(rows, specs, "path")


@pytest.mark.excel
def test_before_a_first_stop_past_0_the_color_runs_from_the_last_stop(captured, specs):
    # Along the axis: 0 degrees across (profile_x), 90 degrees down (profile_y).
    rows = _named(_group(captured, "main"), lambda n: n.startswith("first stop"))
    across = _named(rows, lambda n: n.endswith(", 0"))
    down = _named(rows, lambda n: n.endswith(", 90"))
    _compare(across, specs, "first stop across", which=("profile_x",))
    _compare(down, specs, "first stop down", which=("profile_y",))


@pytest.mark.excel
def test_duplicate_stop_positions(captured, specs):
    rows = _named(_group(captured, "odd"), lambda n: "stop" in n or "0.5" in n)
    _compare(rows, specs, "duplicate stops", which=("profile_x",))


def _like_its_control(rows, controls):
    # Excel against itself: the case draws as the cell with the equivalent degree.
    for (group, cell, zoom), row in rows.items():
        control = controls.get(cell)
        if control is None:
            continue
        other = rows[(group, control, zoom)]
        for part in ("grid", "profile_x", "profile_y"):
            for a, b in zip(
                row[part].replace("/", " ").split(),
                other[part].replace("/", " ").split(),
                strict=True,
            ):
                assert _near(a, b, TOLERANCE // 2), (cell, control, part)


@pytest.mark.excel
def test_a_degree_out_of_range_is_taken_modulo_360(captured, specs):
    rows = _group(captured, "odd")
    _like_its_control(rows, suite.CONTROLS["odd"])
    _compare(_named(rows, lambda n: n.startswith("linear")), specs, "degrees")


@pytest.mark.excel
def test_a_degree_not_finite_is_0(captured, specs):
    rows = _group(captured, "not_finite")
    _like_its_control(rows, suite.CONTROLS["not_finite"])
    _compare(rows, specs, "not finite")


@pytest.mark.excel
def test_a_gradient_without_stops_draws_black(captured):
    rows = _group(captured, "no_stops")
    assert rows
    for (_, cell, _), row in rows.items():
        colors = row["grid"].replace("/", " ").split()
        assert all(_near(c, "000000", 8) for c in colors), (cell, row["name"])


@pytest.mark.excel
def test_a_gradient_covers_its_far_gridlines(captured):
    # sheetpeek draws the fill over the gridline after its right and bottom edges, 1 screen
    # point (one CSS px): 1 point at zoom 100, 0.5 at 200. Measured only where the fill's color at
    # the edge is told from white and the gridline; within a device pixel.
    found, _ = captured
    checked, bad = 0, []
    for (_, cell, zoom), row in found.items():
        for edge in ("fill_right", "fill_bottom"):
            if row[edge] == "":
                continue
            checked += 1
            if abs(float(row[edge]) - 100 / zoom) > 100 / zoom / 2 + 0.01:
                bad.append((cell, zoom, edge, row[edge]))
    REPORT["gradients gridlines"] = f"gradients: {checked} fill edges over the far gridline"
    assert not bad, bad
