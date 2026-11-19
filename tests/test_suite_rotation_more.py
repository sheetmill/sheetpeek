"""The screen suite ``rotation_more`` (ARCHITECTURE.md §4.17), without Excel.

Its workbook is built and read back with sheetpeek as built, and sheetpeek's layout of each case
is the suite's ``EXPECTED``. Its measuring runs on a synthetic capture with known marks. The
comparison with ``tests/captures/rotation_more.tsv`` is skipped until that capture exists.
"""

import math

import openpyxl
import pytest
from capture import rotation, rotation_more

import sheetpeek
from tests.screen_canvas import SCALE, Canvas
from tests.test_rotation import _TOPS

BLACK = (0, 0, 0)
BLUE = (0, 0, 192)


def drawn(path):
    """What sheetpeek draws for each case of the workbook at ``path``, as ``EXPECTED`` holds it.

    Lengths in points from the case's cell (or merge) left and top edges: each run's start of
    baseline, the baseline and its width; each underline's start and end against its run's and
    its top below the baseline; the slanted box's shift (``None``: drawn upright); for the
    ``neighbor`` case, the style drawn upright on each shared edge.
    """
    book = sheetpeek.open(path)
    found = {}
    for name, cell, case in rotation_more.placements():
        rendered = book.sheet(name).layout("excel")
        r = int(cell[1:])
        rows = {x.index: x for x in rendered.rows}
        row = rows[r]
        columns = {c.index: c for c in rendered.columns}
        column = columns[5]
        runs = [run for x in rendered.rows for run in x.runs if run.source == (r, 5)]
        i = list(rows).index(r)
        p = list(columns).index(5)
        borders = rendered.borders
        slanted = [s for s in (borders.slanted if borders else ()) if (s.row, s.column) == (i, p)]
        entry = {
            "height": row.height_pt,
            "angle": runs[0].angle if runs else None,
            "runs": tuple(
                (round(run.x - column.x_pt, 2), round(run.y - row.y_pt, 2), round(run.width, 2))
                for run in runs
            ),
            "rules": tuple(
                (
                    round(rule.x0 - run.x, 2),
                    round(rule.x1 - run.x - run.width, 2),
                    round(rule.y - run.y, 2),
                )
                for run in runs
                for rule in run.rules
            ),
            "slanted": round(slanted[0].shift_pt, 2) if slanted else None,
        }
        if case.neighbors:
            sides = (
                borders.vertical[i][p],
                borders.vertical[i][p + 1],
                borders.horizontal[i][p],
                borders.horizontal[i + 1][p],
            )
            entry["edges"] = tuple(side.style if side else None for side in sides)
        found[case.name] = entry
    return found


def _case(name):
    """The case's window, cell and case."""
    return next(p for p in rotation_more.placements() if p[2].name == name)


# --- the workbook


def test_the_workbook_reads_back_as_built(tmp_path):
    path = tmp_path / "rotation_more.xlsx"
    placed = rotation_more.build(path)
    assert placed == rotation_more.placements()
    assert [case for _, _, case in placed] == rotation_more.CASES
    assert len({case.name for case in rotation_more.CASES}) == len(rotation_more.CASES) == 43
    book = sheetpeek.open(path)
    letters = {"l": "left", "r": "right", "t": "top", "b": "bottom"}
    for name, cell, case in placed:
        sheet = book.sheet(name)
        r = int(cell[1:])
        style = sheet.cell(r, 5).style
        base = case.cell
        stored = base.rotation
        expected = 0 if stored == 255 else (stored if stored <= 90 else 90 - stored)
        assert (style.rotation, style.stacked) == (expected, stored == 255), case.name
        assert style.h_align == (base.horizontal or "general"), case.name
        assert style.v_align == (base.vertical or "bottom"), case.name
        sides = style.borders.sides if style.borders else (None,) * 4
        found = {side for side, value in zip(letters.values(), sides, strict=True) if value}
        if case.merge != "across":  # openpyxl gives the anchor the merge's inner side too
            assert found == {letters[s] for s in base.borders}, case.name
        assert (style.font.underline == "singleAccounting") is case.underline, case.name
        if case.merge:
            merges = {str(m) for m in sheet.merges}
            assert ":".join(rotation_more.area(cell, case)) in merges, case.name
        if case.neighbors:
            assert sheet.cell(r, 4).style.borders.right.style == "thick"
            assert sheet.cell(r, 6).style.borders.left.style == "thick"
            assert sheet.cell(r - 1, 5).style.borders.bottom.style == "thick"
            assert sheet.cell(r + 1, 5).style.borders.top.style == "thick"
            assert sheet.cell(r, 6).value == "x"
    assert drawn(path) == rotation_more.EXPECTED


def test_each_window_holds_five_cases_or_fewer_in_cells_of_their_own():
    names = []
    for name, chosen in rotation_more.windows():
        names.append(name)
        assert 1 <= len(chosen) <= rotation.PER_WINDOW, name
        assert len({case.cell.width for case in chosen}) == 1, name
        used = []
        for k, case in enumerate(chosen):
            row = rotation.FIRST_ROW + 2 * k
            cell = f"E{row}"
            used += dict.fromkeys(rotation_more.area(cell, case))
            if case.neighbors:
                used += [f"D{row}", f"F{row}", f"E{row - 1}", f"E{row + 1}"]
        assert len(used) == len(set(used)), name
        last = rotation.FIRST_ROW + 2 * (len(chosen) - 1)
        assert all(int(c[1:]) < last + 2 and c[0] < "J" for c in used), name  # inside the markers
    assert len(names) == len(set(names)) == 9


def test_the_two_slant_triggers_left_out_of_rotation_are_here():
    # rotation's slant group has all four sides, left and right, top and bottom, the left side
    # alone and the bottom side alone; this suite adds the right alone and the top alone.
    captured = {c.name for c in rotation.cases() + rotation.more_cases() if c.group == "slant"}
    assert {"45 l", "45 b", "45 lrtb", "45 lr", "45 tb"} <= captured
    assert {"45 r", "45 t"}.isdisjoint(captured)
    assert {c.cell.borders for c in rotation_more.CASES if c.group == "trigger"} == {"r", "t"}


def test_no_case_draws_magenta_or_a_fill():
    # text is blue (0000C0) in bordered cells and black elsewhere; sides are black
    assert rotation.BLUE == "FF0000C0"
    for case in rotation_more.CASES:
        assert case.cell.fill is False, case.name


# --- measuring, on a synthetic capture


def _canvas(name, heights=None):
    """A blank canvas of ``name``'s window, the case's cell and case, and the window's cases."""
    window, cell, case = _case(name)
    chosen = dict(rotation_more.windows())[window]
    ws = openpyxl.Workbook().active
    rotation_more.fill(ws, chosen)
    for row, height in (heights or {}).items():
        ws.row_dimensions[row].height = height
    return Canvas(ws, rotation_more.read_cells(chosen)), cell, case, chosen


def _measure(canvas, cell, case, chosen):
    return rotation_more.measure(rotation_more.complete(canvas.capture(), chosen), cell, case)


def test_a_top_side_alone_reads_as_its_extent_and_the_ink_box_as_blue():
    canvas, cell, case, chosen = _canvas("45 t")
    left, top, right, _ = canvas.box(cell)
    canvas.paint(left + 20, top - 1, left + 120, top + 1, BLACK)  # the top side, shifted
    canvas.paint(left + 10, top + 5, left + 30, top + 25, BLUE)  # the text
    found = _measure(canvas, cell, case, chosen)
    assert (found["top_from"], found["top_to"]) == ("10", "60")
    assert "bottom_from" not in found
    assert "border_left_25" not in found  # no side crosses the row's quarters
    assert (found["ink_left"], found["ink_top"], found["ink_bottom"]) == ("5", "2.5", "12.5")
    assert found["ink_right"] == f"{(left + 30 - right) / SCALE:g}"


def test_a_neighbors_sides_read_as_their_width_across_each_edge():
    canvas, cell, case, chosen = _canvas("45 lrtb thick neighbors")
    left, top, right, bottom = canvas.box(cell)
    cx = (left + right) // 2
    canvas.paint(left - 3, top, left + 3, bottom, BLACK)  # thick, 3 pt
    canvas.paint(right - 1, top, right + 1, bottom, BLACK)  # thin, 1 pt
    canvas.paint(cx - 10, top - 3, cx + 10, top + 3, BLACK)
    found = _measure(canvas, cell, case, chosen)
    edges = tuple(found[f"edge_{side}"] for side in ("left", "right", "top", "bottom"))
    assert edges == ("3", "1", "3", "0")


def test_a_merges_sides_are_measured_from_the_merges_edges():
    canvas, cell, case, chosen = _canvas("45 lrtb merged across")
    left, top, _, bottom = canvas.box(cell)
    right = canvas.box(f"F{rotation.FIRST_ROW}")[2]  # every row's F takes F4's column
    canvas.paint(right - 2, top, right, bottom, BLACK)
    found = _measure(canvas, cell, case, chosen)
    assert found["border_right_25"] == found["border_right_75"] == f"{(right - left) / SCALE:g}"
    canvas, cell, case, chosen = _canvas("45 lrtb merged down")
    spanned = rotation_more.merged(rotation_more.complete(canvas.capture(), chosen), cell, case)
    assert spanned.edge(cell, "bottom") - spanned.edge(cell, "top") == (40 + 30) * SCALE


def test_stacked_lines_read_one_band_each_top_first():
    row = int(_case("stacked right")[1][1:])
    canvas, cell, case, chosen = _canvas("stacked right", {row: 47})
    left, top, _, _ = canvas.box(cell)
    right = canvas.geometry[cell]["left"] * SCALE + canvas.geometry[cell]["width"] * SCALE
    spans = ((200, 212), (202, 212), (204, 210))
    for k, (x0, x1) in enumerate(spans):
        y = top + 6 + 30 * k
        canvas.paint(left + x0, y, left + x1, y + 14, BLACK)
    found = _measure(canvas, cell, case, chosen)
    expected = [f"{x0 / SCALE:g}:{(left + x1 - right) / SCALE:g}" for x0, x1 in spans]
    assert found["lines"] == ";".join(expected)


def test_an_upright_turned_underline_reads_against_the_texts_ink():
    row = int(_case("90 accounting text")[1][1:])
    canvas, cell, case, chosen = _canvas("90 accounting text", {row: 85})
    left, top, _, _ = canvas.box(cell)
    canvas.paint(left + 20, top + 10, left + 34, top + 80, BLACK)  # the text, reading upward
    found = _measure(canvas, cell, case, chosen)
    assert "rule_start" not in found  # letters alone are no rule
    canvas.paint(left + 38, top + 14, left + 40, top + 76, BLACK)  # 2 px, past a 4 px gap
    found = _measure(canvas, cell, case, chosen)
    assert (found["rule_start"], found["rule_end"], found["rule_offset"]) == ("2", "-2", "2.75")


def test_a_turned_underline_at_45_degrees_reads_along_the_line():
    row = int(_case("45 accounting text")[1][1:])
    canvas, cell, case, chosen = _canvas("45 accounting text", {row: 85})
    left, top, _, bottom = canvas.box(cell)
    ox, oy = left + 10, bottom - 10  # the baseline's start
    sin = cos = math.sqrt(0.5)
    for y in range(top, bottom):
        for x in range(left, left + 240):
            s = (x - ox) * cos - (y - oy) * sin
            t = (x - ox) * sin + (y - oy) * cos
            if (0 <= s <= 100 and -14 <= t <= 0) or (10 <= s <= 90 and 4 <= t <= 6):
                canvas.paint(x, y, x + 1, y + 1, BLACK)
    found = _measure(canvas, cell, case, chosen)
    assert float(found["rule_start"]) == pytest.approx(5, abs=0.75)
    assert float(found["rule_end"]) == pytest.approx(-5, abs=0.75)
    assert float(found["rule_offset"]) == pytest.approx(2.5, abs=0.75)


def test_every_case_is_measured_on_a_blank_capture():
    for _, chosen in rotation_more.windows():
        ws = openpyxl.Workbook().active
        rotation_more.fill(ws, chosen)
        canvas = Canvas(ws, rotation_more.read_cells(chosen))
        cap = rotation_more.complete(canvas.capture(), chosen)
        for k, case in enumerate(chosen):
            cell = f"E{rotation.FIRST_ROW + 2 * k}"
            found = rotation_more.measure(cap, cell, case)
            assert not {"ink_left", "lines", "rule_start", "top_from"} & set(found), case.name
            if "edges" in case.measure:
                assert {found[f"edge_{s}"] for s in ("left", "right", "top", "bottom")} == {"0"}
            record = rotation_more.record(cap, cell, case)
            assert set(rotation_more.COLUMNS) >= set(record)
            assert (record["group"], record["name"]) == (case.group, case.name)


# --- the capture (suite rotation_more)


def _captured():
    """The captured rows by case name."""
    path = rotation_more.OUTPUT
    if not path.exists():
        pytest.skip("not captured yet: run python -m capture.rotation_more run")
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    return {record["name"]: record for record in (dict(zip(header, r, strict=False)) for r in body)}


def _row(name):
    found = _captured()
    if name not in found:
        pytest.skip(f"{name}: its window was not captured yet")
    return found[name]


@pytest.fixture(scope="module")
def widths(tmp_path_factory):
    """Each case's cell (or merge) width in points, as sheetpeek lays it out."""
    path = tmp_path_factory.mktemp("rotation_more") / "rotation_more.xlsx"
    rotation_more.build(path)
    book = sheetpeek.open(path)
    found = {}
    for name, _, case in rotation_more.placements():
        columns = {c.index: c for c in book.sheet(name).layout("excel").columns}
        last = 6 if case.merge == "across" else 5
        found[case.name] = sum(columns[c].width_pt for c in range(5, last + 1))
    return found


def _ink(name, width):
    """The modeled ink of a one-run case: left, right past the far edge, top and bottom.

    The glyphs' box inside the line, as ``tests/test_rotation.py`` models it: 0.8 pt in from
    each end, from the baseline to the end letter's own top (the texts have no descenders).
    """
    entry = rotation_more.EXPECTED[name]
    case = _case(name)[2]
    ((x, y, w),) = entry["runs"]
    angle = math.radians(entry["angle"])
    cos, sin = math.cos(angle), math.sin(angle)
    size, text = case.cell.size, case.cell.text
    points = [
        (x + s * cos - t * sin, y - s * sin - t * cos)
        for s, letter in ((0.8, text[0]), (w - 0.8, text[-1]))
        for t in (0.0, _TOPS.get(letter, 0.68) * size)
    ]
    xs, ys = [p[0] for p in points], [p[1] for p in points]
    return min(xs), max(xs) - width, min(ys), max(ys)


def _names(*groups, slanted=None):
    return [
        c.name
        for c in rotation_more.CASES
        if c.group in groups
        and (slanted is None or (rotation_more.EXPECTED[c.name]["slanted"] is not None) == slanted)
    ]


INK_BOUND = 1.5
"""Points within which an angled case's ink edges lie: across the line between slanted sides
(``rotation``'s slant cases: median 0.6, 90th percentile 1.1, the largest outside its three
named outliers 1.5), in x and y elsewhere (``rotation``'s unbordered and turned cases: within
1.1 and 1.5)."""


@pytest.mark.excel
@pytest.mark.parametrize(
    "name", _names("trigger", "neighbor", "top", "negative", "vertical", slanted=True)
)
def test_text_between_slanted_sides_sits_where_sheetpeek_puts_it(name, widths):
    # §4.17: top alignment at 15 to 75 and -45 degrees (top_slant_45: placed in height as without
    # borders), left, center and right at negative angles (taken as the mirror of the positive
    # ones), vertical justify and distributed, and the right or top side alone (slanting as any
    # side does). Compared across the line, as rotation's slant cases are.
    record = _row(name)
    left, right, top, bottom = _ink(name, widths[name])
    across = abs(math.sin(math.radians(rotation_more.EXPECTED[name]["angle"])))
    assert abs(float(record["ink_left"]) - left) * across <= INK_BOUND, (name, left)
    assert abs(float(record["ink_right"]) - right) * across <= INK_BOUND, (name, right)
    if _case(name)[2].cell.vertical is not None:  # how high the text sits, as without borders
        assert abs(float(record["ink_top"]) - top) <= INK_BOUND, (name, top)
        assert abs(float(record["ink_bottom"]) - bottom) <= INK_BOUND, (name, bottom)


@pytest.mark.excel
@pytest.mark.parametrize("name", _names("merge", "vertical", "spread", slanted=False))
def test_angled_text_without_slanted_sides_sits_where_sheetpeek_puts_it(name, widths):
    # §4.17: a merge's text placed as without borders, vertical justify and distributed without
    # borders, and fill, justify and distributed at 15, 75, -45 and 90 degrees placed as left.
    record = _row(name)
    left, right, top, bottom = _ink(name, widths[name])
    assert abs(float(record["ink_left"]) - left) <= INK_BOUND, (name, left)
    assert abs(float(record["ink_right"]) - right) <= INK_BOUND, (name, right)
    assert abs(float(record["ink_top"]) - top) <= INK_BOUND, (name, top)
    assert abs(float(record["ink_bottom"]) - bottom) <= INK_BOUND, (name, bottom)


SIDE_BOUND = 1.25
"""Points across the line within which a slanted side's center lies (``rotation``'s slant cases:
the farthest 1.05, at 15 degrees)."""


def _side_names():
    return [
        c.name
        for c in rotation_more.CASES
        if set(c.cell.borders) & {"l", "r"} and not c.neighbors  # D's and F's sides are dark too
    ]


@pytest.mark.excel
@pytest.mark.parametrize("name", _side_names())
def test_the_sides_lie_where_sheetpeek_draws_them(name, widths):
    # The dark run of each side at 25 % and 75 % of the row (of the merge): slanted by Excel's
    # own row height, h·cot θ at the top, or upright for a merge. With one side, the run is that
    # side, and its center is compared.
    record = _row(name)
    case = _case(name)[2]
    entry = rotation_more.EXPECTED[name]
    width = widths[name]
    angle = math.radians(entry["angle"])
    sine = abs(math.sin(angle))
    height = float(record["row_height"])
    if case.merge == "down":
        height += rotation_more.MERGE_HEIGHTS["down"][1]
    shift = height / math.tan(angle) if entry["slanted"] is not None else 0.0
    for share in (25, 75):
        if not record[f"border_left_{share}"]:
            pytest.fail(f"{name}: no side found at {share} %")
        start, end = float(record[f"border_left_{share}"]), float(record[f"border_right_{share}"])
        base = shift * (1 - share / 100)
        if case.cell.borders == "r":
            centers = [((start + end) / 2, base + width)]
        else:
            half = 1 / (2 * sine) if entry["slanted"] is not None else 0.5
            centers = [(start + half, base), (end - half, base + width)]
        for measured, model in centers:
            assert abs(measured - model) * (sine if shift else 1.0) <= SIDE_BOUND, (
                name,
                share,
                measured,
                model,
            )


@pytest.mark.excel
@pytest.mark.parametrize(
    "name", [c.name for c in rotation_more.CASES if "t" in c.cell.borders and not c.neighbors]
)
def test_the_top_side_lies_where_sheetpeek_draws_it(name, widths):
    # Slanted, the top side runs along the cell's top edge from h·cot θ to h·cot θ plus the
    # width (the "45 t" case: whether the top side alone slants); upright, over the cell.
    # A line's end and its antialiasing: within 1.5 pt.
    record = _row(name)
    entry = rotation_more.EXPECTED[name]
    if not record["top_from"]:
        pytest.fail(f"{name}: no top side found")
    height = float(record["row_height"])
    if _case(name)[2].merge == "down":
        height += rotation_more.MERGE_HEIGHTS["down"][1]
    angle = math.radians(entry["angle"])
    shift = height / math.tan(angle) if entry["slanted"] is not None else 0.0
    lo, hi = sorted((shift, shift + widths[name]))
    assert float(record["top_from"]) == pytest.approx(lo, abs=1.5), name
    assert float(record["top_to"]) == pytest.approx(hi, abs=1.5), name


NEIGHBOR_SIDES = {None: (0.0, 0.0), "thin": (0.5, 1.5), "thick": (2.5, 3.5)}
"""Points of dark pixels across an edge for each style (``borders``: thin 1, thick 3)."""


@pytest.mark.excel
def test_a_neighbors_side_on_a_shared_edge_is_drawn_upright():
    # §4.17, slant_neighbor_side: D's, F's and the cells above and below's own thick sides on the
    # edges shared with the slanted cell, drawn upright (the cell's own sides have left).
    name = "45 lrtb thick neighbors"
    record = _row(name)
    expected = rotation_more.EXPECTED[name]["edges"]
    for side, style in zip(("left", "right", "top", "bottom"), expected, strict=True):
        lo, hi = NEIGHBOR_SIDES[style]
        assert lo <= float(record[f"edge_{side}"]) <= hi, (side, style, record[f"edge_{side}"])


STACKED_BEARINGS = (0.15, 0.1)
"""Ems by which a glyph's ink may lie inside its advance on the left and the right (Calibri's
"A", "b" and "c": under 0.1), plus half a point for the device pixel the ink ends on."""


@pytest.mark.excel
@pytest.mark.parametrize("name", _names("stacked"))
def test_stacked_characters_sit_where_sheetpeek_puts_them(name, widths):
    # §4.17, stacked_align: each line's ink inside its character's advance, left, right, center
    # and general in a 20-character column, where the three placements lie about 50 pt apart.
    record = _row(name)
    lines = [tuple(map(float, line.split(":"))) for line in record["lines"].split(";")]
    runs = rotation_more.EXPECTED[name]["runs"]
    assert len(lines) == len(runs), name
    size = _case(name)[2].cell.size
    for (ink_left, ink_right), (x, _, w) in zip(lines, runs, strict=True):
        left, right = x, x + w - widths[name]
        assert -0.5 <= ink_left - left <= STACKED_BEARINGS[0] * size + 0.5, (name, ink_left, left)
        assert -0.5 <= right - ink_right <= STACKED_BEARINGS[1] * size + 0.5, (name, ink_right)


RULE_BOUND = 1.5
"""Points within which an underline's ends lie along the line, against the text's ink: the
model's ink is 0.8 pt in from the run's ends (``tests/test_rotation.py``), within 1.1."""


@pytest.mark.excel
@pytest.mark.parametrize("name", _names("underline"))
def test_an_angled_accounting_underline_runs_where_sheetpeek_draws_it(name):
    # §4.17, angled_acct_underline: the rule's start and end along the turned line against the
    # text's ink, and its distance below the ink across the line (the top of the 1 pt line 3 pt
    # below the baseline: its center 3.5 below, within a point).
    record = _row(name)
    ((start, end, below),) = rotation_more.EXPECTED[name]["rules"]
    assert record["rule_start"], f"{name}: no underline found"
    assert float(record["rule_start"]) == pytest.approx(start - 0.8, abs=RULE_BOUND), name
    assert float(record["rule_end"]) == pytest.approx(end + 0.8, abs=RULE_BOUND), name
    assert float(record["rule_offset"]) == pytest.approx(below + 0.5, abs=1.0), name


@pytest.mark.excel
@pytest.mark.parametrize("name", [c.name for c in rotation_more.CASES])
def test_each_row_is_as_tall_as_sheetpeek_makes_it(name):
    # Set rows exactly; fitted ones (§4.17, fitted rows) within rotation's FITTED_BOUND, 2 pt.
    record = _row(name)
    expected = rotation_more.EXPECTED[name]["height"]
    assert abs(float(record["row_height"]) - expected) <= 2.0, (name, record["row_height"])
