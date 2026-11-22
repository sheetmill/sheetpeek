"""The screen suite ``outline``: its workbooks, its measuring, and sheetpeek against its capture.

ARCHITECTURE.md §4.11. The workbooks are built and read back; the measuring runs on synthetic
windows (a gutter, headings and cells at 2 pixels a screen point, symbols painted where
sheetpeek's model puts them). The comparisons with ``tests/captures/outline.tsv`` are skipped
until the suite has been run against Excel.

Tolerances: Excel draws the gutter at whole device pixels, half a screen point on a 2x display,
and rounds a button's corners, so a length or offset may be off by one device pixel: 0.5
screen points. A flat fill keeps its color through the sRGB conversion within a few levels:
3 per channel.
"""

import math
import re
from pathlib import Path

import openpyxl
import pytest
from capture import outline
from capture.screen import png
from capture.screen.run import Capture
from capture.screen.window import Grid
from openpyxl.utils import column_index_from_string, get_column_letter

import sheetpeek
from tests.screen_canvas import SCALE, Canvas

CAPTURED = Path(__file__).parent / "captures" / "outline.tsv"
TOLERANCE = 0.5
COLOR_TOLERANCE = 3
LANE, MARGIN, END = 12, 2, 3
BUTTON, BAR, INSET, START, TICK, DOT = 12, 2, 3, 2, 6, 2
"""§4.11's geometry in screen points."""


def centered(length, size):
    """Where a box ``size`` long starts in a row ``length`` long, below its 1-point gridline."""
    return 1 + math.ceil((length - 1 - size) / 2)


def expected_pieces(marks, length, after):
    """The pieces the measuring should find for ``marks`` (lane, kind) in a row ``length`` long:
    (kind, across, along, across size, along size) in screen points, as §4.11 states them."""
    found = []
    button = centered(length, BUTTON)
    for lane, kind in marks:
        x0 = MARGIN + LANE * (lane - 1)
        x = x0 + INSET
        if kind == "bar":
            found.append(("bar", x, 0, BAR, length))
        elif kind == "start" and after:
            found += [("tick", x, START, TICK, BAR), ("bar", x, START + BAR, BAR, length - 4)]
        elif kind == "start":
            end = length - START  # the tick ends at the row's edge (captured, summaries)
            found += [("bar", x, 0, BAR, end), ("tick", x, end, TICK, BAR)]
        elif kind == "open":
            if after and button:
                found.append(("bar", x, 0, BAR, button))
            found.append(("minus", x0, button, BUTTON, BUTTON))
            if not after and length - button - BUTTON:
                found.append(("bar", x, button + BUTTON, BAR, length - button - BUTTON))
        elif kind == "closed":
            found.append(("plus", x0, button, BUTTON, BUTTON))
        else:
            found.append(("dot", x, centered(length, DOT), DOT, DOT))
    return found


def gutter_size(lanes):
    return MARGIN + LANE * lanes + END


def parse_marks(text):
    return [(int(lane), kind) for lane, _, kind in (m.partition(":") for m in text.split())]


def layout(tmp_path, name):
    """sheetpeek's outline for a group: (row lanes, column lanes), rows and columns as text."""
    wb = openpyxl.Workbook()
    outline.fill(wb.active, name)
    path = tmp_path / f"outline_{name}.xlsx"
    wb.save(path)
    post = outline.postprocess(name)
    if post is not None:
        post(path)
    found = sheetpeek.read(path).layout("excel").outline
    if found is None:
        return None
    rows, columns = {}, {}
    for mark in found.rows:
        rows.setdefault(mark.index, []).append(f"{mark.lane}:{mark.kind}")
    for mark in found.columns:
        columns.setdefault(get_column_letter(mark.index), []).append(f"{mark.lane}:{mark.kind}")
    return {
        "lanes": (found.row_lanes, found.column_lanes),
        "rows": {k: " ".join(v) for k, v in rows.items()},
        "columns": {k: " ".join(v) for k, v in columns.items()},
    }


# --- the workbooks


def test_every_group_builds_and_reads_back_as_built(tmp_path):
    names = outline.build(tmp_path)
    assert names == list(outline.GROUPS)
    for name, group in outline.GROUPS.items():
        sheet = sheetpeek.read(tmp_path / f"outline_{name}.xlsx")
        assert sheet.outline is not None, name
        levels = {r.index: (r.outline_level, r.hidden, r.collapsed) for r in sheet.rows}
        for row, (level, hidden, collapsed, height) in group.rows.items():
            assert levels[row] == (level, hidden, collapsed), (name, row)
            if height is not None:
                assert sheet.rows[row - 1].height == height
        columns = {c.letter: (c.outline_level, c.hidden, c.collapsed) for c in sheet.columns}
        for letter, (level, hidden, collapsed, _) in group.columns.items():
            assert columns[letter] == (level, hidden, collapsed), (name, letter)
        assert (sheet.outline.summary_below, sheet.outline.summary_right) == (
            group.below,
            group.right,
        )
        # the view alone decides: Excel ignores <outlinePr showOutlineSymbols> (captured)
        shown = {None: True, "0": False, "1": True, "yes": True}[group.view]
        assert sheet.outline.show_symbols is shown, name


def test_the_files_state_what_each_group_is_about(tmp_path):
    import zipfile

    outline.build(tmp_path)

    def part(name):
        with zipfile.ZipFile(tmp_path / f"outline_{name}.xlsx") as archive:
            return archive.read("xl/worksheets/sheet1.xml").decode()

    assert "outlineLevelRow" not in part("main")
    assert 'outlineLevelRow="3"' in part("declared")
    assert 'outlineLevelCol="2"' in part("declared")
    assert 'outlineLevelRow="1"' in part("declared_small")
    assert '<sheetView showOutlineSymbols="0"' in part("view_off")
    assert re.search(r'<outlinePr [^>]*showOutlineSymbols="0"', part("props_off"))
    assert '<sheetView showOutlineSymbols="yes"' in part("flag_yes")
    for name in ("view_on_props_off", "view_off_props_on"):
        view, props = outline.GROUPS[name].view, outline.GROUPS[name].props
        assert f'<sheetView showOutlineSymbols="{view}"' in part(name)
        assert f'showOutlineSymbols="{props}"' in part(name).split("<sheetViews>")[0]
    assert outline.REFUSALS == ("flag_yes",)
    assert all(outline.GROUPS[name].zooms == (100, 200) for name in outline.LIGHT)


def test_expected_is_what_sheetpeek_lays_out(tmp_path):
    for name in outline.GROUPS:
        expected = outline.EXPECTED[name]
        if expected == "the same as main":
            expected = outline.EXPECTED["main"]
        assert layout(tmp_path, name) == expected, name


def test_no_marker_shares_a_measured_row_or_column(tmp_path):
    for group in outline.GROUPS.values():
        rows, columns = outline.measured(group)
        letters, row = re.fullmatch(r"([A-Z]+)(\d+)", group.marker).groups()
        assert int(row) > max(rows)
        assert column_index_from_string(letters) > column_index_from_string(columns[-1])


# --- the measuring, on synthetic windows

GUTTER, HEADINGS, BAR_INK, BORDER, PAPER = (85,) * 3, (22,) * 3, (153,) * 3, (51,) * 3, (255,) * 3
ROW_HEADINGS, COLUMN_HEADINGS = 22, 21
"""The headings' widths in screen points, as a dark window at zoom 100 shows them."""


class Window:
    """A synthetic window: gutters, headings and the cells of a group at zoom 100."""

    def __init__(self, name, lanes, headings=HEADINGS, gutter=GUTTER):
        ws = openpyxl.Workbook().active
        self.group = outline.fill(ws, name)
        cells = Canvas(ws, outline.cells(self.group))
        self.geometry = cells.geometry
        rows, columns = lanes
        self.gx = gutter_size(rows) * SCALE if rows else 0
        self.gy = gutter_size(columns) * SCALE if columns else 0
        self.ox, self.oy = self.gx + ROW_HEADINGS * SCALE, self.gy + COLUMN_HEADINGS * SCALE
        self.width, self.height = self.ox + cells.width, self.oy + cells.height
        self.rows = [bytearray(b"\xff" * (self.width * 3)) for _ in range(self.height)]
        self.paint(0, 0, self.width, self.oy, gutter)
        self.paint(self.gx, self.oy, self.ox, self.height, headings)
        self.paint(self.ox, self.gy, self.width, self.oy, headings)
        self.paint(self.gx, self.gy, self.ox, self.oy, headings)
        self.paint(0, self.oy, self.gx, self.height, gutter)

    def paint(self, x0, y0, x1, y1, rgb):
        for y in range(y0, y1):
            for x in range(x0, x1):
                self.rows[y][3 * x : 3 * x + 3] = bytes(rgb)

    def piece(self, axis, at, kind, across, along, wide, long):
        """Paint one expected piece (screen points) in the row or column whose stretch starts
        at ``at`` pixels along."""
        a0, b0 = round(across * SCALE), at + round(along * SCALE)
        a1, b1 = a0 + round(wide * SCALE), b0 + round(long * SCALE)

        def rect(p0, q0, p1, q1, rgb):
            if axis == "rows":
                self.paint(p0, q0, p1, q1, rgb)
            else:
                self.paint(q0, p0, q1, p1, rgb)

        if kind in ("bar", "tick", "dot"):
            rect(a0, b0, a1, b1, BAR_INK)
            return
        rect(a0, b0, a1, b1, BORDER)
        rect(a0 + 2, b0 + 2, a1 - 2, b1 - 2, PAPER)
        middle_a, middle_b = (a0 + a1) // 2, (b0 + b1) // 2
        rect(middle_a - 6, middle_b - 1, middle_a + 6, middle_b + 1, BORDER)
        if kind == "plus":
            rect(middle_a - 1, middle_b - 6, middle_a + 1, middle_b + 6, BORDER)

    def capture(self):
        image = png.Image(self.width, self.height, 3, self.rows, "sRGB")
        grid = Grid(float(self.ox), float(self.oy), float(SCALE))
        return Capture(100, image, grid, {}, self.geometry, [], "mac")

    def start(self, axis, index):
        cell = f"A{index}" if axis == "rows" else f"{index}1"
        g = self.geometry[cell]
        if axis == "rows":
            return self.oy + round(g["top"] * SCALE), g["height"]
        return self.ox + round(g["left"] * SCALE), g["width"]


def found_pieces(records, axis, index):
    return sorted(
        (
            r["kind"],
            float(r["across"]),
            float(r["along"]),
            float(r["across_size"]),
            float(r["along_size"]),
        )
        for r in records
        if r["axis"] == axis and r["index"] == str(index) and r["kind"] != "none"
    )


def test_the_gutters_and_their_symbols_read_as_painted():
    expected = outline.EXPECTED["main"]
    window = Window("main", expected["lanes"])
    painted = {}
    for axis, after in (("rows", True), ("columns", True)):
        for index, text in expected[axis].items():
            at, length = window.start(axis, index)
            painted[(axis, index)] = expected_pieces(parse_marks(text), length, after)
            for piece in painted[(axis, index)]:
                window.piece(axis, at, *piece)
    records = outline.measure(window.capture(), "main")
    gutters = {r["index"]: r for r in records if r["axis"] == "gutter"}
    assert (gutters["rows"]["across_size"], gutters["rows"]["along_size"]) == ("53", "22")
    assert (gutters["columns"]["across_size"], gutters["columns"]["along_size"]) == ("41", "21")
    assert gutters["rows"]["ink"] == "555555"
    for (axis, index), pieces in painted.items():
        assert found_pieces(records, axis, index) == sorted(pieces), (axis, index)
    # the rows and columns without symbols read as none
    assert [r["kind"] for r in records if r["axis"] == "rows" and r["index"] == "2"] == ["none"]
    minus = next(r for r in records if r["index"] == "7" and r["kind"] == "minus")
    assert (minus["ink"], minus["paper"], minus["stroke"], minus["stroke_size"]) == (
        "333333",
        "FFFFFF",
        "333333",
        "6,1",
    )
    plus = next(r for r in records if r["index"] == "11" and r["kind"] == "plus")
    assert plus["stroke_size"] == "6,6"
    bar = next(r for r in records if r["index"] == "5" and r["kind"] == "bar")
    assert bar["ink"] == "999999"
    assert set(outline.COLUMNS) >= set(records[0]) | set(minus)


def test_summaries_above_read_as_painted():
    expected = outline.EXPECTED["summaries"]
    window = Window("summaries", expected["lanes"])
    painted = {}
    for axis in ("rows", "columns"):
        for index, text in expected[axis].items():
            at, length = window.start(axis, index)
            painted[(axis, index)] = expected_pieces(parse_marks(text), length, after=False)
            for piece in painted[(axis, index)]:
                window.piece(axis, at, *piece)
    records = outline.measure(window.capture(), "summaries")
    for (axis, index), pieces in painted.items():
        assert found_pieces(records, axis, index) == sorted(pieces), (axis, index)


def test_level_buttons_side_by_side_read_one_each():
    window = Window("main", (4, 3))
    band = window.oy - window.gy
    top = window.gy + centered(band // SCALE, BUTTON) * SCALE
    for lane in range(1, 5):
        window.piece("rows", top, "minus", MARGIN + LANE * (lane - 1), 0, BUTTON, BUTTON)
    records = outline.measure(window.capture(), "main")
    levels = [r for r in records if r["axis"] == "levels" and r["index"] == "rows"]
    assert [(r["kind"], r["across"], r["across_size"], r["along_size"]) for r in levels] == [
        ("level", str(MARGIN + LANE * k), "12", "12") for k in range(4)
    ]
    assert levels[0]["along"] == str(centered(band // SCALE, BUTTON))
    assert levels[0]["size"] == "21"


def test_a_window_without_a_gutter_reads_none():
    window = Window("view_off", (0, 0))
    records = outline.measure(window.capture(), "view_off")
    gutters = {r["index"]: r for r in records if r["axis"] == "gutter"}
    assert gutters["rows"]["across_size"] == gutters["columns"]["across_size"] == ""
    assert gutters["rows"]["along_size"] == "22"
    assert [r["axis"] for r in records] == ["gutter", "gutter"]


def test_a_light_gutter_reads_with_its_own_colors():
    window = Window("view_on_props_off", (4, 3), headings=(230, 230, 230), gutter=(242,) * 3)
    records = outline.measure(window.capture(), "view_on_props_off")
    gutters = {r["index"]: r for r in records if r["axis"] == "gutter"}
    assert (gutters["rows"]["across_size"], gutters["rows"]["ink"]) == ("53", "F2F2F2")
    assert gutters["columns"]["across_size"] == "41"


# --- the capture (suite outline)


def _captured():
    if not CAPTURED.exists():
        pytest.skip("not captured yet: run the outline suite against Excel")
    lines = CAPTURED.read_text(encoding="utf-8").splitlines()
    rows = [line.split("\t") for line in lines if line and not line.startswith("#")]
    header, body = rows[0], rows[1:]
    return [dict(zip(header, row, strict=False)) for row in body]


def _close(a, b, tolerance=TOLERANCE):
    return abs(a - b) <= tolerance


def _matches(found, expected):
    """Whether two piece lists pair up one to one, kinds equal and numbers within tolerance."""
    left = list(found)

    def near(f, piece):
        return f[0] == piece[0] and all(_close(a, b) for a, b in zip(f[1:], piece[1:], strict=True))

    for piece in expected:
        hit = next((f for f in left if near(f, piece)), None)
        if hit is None:
            return False
        left.remove(hit)
    return not left


def _groups(records):
    found = {}
    for r in records:
        found.setdefault(
            (r["appearance"], r["group"], r["zoom"], r["axis"], r["index"]), []
        ).append(r)
    return found


def _compare(tmp_path, records, names, lanes_only=None):
    """Each measured row and column of ``names`` against sheetpeek's marks; the mismatches."""
    misses = []
    for name in names:
        expected = layout(tmp_path, name)
        group = outline.GROUPS[name]
        for key, found in _groups(r for r in records if r["group"] == name).items():
            axis, index = key[3], key[4]
            if axis not in ("rows", "columns"):
                continue
            marks = parse_marks(expected[axis].get(int(index) if axis == "rows" else index, ""))
            after = group.below if axis == "rows" else group.right
            want = expected_pieces(marks, float(found[0]["size"]), after)
            got = found_pieces(found, axis, index)
            if lanes_only is not None:
                want = sorted(
                    (k, round((a - MARGIN) / LANE)) for k, a, *_ in want if k in lanes_only
                )
                got = sorted((k, round((a - MARGIN) / LANE)) for k, a, *_ in got if k in lanes_only)
                if want != got:
                    misses.append((key, got, want))
            elif not _matches(got, want):
                misses.append((key, got, want))
    return misses


@pytest.mark.excel
def test_buttons_are_minus_while_a_detail_shows_and_plus_when_collapsed_as_captured(tmp_path):
    records = _captured()
    misses = _compare(tmp_path, records, ("main", "heights"), lanes_only=("minus", "plus"))
    assert not misses, misses


@pytest.mark.excel
def test_each_visible_detail_gets_a_dot_in_the_lane_after_its_level_as_captured(tmp_path):
    records = _captured()
    misses = _compare(tmp_path, records, ("main", "heights"), lanes_only=("dot",))
    assert not misses, misses


@pytest.mark.excel
def test_summaries_above_and_left_mirror_the_rule_as_captured(tmp_path):
    misses = _compare(tmp_path, _captured(), ("summaries",))
    assert not misses, misses


@pytest.mark.excel
def test_bars_ticks_buttons_and_dots_sit_at_fixed_screen_sizes_as_captured(tmp_path):
    # Both zooms and rows of 15, 30 and 60 points: every piece where §4.11 puts it.
    misses = _compare(tmp_path, _captured(), ("main", "heights"))
    assert not misses, misses


def _gutters(records, name, appearance="dark"):
    return {
        (r["zoom"], r["index"]): r
        for r in records
        if r["group"] == name and r["axis"] == "gutter" and r["appearance"] == appearance
    }


@pytest.mark.excel
def test_the_gutter_is_as_deep_as_the_deepest_level_whatever_is_declared_as_captured():
    records = _captured()
    rows, columns = outline.EXPECTED["main"]["lanes"]
    for name in ("main", "declared", "declared_small", "heights", "summaries"):
        expected = outline.EXPECTED[name]
        lanes = (rows, columns) if expected == "the same as main" else expected["lanes"]
        for (zoom, index), r in _gutters(records, name).items():
            want = gutter_size(lanes[0] if index == "rows" else lanes[1])
            assert r["across_size"], (name, zoom, index)
            assert _close(float(r["across_size"]), want), (name, zoom, index, r["across_size"])


@pytest.mark.excel
def test_show_outline_symbols_in_the_view_wins_over_outline_properties_as_captured():
    records = _captured()
    names = ("main", "view_off", "props_off", "view_on_props_off", "view_off_props_on")
    for name in names:
        expected = outline.EXPECTED[name]
        gutters = _gutters(records, name)
        if not gutters:
            continue
        drawn = all(r["across_size"] for r in gutters.values())
        assert drawn is (expected is not None), name
        if expected is None:
            assert not [r for r in records if r["group"] == name and r["axis"] != "gutter"]


@pytest.mark.excel
def test_a_show_outline_symbols_of_yes_reads_as_shown_or_is_refused_as_captured():
    rows = [r for r in _captured() if r["group"] == "flag_yes"]
    if len(rows) == 1 and rows[0]["kind"].startswith("refused: "):
        return  # Excel refuses the file: no value but 0, 1, false and true reaches the screen
    gutters = _gutters(rows, "flag_yes")
    assert gutters
    assert all(r["across_size"] for r in gutters.values())


@pytest.mark.excel
def test_level_buttons_one_a_lane_centered_on_the_headings_as_captured():
    records = _captured()
    for name in ("main", "heights", "summaries"):
        lanes = outline.EXPECTED[name]["lanes"]
        for key, found in _groups(r for r in records if r["group"] == name).items():
            if key[3] != "levels":
                continue
            count = lanes[0] if key[4] == "rows" else lanes[1]
            levels = _level_buttons(found)
            assert len(levels) == count, (key, levels)
            size = float(found[0]["size"])
            for k, (across, along, wide, long) in enumerate(levels):
                assert _close(across, MARGIN + LANE * k), (key, k, levels)
                assert _close(wide, BUTTON), (key, k, levels)
                assert _close(long, BUTTON), (key, k, levels)
                # the headings' band has no gridline term (gutter buttons do: centered)
                assert _close(along, math.ceil((size - BUTTON) / 2)), (key, k, levels)


def _level_buttons(found):
    """The level buttons as (across, along, across size, along size), one per lane.

    The measuring cuts them unlike the gutter's buttons, which stand apart: in the dark
    appearance buttons that touch read as one piece (48 points for four), and in the light one
    each reads as its 1.5-point top highlight and its body. So a piece a whole number of
    buttons wide is split into buttons, and the pieces of one lane are joined.
    """
    lanes = {}
    for r in found:
        if r["kind"] != "level":
            continue
        across, along = float(r["across"]), float(r["along"])
        wide, long = float(r["across_size"]), float(r["along_size"])
        for k in range(max(1, round(wide / LANE))):
            start = across + LANE * k
            box = lanes.setdefault(round(start), [start, along, along + long])
            box[1], box[2] = min(box[1], along), max(box[2], along + long)
    return [
        (start, top, _width(found, start), bottom - top)
        for start, top, bottom in sorted(lanes.values())
    ]


def _width(found, start):
    """The width of the level button whose lane starts at ``start``."""
    for r in found:
        across, wide = float(r["across"]), float(r["across_size"])
        if r["kind"] == "level" and across - 1 <= start < across + wide - 1:
            return wide / max(1, round(wide / LANE))
    return 0.0


def _css_colors(appearance):
    """sheetpeek's outline colors in an appearance, from the HTML backend's stylesheet."""
    from sheetpeek.backends.html._outline import _CSS

    blocks = _CSS.split(".sheetpeek[data-color-scheme=")
    light = dict(re.findall(r"--sheetpeek-outline-([\w-]+): #(\w+);", blocks[0]))
    if appearance == "light":
        return {k: v.upper() for k, v in light.items()}
    dark = dict(re.findall(r"--sheetpeek-outline-([\w-]+): #(\w+);", blocks[1]))
    return {k: v.upper() for k, v in (light | dark).items()}


def _same(found, want):
    if not found:
        return False
    a = [int(found[i : i + 2], 16) for i in (0, 2, 4)]
    b = [int(want[i : i + 2], 16) for i in (0, 2, 4)]
    return all(abs(x - y) <= COLOR_TOLERANCE for x, y in zip(a, b, strict=False))


@pytest.mark.excel
@pytest.mark.parametrize("appearance", outline.APPEARANCES)
def test_the_gutter_colors_in_each_appearance_as_captured(appearance):
    records = [r for r in _captured() if r["appearance"] == appearance and r["group"] == "main"]
    if not records:
        pytest.skip(f"the {appearance} appearance is not captured yet")
    want = _css_colors(appearance)
    gutters = [r for r in records if r["axis"] == "gutter" and r["index"] == "rows"]
    # not the level band, where the light appearance's buttons read with half-point slivers
    # between them (see _level_buttons)
    bars = [r for r in records if r["kind"] in ("bar", "tick", "dot") and r["axis"] != "levels"]
    buttons = [r for r in records if r["kind"] in ("minus", "plus")]
    assert gutters
    assert bars
    assert buttons
    assert all(_same(r["ink"], want["paper"]) for r in gutters), gutters
    assert all(_same(r["ink"], want["ink"]) for r in bars), {r["ink"] for r in bars}
    assert all(_same(r["ink"], want["button"]) for r in buttons), {r["ink"] for r in buttons}
    assert all(_same(r["paper"], want["button-paper"]) for r in buttons)
    assert all(_same(r["stroke"], want["stroke"]) for r in buttons)
