"""The screen suite ``border_conflicts`` (which of two meeting sides Excel draws, §4.12).

Its workbooks are built and read back with sheetpeek as built, and sheetpeek's resolved side on
each edge is the suite's ``EXPECTED`` (``ORDER`` for the style pairs). Its measuring runs on a
synthetic capture with known lines painted. The comparison with
``tests/captures/border_conflicts.tsv`` is skipped until that capture exists.
"""

import zipfile

import openpyxl
import pytest
from capture import border_conflicts as bc
from capture.screen.window import GRIDLINE, MARKER
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from tests.screen_canvas import Canvas


def _place(rendered, edge):
    """The resolved side on an edge of the laid-out grid."""
    letters, row = coordinate_from_string(edge.cell)
    column = column_index_from_string(letters)
    i = next(k for k, r in enumerate(rendered.rows) if r.index == row)
    p = next(k for k, c in enumerate(rendered.columns) if c.index == column)
    borders = rendered.borders
    return {
        "left": lambda: borders.vertical[i][p],
        "right": lambda: borders.vertical[i][p + 1],
        "top": lambda: borders.horizontal[i][p],
        "bottom": lambda: borders.horizontal[i + 1][p],
    }[edge.side]()


def drawn(path, group):
    """What sheetpeek draws on each edge of a group's workbook: ``style RRGGBB`` or ``none``."""
    rendered = sheetpeek.read(path).layout("excel")
    found = {}
    for edge in bc.edges(group):
        side = _place(rendered, edge)
        found[edge.name] = "none" if side is None else f"{side.style} {side.drawn_color}"
    return found


# --- the workbooks


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    folder = tmp_path_factory.mktemp("border_conflicts")
    assert bc.build(folder) == [edge for group in bc.GROUPS for edge in bc.edges(group)]
    return folder


@pytest.mark.parametrize("group", bc.GROUPS)
def test_each_group_reads_back_as_built_and_draws_the_expected_sides(built, group):
    found = drawn(built / f"border_conflicts_{group}.xlsx", group)
    expected = {edge.name: bc.expected(edge) for edge in bc.edges(group)}
    assert found == expected


def test_the_style_pairs_are_every_pair_in_both_orders_and_orientations():
    pairs = [(e.first.style, e.second.style, e.vertical) for e in bc.edges("styles") if e.first]
    assert len(pairs) == len(set(pairs)) == 13 * 12 * 2
    references = [e for e in bc.edges("styles") if e.kind == "ref"]
    assert sorted(e.second.style for e in references if e.vertical) == sorted(bc.STYLES)
    assert sorted(e.second.style for e in references if not e.vertical) == sorted(bc.STYLES)
    assert sorted(bc.ORDER) == sorted(bc.STYLES)


def test_each_edge_is_measured_on_a_cell_of_its_own_side():
    for group in bc.GROUPS:
        places = [(e.cell, e.side) for e in bc.edges(group)]
        assert len(places) == len(set(places)), group


def test_theme_and_indexed_colors_resolve_as_the_cases_state(built):
    sheet = sheetpeek.read(built / "border_conflicts_colors.xlsx")
    for edge in bc.edges("colors"):
        if edge.kind != "color":
            continue
        letters, row = coordinate_from_string(edge.cell)
        side = sheet.cell(row, column_index_from_string(letters)).style.borders.sides[
            ("left", "right", "top", "bottom").index(edge.side)
        ]
        assert side.drawn_color == edge.first.rgb, edge.name


def test_the_crafted_borders_hold_start_and_end(built):
    with zipfile.ZipFile(built / "border_conflicts_start_end.xlsx") as package:
        styles = package.read("xl/styles.xml").decode()
    assert '<border><start style="thick"><color rgb="FF000000"/></start><top/>' in styles
    assert '<start style="thick"><color rgb="FF000000"/></start><left style="thin">' in styles
    assert "<vertical" not in styles
    sheet = sheetpeek.read(built / "border_conflicts_start_end.xlsx")
    assert sheet.cell(4, 3).style.borders.left.style == "thick"
    assert sheet.cell(8, 3).style.borders.left.style == "thick"  # start over left, as Excel


def test_left_before_start_is_refused_as_excel_refuses_it(built, tmp_path):
    source, path = built / "border_conflicts_start_end.xlsx", tmp_path / "left_first.xlsx"
    start = '<start style="thick"><color rgb="FF000000"/></start>'
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(path, "w") as new:
        for item in old.infolist():
            data = old.read(item.filename)
            if item.filename == "xl/styles.xml":
                text = data.decode()
                left = text.index("<left", text.index(start + "<left"))
                end = text.index("</left>", left) + len("</left>")
                text = text[: left - len(start)] + text[left:end] + start + text[end:]
                data = text.encode()
            new.writestr(item, data)
    with pytest.raises(sheetpeek.PackageError, match="left before start"):
        sheetpeek.read(path)


def test_no_case_draws_near_the_marker_color():
    marker = tuple(int(MARKER[k : k + 2], 16) for k in (0, 2, 4))
    for edge in bc.EDGES:
        for line in (edge.first, edge.second):
            if line is None:
                continue
            rgb = tuple(int(line.rgb[k : k + 2], 16) for k in (0, 2, 4))
            assert sum((a - b) ** 2 for a, b in zip(rgb, marker, strict=True)) ** 0.5 > 200


def test_the_color_pairs_separate_what_they_say():
    def luma(c, weights=(0.299, 0.587, 0.114)):
        return sum(w * int(c[k : k + 2], 16) for w, k in zip(weights, (0, 2, 4), strict=True))

    for label, a, b in bc.COLOR_PAIRS:
        la, lb = luma(a.rgb), luma(b.rgb)
        if label == "equal luma":
            assert la == lb, (a, b)
            continue
        assert abs(la - lb) >= 3, (a, b)
        if "709" in label:
            assert (la < lb) != (
                luma(a.rgb, (0.2126, 0.7152, 0.0722)) < luma(b.rgb, (0.2126, 0.7152, 0.0722))
            )
        if "max" in label:
            assert (la < lb) != (max(bytes.fromhex(a.rgb)) < max(bytes.fromhex(b.rgb)))
        if "mean" in label:
            assert (la < lb) != (sum(bytes.fromhex(a.rgb)) < sum(bytes.fromhex(b.rgb)))


def test_the_cells_read_cover_every_column_and_row_and_stay_off_hidden_ones():
    for group in bc.GROUPS:
        found = [coordinate_from_string(c) for c in bc.reads(group)]
        for edge in bc.edges(group):
            letters, row = coordinate_from_string(edge.cell)
            assert letters in {c for c, _ in found}, edge.name
            assert row in {r for _, r in found}, edge.name
        if group == "colors":  # column L and row 13 are hidden
            assert all(c != "L" and r != 13 for c, r in found)
        marker = coordinate_from_string(bc.MARKERS[group][1])
        assert all(
            column_index_from_string(c) < column_index_from_string(marker[0]) for c, _ in found
        )
        assert all(r < marker[1] for _, r in found)


# --- measuring, on a synthetic capture

BLACK = (0, 0, 0)


def _canvas(group):
    ws = openpyxl.Workbook().active
    bc.fill(ws, group)
    return Canvas(ws, bc.reads(group))


def _paint(canvas, edge, rows, dashes=None, rgb=BLACK):
    """A line on an edge: ``rows`` device rows' offsets from the boundary, ``dashes`` the on and
    off runs of one period in pixels (solid without), from the edge's start."""
    letters, row = coordinate_from_string(edge.cell)
    column = next(c for c in canvas.geometry if coordinate_from_string(c)[0] == letters)
    line = next(c for c in canvas.geometry if coordinate_from_string(c)[1] == row)
    x0, _, x1, _ = canvas.box(column)
    _, y0, _, y1 = canvas.box(line)
    boundary = {"left": x0, "right": x1, "top": y0, "bottom": y1}[edge.side]
    start, end = (y0, y1) if edge.vertical else (x0, x1)
    flags = []
    for k, length in enumerate(dashes or (end - start,)):
        flags += [k % 2 == 0] * length
    for offset in rows:
        for b in range(start, end):
            if flags[(b - start) % len(flags)]:
                x, y = (boundary + offset, b) if edge.vertical else (b, boundary + offset)
                canvas.paint(x, y, x + 1, y + 1, rgb)


def _edge(group, name):
    return next(e for e in bc.edges(group) if e.name == name)


def test_a_shared_edge_is_matched_to_the_reference_it_looks_like():
    canvas = _canvas("styles")
    _paint(canvas, _edge("styles", "ref v thin"), (0, 1))
    _paint(canvas, _edge("styles", "ref v thick"), (-2, -1, 0, 1, 2, 3))
    _paint(canvas, _edge("styles", "ref v dashed"), (0, 1), (6, 2))
    _paint(canvas, _edge("styles", "ref v double"), (-2, -1, 2, 3))
    _paint(canvas, _edge("styles", "v thick|double"), (-2, -1, 2, 3))
    _paint(canvas, _edge("styles", "v thin|dashed"), (0, 1))
    _paint(canvas, _edge("styles", "v dashed|thin"), (0, 1), (6, 2))
    records = {r["name"]: r for r in bc.measure(canvas.capture(), "styles")}
    assert records["ref v thick"]["offsets"] == "-1 -0.5 0 0.5 1 1.5"
    assert (records["ref v dashed"]["ons"], records["ref v dashed"]["offs"]) == ("3", "1")
    assert records["ref v dashed"]["pattern"] == "on 3 off 1"
    assert records["v thick|double"]["drawn"] == "double"
    assert records["v thin|dashed"]["drawn"] == "thin"
    assert records["v dashed|thin"]["drawn"] == "dashed"
    assert records["v dashed|thin"]["distance"] == "0.00"
    assert records["h thin/medium"]["drawn"] == "none"  # nothing painted there
    assert set(bc.COLUMNS) >= set(records["v thin|dashed"])


def test_short_dashes_in_a_row_tell_dash_dot_dot_from_dash_dot():
    edge = _edge("styles", "h dashDot/thin")
    found = {}
    for name, dashes in (  # in pixels, at 2 a point
        ("mediumDashed", (18, 6)),
        ("dashDot", (18, 6, 6, 6)),
        ("dashDotDot", (18, 6, 6, 6, 6, 6)),
    ):
        canvas = _canvas("styles")
        _paint(canvas, edge, (0, 1), dashes)
        record = bc.measure_edge(canvas.capture(), edge)
        found[name] = (record["ons"], record["offs"], record["short_seq"])
    assert found == {
        "mediumDashed": ("9", "3", "0"),
        "dashDot": ("3 9", "3", "1"),
        "dashDotDot": ("3 9", "3", "2"),
    }


def test_a_line_takes_its_most_common_ink_and_the_gridline_is_not_ink():
    canvas = _canvas("colors")
    edge = _edge("colors", "v 601 vs 709: 40C080|FF8000")
    _paint(canvas, edge, (0, 1), rgb=(0x40, 0xC0, 0x80))
    other = _edge("colors", "h 601 vs 709: 40C080/FF8000")
    _paint(canvas, other, (0, 1), rgb=GRIDLINE["mac"])
    records = {r["name"]: r for r in bc.measure(canvas.capture(), "colors")}
    assert records[edge.name]["color"] == "40C080"
    assert records[edge.name]["solid"] == "1.00"
    assert records[other.name]["drawn"] == "none"


def test_slanted_gaps_count_as_gaps_by_the_lines_own_color():
    canvas = _canvas("styles")
    edge = _edge("styles", "ref h slantDashDot")
    _paint(canvas, edge, (-2, -1, 0, 1))
    box = bc.bounds(canvas.capture(), edge.cell)
    top = round(box["top"])
    for x in range(round(box["left"]) + 2, round(box["right"]), 24):  # gaps of 1 pt, light
        canvas.paint(x, top - 2, x + 2, top + 2, (120, 120, 120))
    record = bc.measure_edge(canvas.capture(), edge)
    assert record["coverage"] == "1.00"  # the gaps are ink, as luma alone would count them
    assert (record["ons"], record["offs"]) == ("11", "1")
    assert float(record["solid"]) < 1


def test_every_edge_is_measured_on_a_blank_capture():
    for group in bc.GROUPS:
        for record in bc.measure(_canvas(group).capture(), group):
            assert record["drawn"] == "none", record["name"]


# --- the capture (suite border_conflicts)


def _captured():
    """The captured rows by group and edge name, and each group's refusal, if any."""
    path = bc.OUTPUT
    if not path.exists():
        pytest.skip("not captured yet: run python -m capture.border_conflicts run")
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    found: dict[str, dict[str, dict[str, str]]] = {}
    for row in body:
        record = dict(zip(header, row, strict=False))
        if record["name"].startswith("refused:"):
            found[record["group"]] = {"refused": record}
        elif record["zoom"] == str(bc.ZOOM):
            found.setdefault(record["group"], {})[record["name"]] = record
    return found


def _group(group):
    found = _captured()
    if group not in found:
        pytest.skip(f"group {group} not captured yet")
    if "refused" in found[group]:
        pytest.skip(f"Excel refused the file: {found[group]['refused']['name']}")
    return found[group]


def _rgb(text):
    return tuple(bytes.fromhex(text))


def _far(a, b):
    return sum((x - y) ** 2 for x, y in zip(_rgb(a), _rgb(b), strict=True)) ** 0.5


@pytest.mark.excel
@pytest.mark.parametrize("group", bc.GROUPS)
def test_the_references_are_told_apart(group):
    # Each style alone must lie nearer its own record than any other style's by a clear margin,
    # or a shared edge could not be matched: the nearest other reference at 1 or more (one row
    # of difference counts 10, a run length 2, a share of 0.5 one, the short-dash sequence 2 a
    # step). A dash's phase moves a share by up to 0.15 on these edges, about 0.3 of distance:
    # with every style painted as borders.tsv captured it at random phases, the closest two
    # references (dashDot and dashDotDot, medium or not) are 2 apart and every matched pair
    # within 0.5 of its reference.
    captured = _group(group)
    references = {
        orientation: {
            r["name"].split()[-1]: r
            for r in captured.values()
            if r["kind"] == "ref" and r["name"].split()[1] == orientation
        }
        for orientation in ("v", "h")
    }
    for orientation, found in references.items():
        for style, record in found.items():
            others = {s: r for s, r in found.items() if s != style}
            if not others:
                continue
            score, nearest = bc.identify(record, others)[0]
            assert score >= 1, (orientation, style, nearest, score)


MIXED_LINES = {
    "h hair/slantDashDot": "a hair bottom side over a slantDashDot top side draws a line that is "
    "neither (distance 10.68 from slantDashDot, the nearer); the other placement draws "
    "slantDashDot alone (0.08), the order's answer",
}
"""Edges not compared, each with its reason."""


@pytest.mark.excel
def test_the_stronger_style_wins_as_captured():
    # The style order (§4.12, "Conflicts"): every pair of the 13 styles, black, both orders and
    # orientations. A drawn line within 1 of a reference is that reference's style; past 1 it
    # is neither style alone (a mixed line), which the rule says never happens.
    captured = _group("styles")
    wrong, mixed = [], []
    for edge in bc.edges("styles"):
        if edge.kind != "style" or edge.name in MIXED_LINES:
            continue
        record = captured[edge.name]
        if float(record["distance"]) >= 1:
            mixed.append((edge.name, record["drawn"], record["distance"]))
        elif record["drawn"] != bc.expected(edge).split()[0]:
            wrong.append((edge.name, record["drawn"]))
    assert not mixed, mixed
    assert not wrong, wrong


@pytest.mark.excel
@pytest.mark.xfail(
    strict=True,
    reason="16.115: Excel's rule for two colors of one style is not found. Rec. 601 luma, "
    "sheetpeek's, fits 28 of the 68 cases; the larger RRGGBB, the largest channel and the red "
    "channel fit 60, Rec. 709 luma 56; no single measure fits all. A known deviation, not yet "
    "modeled",
)
def test_the_darker_color_wins_as_captured():
    # Rec. 601 luma, equal luma to the larger RRGGBB, theme, tinted and indexed colors compared
    # after resolution (§4.12, "Conflicts"). The drawn color is the nearer of the two sides'
    # (the pairs are at least 150 apart) and within 24 of it: the capture's conversion to sRGB
    # moves a saturated color by a few units.
    captured = _group("colors")
    wrong = []
    for edge in bc.edges("colors"):
        if edge.kind != "color":
            continue
        record = captured[edge.name]
        expected = bc.expected(edge).split()[1]
        other = edge.second.rgb if expected == edge.first.rgb else edge.first.rgb
        color = record["color"]
        if not (_far(color, expected) <= 24 and _far(color, expected) < _far(color, other)):
            wrong.append((edge.name, color, expected))
        assert record["drawn"] == "thin", edge.name
    assert not wrong, wrong


@pytest.mark.excel
def test_black_and_white_against_colors_as_captured():
    # The common case the color pairs leave out: black and white against six colors, both
    # orders and orientations, against sheetpeek's rule (black the darker, white the lighter).
    captured = _group("black_white")
    wrong = []
    for edge in bc.edges("black_white"):
        if edge.kind != "color":
            continue
        record = captured[edge.name]
        expected = bc.expected(edge).split()[1]
        other = edge.second.rgb if expected == edge.first.rgb else edge.first.rgb
        color = record["color"]
        if not (_far(color, expected) <= 24 and _far(color, expected) < _far(color, other)):
            wrong.append((edge.name, color, expected))
        assert record["drawn"] == "thin", edge.name
    assert not wrong, wrong


@pytest.mark.excel
def test_sides_across_a_hidden_column_or_row_compete_as_captured():
    # The visible neighbors' facing sides meet on one edge (§4.12, "Hidden rows and columns"):
    # the hidden cell's own sides (double, thick) are not drawn.
    captured = _group("colors")
    for edge in bc.edges("colors"):
        if edge.kind != "hidden":
            continue
        record = captured[edge.name]
        style, color = bc.expected(edge).split()
        assert record["drawn"] == style, (edge.name, record["drawn"])
        assert float(record["distance"]) < 1, edge.name
        assert _far(record["color"], color) <= 24, (edge.name, record["color"])


@pytest.mark.excel
def test_start_and_end_are_drawn_as_captured():
    # A border's start and end as left and right, over left and right where both are given
    # (§4.12, "Borders"). A refusal of the crafted file skips this test.
    captured = _group("start_end")
    for edge in bc.edges("start_end"):
        if edge.kind == "ref":
            continue
        record = captured[edge.name]
        assert record["drawn"] == bc.expected(edge).split()[0], (edge.name, record["drawn"])
