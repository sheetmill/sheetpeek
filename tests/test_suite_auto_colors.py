"""The screen suite ``auto_colors`` (automatic and system colors, ARCHITECTURE.md §3.3, §4.12,
§4.14), without Excel.

Its workbooks are built and read back with sheetpeek as built, and what sheetpeek draws for each
case is the suite's ``EXPECTED``. Its measuring runs on a synthetic capture with known marks. The
comparison with ``tests/captures/auto_colors.tsv`` is skipped until that capture exists.
"""

import zipfile

import openpyxl
import pytest
from capture import auto_colors
from capture.screen.window import MARKER
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from tests.screen_canvas import SCALE, Canvas

BLACK, WHITE, BLUE, YELLOW = (0, 0, 0), (255, 255, 255), (0, 0, 255), (255, 255, 0)
GRID = (212, 212, 212)


def _rgb(text):
    return tuple(int(text[k : k + 2], 16) for k in (0, 2, 4))


def _distance(a, b):
    return sum((x - y) ** 2 for x, y in zip(_rgb(a), _rgb(b), strict=True)) ** 0.5


def _cell(sheet, address):
    letters, row = coordinate_from_string(address)
    return sheet.cell(row, column_index_from_string(letters))


def drawn(path, group):
    """What sheetpeek draws for each case of the workbook at ``path``, as ``EXPECTED`` holds it."""
    sheet = sheetpeek.read(path)
    found = {}
    for case in auto_colors.cases(group):
        style = _cell(sheet, case.cell).style
        if case.kind == "fill":
            spec = style.fill_spec
            if spec is None:
                found[case.name] = "none"
            elif spec.kind == "solid":
                found[case.name] = f"solid {spec.foreground}"
            else:
                found[case.name] = f"{spec.kind} {spec.foreground}/{spec.background}"
        elif case.kind == "font":
            found[case.name] = style.font.color or "000000"
        else:
            b = style.borders
            sides = {f"{s.style} {s.drawn_color}" for s in (b.left, b.right, b.top, b.bottom)}
            assert len(sides) == 1, case.name
            found[case.name] = sides.pop()
    return found


# --- the workbooks


def test_the_auto_colors_workbooks_read_back_as_built(tmp_path):
    found = auto_colors.build(tmp_path)
    assert found == auto_colors.CASES
    assert len({c.name for c in found}) == len(found)
    assert set(auto_colors.EXPECTED) == {c.name for c in found}
    for group in auto_colors.GROUPS:
        path = tmp_path / f"auto_colors_{group}.xlsx"
        sheet = sheetpeek.read(path)
        for case in auto_colors.cases(group):
            cell = _cell(sheet, case.cell)
            assert cell.value == ("HH" if case.kind == "font" else None), case.name
        expected = {c.name: auto_colors.EXPECTED[c.name] for c in auto_colors.cases(group)}
        assert drawn(path, group) == expected, group


def test_each_color_is_written_as_the_case_states_it(tmp_path):
    auto_colors.build(tmp_path)
    styles = {}
    for group in auto_colors.GROUPS:
        with zipfile.ZipFile(tmp_path / f"auto_colors_{group}.xlsx") as package:
            styles[group] = package.read("xl/styles.xml").decode()
        assert "<colors>" not in styles[group]  # the default palette applies
        assert 'rgb="FF0A' not in styles[group]  # every placeholder rewritten
    main, empty = styles["main"], styles["empty"]
    assert '<patternFill patternType="solid"><bgColor rgb="FFFFFF00" /></patternFill>' in main
    assert '<patternFill patternType="solid"></patternFill>' in main
    assert '<patternFill patternType="lightGrid"></patternFill>' in main
    assert '<left style="medium"></left>' in main
    assert '<fgColor indexed="81" />' in main
    assert "<color/>" not in main
    assert "<fgColor/>" not in main
    assert '<patternFill patternType="solid"><fgColor/></patternFill>' in empty
    assert '<fgColor rgb="FF0000FF" /><bgColor/>' in empty
    assert '<left style="medium"><color/></left>' in empty
    assert '<family val="2" /><color/><sz val="11" />' in empty


def test_every_case_has_a_cell_of_its_own_with_empty_cells_around_it():
    for group in auto_colors.GROUPS:
        found = auto_colors.cases(group)
        cells = [coordinate_from_string(c.cell) for c in found]
        assert len(set(cells)) == len(cells), group
        for letters, row in cells:
            assert row % 2 == 0
            assert letters in auto_colors.SLOTS
        letters, row = coordinate_from_string(auto_colors.MARKERS[group][1])
        assert row > max(r for _, r in cells), group
        last = max(column_index_from_string(c) for c, _ in cells)
        assert column_index_from_string(letters) > last, group


def test_no_case_draws_near_the_marker_color():
    marker = MARKER
    for color in (auto_colors.BLUE, auto_colors.YELLOW, auto_colors.GRAY, "000000", "FFFFFF"):
        assert _distance(color, marker) > 200


# --- measuring, on a synthetic capture


def _canvas(group="main"):
    ws = openpyxl.Workbook().active
    auto_colors.fill(ws, group)
    return Canvas(ws, auto_colors._cells(group))


def _case(name):
    return next(c for c in auto_colors.CASES if c.name == name)


def _gridlines(canvas, cell):
    x0, y0, x1, y1 = canvas.box(cell)
    canvas.paint(x0, y0 - 4, x0 + 2, y1 + 4, GRID)
    canvas.paint(x1, y0 - 4, x1 + 2, y1 + 4, GRID)
    canvas.paint(x0 - 4, y0, x1 + 4, y0 + 2, GRID)
    canvas.paint(x0 - 4, y1, x1 + 4, y1 + 2, GRID)


def test_a_solid_fill_is_its_color_inside_and_covers_its_gridlines():
    canvas = _canvas()
    filled, empty = _case("solid blue"), _case("solid fg indexed 64")
    _gridlines(canvas, empty.cell)
    _gridlines(canvas, filled.cell)
    x0, y0, x1, y1 = canvas.box(filled.cell)
    canvas.paint(x0, y0, x1 + 2, y1 + 2, BLUE)  # Excel's fill runs over the far gridlines
    cap = canvas.capture()
    assert auto_colors.measure(cap, filled) == {"interior": "0000FF:1", "gridlines": "0000"}
    assert auto_colors.measure(cap, empty) == {"interior": "FFFFFF:1", "gridlines": "1111"}


def test_a_pattern_gives_its_two_colors_by_share():
    canvas = _canvas()
    case = _case("lightGrid blue on yellow")
    x0, y0, x1, y1 = canvas.box(case.cell)
    canvas.paint(x0, y0, x1, y1, YELLOW)
    # lightGrid: a 4 by 4 point tile, its top row and left column set (7 of 16 points)
    for y in range(y0, y1):
        for x in range(x0, x1):
            px, py = (x // SCALE) % 4, (y // SCALE) % 4
            if px == 0 or py == 0:
                canvas.paint(x, y, x + 1, y + 1, BLUE)
    measured = auto_colors.measure(canvas.capture(), case)
    shares = dict(part.split(":") for part in measured["interior"].split())
    assert list(shares) == ["FFFF00", "0000FF"]
    assert float(shares["0000FF"]) == pytest.approx(7 / 16, abs=0.02)


def test_a_font_takes_the_color_of_its_ink_farthest_from_the_cell():
    canvas = _canvas()
    case = _case("font indexed 65 over gray")
    x0, y0, x1, y1 = canvas.box(case.cell)
    canvas.paint(x0, y0, x1, y1, (128, 128, 128))
    canvas.paint(x0 + 6, y0 + 14, x0 + 30, y0 + 34, (200, 200, 200))  # anti-aliased edges
    canvas.paint(x0 + 6, y0 + 14, x0 + 30, y0 + 18, WHITE)  # a core of a fifth
    measured = auto_colors.measure(canvas.capture(), case)
    assert measured["ink_color"] == "FFFFFF"
    assert measured["ink_pixels"] == str(24 * 20)
    blank = auto_colors.measure(_canvas().capture(), _case("font auto"))
    assert (blank["ink_color"], blank["ink_pixels"]) == ("none", "0")


def test_a_border_is_measured_across_each_edge():
    canvas = _canvas()
    case = _case("border blue")
    x0, y0, x1, y1 = canvas.box(case.cell)
    canvas.paint(x0 - 2, y0 - 2, x0 + 2, y1 + 2, BLUE)  # left, 2 pt across the boundary
    canvas.paint(x0 - 2, y1 - 2, x1 + 2, y1 + 2, BLUE)  # bottom
    measured = auto_colors.measure(canvas.capture(), case)
    assert measured["edge_left"] == "FFFFFFx4 0000FFx4 FFFFFFx4"
    assert measured["edge_bottom"] == "FFFFFFx4 0000FFx4 FFFFFFx4"
    assert measured["edge_top"] == measured["edge_right"] == "FFFFFFx12"
    assert _line(measured["edge_left"], "0") == "0000FF"
    assert _line(measured["edge_top"], "1") == "none"
    assert _line(measured["edge_top"], "0") == "FFFFFF"


def test_every_case_is_measured_on_a_blank_capture():
    for group in auto_colors.GROUPS:
        cap = _canvas(group).capture()
        for case in auto_colors.cases(group):
            measured = auto_colors.measure(cap, case)
            assert set(auto_colors.COLUMNS) >= set(measured)
            assert measured["gridlines"] == "0000"  # a blank canvas has none
            assert measured["interior"] == "FFFFFF:1"


# --- the capture (suite auto_colors)


def _captured():
    """The captured rows by group and case name, and each group's refusal, if any."""
    path = auto_colors.OUTPUT
    if not path.exists():
        pytest.skip("not captured yet: run python -m capture.auto_colors run")
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    found: dict[str, dict[str, dict[str, str]]] = {}
    for row in body:
        record = dict(zip(header, row, strict=False))
        if record["name"].startswith("refused:"):
            found[record["group"]] = {"refused": record}
        elif record["zoom"] == str(auto_colors.ZOOM):
            found.setdefault(record["group"], {})[record["name"]] = record
    return found


def _records(kind):
    """Each captured case of ``kind`` with what sheetpeek draws, groups Excel refused left out."""
    found = []
    for group, records in _captured().items():
        if "refused" in records:
            continue  # the refusal is the answer; no case of that group is drawn
        found.extend(
            (case, records[case.name], auto_colors.EXPECTED[case.name])
            for case in auto_colors.cases(group)
            if case.kind == kind
        )
    if not found:
        pytest.skip(f"no {kind} case captured")
    return found


def _shares(text):
    return {color: float(share) for color, share in (p.split(":") for p in text.split())}


def _line(runs, gridline):
    """The color an edge's strip draws: its longest run that is neither paper nor gridline.

    With none, the edge is white where the gridline is hidden (``gridline`` ``0``), else bare.
    """
    best, length = None, 0
    for run in runs.split():
        color, _, count = run.partition("x")
        if _distance(color, "FFFFFF") <= 8 or _distance(color, "D4D4D4") <= 8:
            continue
        if int(count) > length:
            best, length = color, int(count)
    if best is not None:
        return best
    return "FFFFFF" if gridline == "0" else "none"


# Tolerances. The fills suite's cells of the same size (12 characters by 24 pt, 4 pixels in)
# measured each pattern's share exactly (lightGrid 0.4375) and solid colors to the unit
# (0,0,255) after the sRGB conversion: 0.01 of share and a distance of 8 leave room for the
# color profile's rounding and nothing more, while every color the cases draw is more than 100
# from every other. A font's ink is anti-aliased; the tenth of its pixels farthest from the
# cell's color is the stems' core, whole device pixels at 2 a point: 40 takes the profile and
# the hinting, and still tells black, white, gray and blue apart.
COLOR = 8
SHARE = 0.01
INK = 40


def _agreeing(case, expected):
    """The record a capture that agrees with sheetpeek would hold for ``case``."""
    record = {"group": case.group, "name": case.name, "kind": case.kind, "cell": case.cell}
    record |= {"zoom": str(auto_colors.ZOOM), "gridlines": "0000"}
    if expected == "none":
        return record | {"interior": "FFFFFF:1", "gridlines": "1111"}
    if case.kind == "fill":
        kind, colors = expected.split()
        fg, _, bg = colors.partition("/")
        if kind == "solid" or fg == bg:
            return record | {"interior": f"{fg}:1"}
        return record | {"interior": f"{bg}:0.5625 {fg}:0.4375"}
    if case.kind == "font":
        return record | {"interior": "FFFFFF:0.9", "ink_color": expected, "ink_pixels": "300"}
    color = expected.split()[1]
    strip = "FFFFFFx12" if color == "FFFFFF" else f"FFFFFFx4 {color}x4 FFFFFFx4"
    return record | {f"edge_{edge}": strip for edge in auto_colors.EDGES}


def test_the_comparison_accepts_a_capture_that_agrees_with_sheetpeek(tmp_path, monkeypatch):
    path = tmp_path / "auto_colors.tsv"
    lines = ["# suite: auto_colors", "\t".join(auto_colors.COLUMNS)]
    for case in auto_colors.CASES:
        if case.group == "empty":
            continue
        record = _agreeing(case, auto_colors.EXPECTED[case.name])
        lines.append("\t".join(record.get(c, "") for c in auto_colors.COLUMNS))
    lines.append("\t".join(["empty", "refused: Excel could not open the file"]))
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    monkeypatch.setattr(auto_colors, "OUTPUT", path)
    test_fills_in_automatic_and_system_colors_draw_as_sheetpeek_draws_them()
    test_fonts_in_automatic_and_system_colors_draw_as_sheetpeek_draws_them()
    test_borders_in_automatic_and_system_colors_draw_as_sheetpeek_draws_them()
    # and refuses one that does not: indexed 64 drawn as a black solid fill
    lines = [
        line.replace("FFFFFF:1\t1111", "000000:1\t0000")
        if "solid fg indexed 64\t" in line
        else line
        for line in lines
    ]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    with pytest.raises(AssertionError):
        test_fills_in_automatic_and_system_colors_draw_as_sheetpeek_draws_them()


@pytest.mark.excel
def test_fills_in_automatic_and_system_colors_draw_as_sheetpeek_draws_them():
    for case, record, expected in _records("fill"):
        shares = _shares(record["interior"])
        top = max(shares, key=shares.get)
        if expected == "none":
            # no fill: the paper and every gridline show (a white fill hides them)
            assert record["gridlines"] == "1111", case.name
            assert _distance(top, "FFFFFF") <= COLOR, case.name
            assert shares[top] >= 1 - SHARE, case.name
            continue
        assert record["gridlines"] == "0000", (case.name, "a fill covers its gridlines")
        kind, colors = expected.split()
        fg, _, bg = colors.partition("/")
        if kind == "solid" or fg == bg:
            assert _distance(top, fg) <= COLOR, (case.name, record["interior"])
            assert shares[top] >= 1 - SHARE, (case.name, record["interior"])
            continue
        assert len(shares) == 2, (case.name, record["interior"])
        near = {name: next(c for c in shares if _distance(c, name) <= COLOR) for name in (fg, bg)}
        assert shares[near[fg]] == pytest.approx(auto_colors.GRID_SHARE, abs=SHARE), case.name
        assert shares[near[bg]] == pytest.approx(1 - auto_colors.GRID_SHARE, abs=SHARE), case.name


@pytest.mark.excel
def test_fonts_in_automatic_and_system_colors_draw_as_sheetpeek_draws_them():
    for case, record, expected in _records("font"):
        assert record["ink_color"] != "none", (case.name, "no text drawn")
        assert _distance(record["ink_color"], expected) <= INK, (case.name, record["ink_color"])


@pytest.mark.excel
def test_borders_in_automatic_and_system_colors_draw_as_sheetpeek_draws_them():
    for case, record, expected in _records("border"):
        _, color = expected.split()
        for k, edge in enumerate(auto_colors.EDGES):
            line = _line(record[f"edge_{edge}"], record["gridlines"][k])
            assert line != "none", (case.name, edge, "no line drawn")
            assert _distance(line, color) <= COLOR, (case.name, edge, line)
