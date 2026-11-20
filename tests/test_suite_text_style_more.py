"""The screen suite ``text_style_more`` (text styles at other sizes and fonts, ARCHITECTURE.md
§4.13, §4.7 and §3.6), without Excel.

Its workbook is built and read back with sheetpeek as built, and sheetpeek's layout of each case
is the suite's ``EXPECTED``. Its measuring runs on a synthetic capture with known marks. The
comparison with ``tests/captures/text_style_more.tsv`` is skipped until that capture exists.
"""

import openpyxl
import pytest
from capture import text_style_more as suite
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from tests.screen_canvas import Canvas

BLACK = (0, 0, 0)


def _at(cell):
    letters, row = coordinate_from_string(cell)
    return row, column_index_from_string(letters)


def drawn(path, group):
    """What sheetpeek draws for each case of ``group`` in the workbook at ``path``, by cell.

    ``baselines`` from the row's top; ``left`` the first glyph's advance from the cell's left
    edge, ``right`` the end of the widest line from the (last) cell's right edge; ``scale`` the
    drawn size over the font's in sevenths; ``rules`` (kind, lines, offset from the run's
    baseline, top from the row's top, start from the cell's left, end from the right).
    """
    rendered = sheetpeek.read(path, group.capitalize()).layout("excel")
    columns = {c.index: c for c in rendered.columns}
    found = {}
    for case in suite.cases(group):
        r, c = _at(case.cell)
        last = _at(case.last)[1] if case.last else c
        row = next(x for x in rendered.rows if x.index == r)
        runs = [t for t in row.runs if t.source == (r, c)]
        left = columns[c].x_pt
        right = columns[last].x_pt + columns[last].width_pt
        record = {
            "row_height": row.height_pt,
            "text": rendered.cell(r, c).text,
            "lines": len({t.y for t in runs}),
            "baselines": tuple(sorted({t.y - row.y_pt for t in runs})),
            "left": min(t.x for t in runs) - left,
            "right": max(t.x + t.width for t in runs) - right,
            "scale": round(runs[0].scale * 7, 4),
            "rules": tuple(
                (u.kind, u.lines, u.y - t.y, u.y - row.y_pt, u.x0 - left, u.x1 - right)
                for t in runs
                for u in t.rules
            ),
        }
        found[case.cell] = {key: record[key] for key in FIELDS[group]}
    return found


FIELDS = {
    "vert": ("row_height", "text", "baselines"),
    "rules": ("baselines", "rules"),
    "steps": ("scale",),
    "shrunk": ("row_height", "baselines", "left", "scale"),
    "wrap": ("row_height", "lines", "left"),
    "faces": ("baselines",),
    "overflow": ("row_height", "baselines"),
}
"""What the comparison of each group needs of sheetpeek's layout."""


# --- the workbook


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    path = tmp_path_factory.mktemp("text_style_more") / "text_style_more.xlsx"
    return path, suite.build(path)


def test_the_workbook_reads_back_as_built(built):
    path, found = built
    assert found == suite.cases()
    for group in suite.GROUPS:
        sheet = sheetpeek.read(path, group.capitalize())
        for case in suite.cases(group):
            assert sheet.cell(*_at(case.cell)).value is not None, case.name
        assert drawn(path, group) == suite.EXPECTED[group], group


def test_the_cases_are_what_their_names_say(built):
    path, _ = built
    expected = suite.EXPECTED
    # the shrunk texts take the steps their names give, in every row
    for case in suite.cases("shrunk"):
        if " shrunk " in case.name and "/7" in case.name and "20 pt" not in case.name:
            step = int(case.name.split(" shrunk ")[1].split("/")[0])
            assert expected["shrunk"][case.cell]["scale"] == step, case.name
    # the wrapped pairs measure what their keys say, and sheetpeek wraps the wider ones
    import xlformat

    font = xlformat.Font("Calibri", 11)
    for table in (suite.INDENT_NARROW, suite.INDENT_WIDE):
        for width, text in table.items():
            assert xlformat.text_width(text, font=font) == width
    sheet = sheetpeek.read(path, "Rules")
    for row, size in zip((4, 6, 8, 10), suite.RULE_SIZES, strict=True):
        assert sheet.cell(row, 3).font.size == size


def test_every_case_has_a_cell_of_its_own_left_of_and_below_the_markers():
    for group in suite.GROUPS:
        found = suite.cases(group)
        cells = [c.cell for c in found]
        assert len(set(cells)) == len(cells), group
        first, second = suite.MARKERS[group]
        marker_column = column_index_from_string(coordinate_from_string(second)[0])
        marked = {first, second}
        for cell in suite._boxes(found):
            row, column = _at(cell)
            assert column < marker_column, (group, cell)
            assert row > 3, (group, cell)
            assert cell not in marked


def test_few_cells_are_read():
    # Geometry comes from one cell a column and one a row; Excel's reads take about 1.3 s a cell,
    # about four minutes for the seven groups (187 cells).
    total = sum(len(suite.references(group)) for group in suite.GROUPS)
    assert total < 200
    for group in suite.GROUPS:
        refs = suite.references(group)
        assert len(refs) == len(set(refs))


def test_no_case_draws_magenta():
    # Every case is text in the default color on white: no font color or fill is set anywhere.
    for group in suite.GROUPS:
        ws = openpyxl.Workbook().active
        suite.fill(ws, group)
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    color = cell.font.color
                    assert color is None or color.theme == 1, cell.coordinate  # dark 1
                    assert cell.fill.fill_type is None, cell.coordinate


# --- measuring, on a synthetic capture


def _canvas(group, heights=()):
    ws = openpyxl.Workbook().active
    found = suite.fill(ws, group)
    for row, height in heights:
        ws.row_dimensions[row].height = height
    cells = list(dict.fromkeys([*suite._boxes(found), *suite.references(group)]))
    return Canvas(ws, cells)


def _case(group, cell):
    return next(c for c in suite.cases(group) if c.cell == cell)


def test_boxes_are_put_together_from_the_reference_cells():
    ws = openpyxl.Workbook().active
    suite.fill(ws, "steps")
    canvas = Canvas(ws, suite.references("steps"))
    cap = canvas.capture()
    full = Canvas(ws, ["Q17", "C4"]).geometry
    suite.complete(cap, ["Q17", "C4"])
    assert cap.geometry["Q17"] == full["Q17"]
    assert cap.geometry["C4"] == full["C4"]


def test_each_lines_baseline_is_the_bottom_of_its_left_quarter():
    canvas = _canvas("faces")
    x0, y0, _, _ = canvas.box("C10")
    # three lines "Hxg": an H 10 px wide ending on the baseline, then a g 6 px lower, 30 px right
    for top in (20, 60, 100):
        canvas.paint(x0 + 6, y0 + top, x0 + 16, y0 + top + 14, BLACK)
        canvas.paint(x0 + 36, y0 + top + 4, x0 + 46, y0 + top + 20, BLACK)
    measured = suite.measure(canvas.capture(), _case("faces", "C10"))
    assert measured["baselines"] == "17 37 57"
    assert measured["lines"] == "3"
    assert measured["ink_bottom"] == "60"


def test_rules_are_found_with_their_tops_from_the_cells_top():
    canvas = _canvas("rules", heights=[(8, 30)])
    x0, y0, x1, _ = canvas.box("G8")
    canvas.paint(x0 + 6, y0 + 10, x0 + 40, y0 + 40, BLACK)  # the glyphs, baseline at 40 px
    canvas.paint(x0 + 6, y0 + 48, x1 - 4, y0 + 50, BLACK)  # an accounting rule 4 pt below
    measured = suite.measure(canvas.capture(), _case("rules", "G8"))
    assert measured["baseline"] == "20"
    assert measured["rule_offset"] == "4"
    assert measured["rule_tops"] == "24"
    assert measured["rule_thickness"] == "1"


def test_shrink_is_measured_against_its_reference():
    canvas = _canvas("steps")
    x0, y0, _, _ = canvas.box("E10")
    r0, ry, _, _ = canvas.box("M10")
    canvas.paint(x0 + 6, y0 + 12, x0 + 46, y0 + 22, BLACK)
    canvas.paint(r0 + 6, ry + 4, r0 + 76, ry + 24, BLACK)
    measured = suite.measure(canvas.capture(), _case("steps", "E10"))
    assert (measured["width_ratio"], measured["height_ratio"]) == ("0.5714", "0.5")
    assert (measured["ink_width"], measured["ref_width"]) == ("20", "35")


def test_every_case_is_measured_on_a_blank_capture():
    for group in suite.GROUPS:
        cap = _canvas(group).capture()
        for case in suite.cases(group):
            measured = suite.measure(cap, case)
            assert measured is not None, case.name
            assert set(suite.COLUMNS) >= set(measured), case.name


# --- the capture (suite text_style_more)


def _captured(zoom=100):
    path = suite.OUTPUT
    if not path.exists():
        pytest.skip("not captured yet: run python -m capture.text_style_more run")
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    found: dict[str, dict[str, dict[str, str]]] = {}
    for row in body:
        record = dict(zip(header, row, strict=False))
        if record["zoom"] == str(zoom):
            found.setdefault(record["group"], {})[record["cell"]] = record
    return found


def _floats(text):
    return [float(part) for part in text.split()]


def _group(name, zoom=100):
    found = _captured(zoom)
    if name not in found:
        pytest.skip(f"group {name} not captured yet")
    return found[name]


PINNED: dict[str, frozenset[str]] = {
    "test_accounting_underlines_in_merges_wrapped_and_distributed_text": frozenset(
        (
            "doubleAccounting distributed",
            "doubleAccounting wrapped",
            "singleAccounting distributed",
            "singleAccounting wrapped",
        )
    ),
    "test_an_accounting_baseline_in_a_fixed_row": frozenset(
        ("plain fixed 15", "plain fixed 20", "plain fixed 30")
    ),
    "test_baselines_of_sizes_and_faces_xlformat_has_not_measured": frozenset(
        (
            "Arial 11 bottom one line",
            "Arial 11 bottom three lines",
            "Arial 11 center one line",
            "Arial 11 center three lines",
            "Calibri 11 bottom one line",
            "Calibri 11 bottom three lines",
            "Calibri 13 bottom one line",
            "Calibri 13 bottom three lines",
            "Calibri 13 center three lines",
            "Calibri 13 top one line",
            "Calibri 13 top three lines",
            "Calibri 17 bottom one line",
            "Calibri 17 bottom three lines",
            "Calibri 17 center three lines",
            "Calibri 17 top one line",
            "Calibri 17 top three lines",
            "Courier New 11 bottom one line",
            "Courier New 11 bottom three lines",
            "Times New Roman 11 bottom one line",
            "Times New Roman 11 bottom three lines",
            "Times New Roman 11 center one line",
            "Times New Roman 11 center three lines",
            "Times New Roman 11 top three lines",
        )
    ),
    "test_rules_sit_where_sheetpeek_puts_them_at_other_sizes:100": frozenset(
        (
            "double 14",
            "double 20",
            "double 36",
            "doubleAccounting 36",
            "single 36",
            "singleAccounting 36",
            "strike 36",
        )
    ),
    "test_rules_sit_where_sheetpeek_puts_them_at_other_sizes:200": frozenset(
        (
            "double 14",
            "double 20",
            "double 36",
            "double 9",
            "doubleAccounting 20",
            "doubleAccounting 36",
            "single 14",
            "single 36",
            "single 9",
            "singleAccounting 36",
            "strike 14",
            "strike 20",
            "strike 36",
            "strike 9",
        )
    ),
    "test_shrink_steps_at_other_sizes": frozenset(
        (
            "E10",
            "E11",
            "E12",
            "E6",
            "E9",
            "G10",
            "G13",
            "G14",
            "G15",
            "G16",
            "G4",
            "G6",
            "G7",
            "G9",
            "I10",
        )
    ),
    "test_shrunk_text_keeps_the_baseline_sheetpeek_gives_it:100": frozenset(
        (
            "26 plain 11, row 12",
            "26 shrunk 4/7, row 12",
            "center 15 shrunk 2/7",
            "center 15 shrunk 4/7",
            "center 15 shrunk 6/7",
            "center 30 shrunk 2/7",
            "center 30 shrunk 4/7",
            "center 60 shrunk 2/7",
            "center 60 shrunk 4/7",
            "fitted plain 11, row 14",
            "fitted shrunk 4/7 beside 20 pt shrunk",
            "fitted shrunk 4/7, row 14",
            "top 15 shrunk 2/7",
            "top 15 shrunk 4/7",
            "top 15 shrunk 6/7",
            "top 30 shrunk 2/7",
            "top 30 shrunk 4/7",
            "top 30 shrunk 6/7",
            "top 60 shrunk 2/7",
            "top 60 shrunk 4/7",
            "top 60 shrunk 6/7",
        )
    ),
    "test_shrunk_text_keeps_the_baseline_sheetpeek_gives_it:200": frozenset(
        (
            "26 plain 11, row 12",
            "26 shrunk 4/7, row 12",
            "center 15 shrunk 2/7",
            "center 15 shrunk 4/7",
            "center 30 shrunk 2/7",
            "center 30 shrunk 4/7",
            "center 60 shrunk 2/7",
            "center 60 shrunk 4/7",
            "fitted plain 11, row 14",
            "fitted shrunk 4/7 beside 20 pt shrunk",
            "fitted shrunk 4/7, row 14",
            "top 15 shrunk 2/7",
            "top 15 shrunk 4/7",
            "top 15 shrunk 6/7",
            "top 30 shrunk 2/7",
            "top 30 shrunk 4/7",
            "top 30 shrunk 6/7",
            "top 60 shrunk 2/7",
            "top 60 shrunk 4/7",
            "top 60 shrunk 6/7",
        )
    ),
    "test_superscript_and_subscript_fit_rows_and_baselines_as_captured": frozenset(
        (
            "both 14 subscript",
            "both 14 superscript",
            "both 24 subscript",
            "both 24 superscript",
            "both 36 subscript",
            "both 36 superscript",
            "plain 36",
            "subscript 14",
            "subscript 18",
            "subscript 36",
            "subscript 9",
            "superscript 18",
            "superscript 36",
        )
    ),
}
"""Known deviations by test (and zoom), the cases that differ; each comparison asserts exactly
these, so a new difference fails and so does a fixed one (unpin it). Their reasons (the vertical
metrics of styled text, not yet modeled):

- test_accounting_underlines_in_merges_wrapped_and_distributed_text: Excel fits wrapped lines with
  an accounting underline 3 pt taller each (three lines take 57 pt) and draws the one rule under
  the last line there; under distributed words its rule lies a point lower.
- test_an_accounting_baseline_in_a_fixed_row: in a fixed row holding an accounting underline, Excel
  puts the plain bottom-aligned text beside it on the accounting cell's baseline, 6 pt above the
  bottom (sheetpeek 3).
- test_baselines_of_sizes_and_faces_xlformat_has_not_measured: bottom-aligned text in a tall fixed
  row sits higher in Excel (Calibri 11: 5 pt above the bottom of a 60 pt row for one line, 7 in 90
  for three; sheetpeek 3), and sizes and faces xlformat has not measured take other pitches.
- test_lines_that_do_not_fit_their_block_start_where_sheetpeek_starts_them: three wrapped lines in
  a 20 pt row: Excel starts the block where its lines fall, sheetpeek shows one line.
- test_rules_sit_where_sheetpeek_puts_them_at_other_sizes: rule offsets and ends at sizes other
  than 11 (and at 36, past the row table) are scaled from 11 pt; Excel's differ by a point or two.
- test_shrink_steps_at_other_sizes: the step at sizes and faces other than Calibri 11 differs from
  the largest seventh that fits.
- test_shrunk_text_keeps_the_baseline_sheetpeek_gives_it: shrunk text's baseline in Excel is not
  the unshrunk size's that sheetpeek keeps.
- test_superscript_and_subscript_fit_rows_and_baselines_as_captured: superscript and subscript rows
  and baselines at 9 to 36 pt are not the 11 pt growth scaled, and 36 is past the row table.
"""


def _pinned(key, wrong):
    names = set(wrong) if isinstance(wrong, dict) else {w[0] for w in wrong}
    assert names == set(PINNED.get(key, ())), (key, sorted(names ^ set(PINNED.get(key, ()))))


@pytest.mark.excel
def test_superscript_and_subscript_fit_rows_and_baselines_as_captured():
    # Row heights are Excel's own reads (exact). Each case has a row of its own, so the row's
    # height is what that one font asks for: alone at 9, 14, 18 and 36 pt (growth linear
    # between and beyond 11 and 24 pt), and both kinds in one row at 14, 24 and 36 pt (half the
    # one-kind growth again). The baseline of "HHH" is its ink's bottom, on a device pixel: half
    # a point at zoom 100, the tolerance.
    captured = _group("vert")
    wrong = []
    for case in suite.cases("vert"):
        if case.name.startswith("number"):
            continue
        record, expected = captured[case.cell], suite.EXPECTED["vert"][case.cell]
        if float(record["row_height"]) != expected["row_height"]:
            wrong.append((case.name, "row", record["row_height"], expected["row_height"]))
        baseline = _floats(record["baselines"])
        if len(baseline) != 1 or abs(baseline[0] - expected["baselines"][0]) > 0.5:
            wrong.append((case.name, "baseline", record["baselines"], expected["baselines"]))
    _pinned("test_superscript_and_subscript_fit_rows_and_baselines_as_captured", wrong)


@pytest.mark.excel
def test_superscript_numbers_show_what_their_drawn_size_allows():
    # Excel's shown text (its read, exact) for the raised, lowered and plain numbers: digits or
    # "#", as sheetpeek measures them at the drawn size. The plain numbers are the control.
    captured = _group("vert")
    wrong = []
    for case in suite.cases("vert"):
        if not case.name.startswith("number"):
            continue
        record, expected = captured[case.cell], suite.EXPECTED["vert"][case.cell]
        if record["text"] != expected["text"]:
            wrong.append((case.name, record["text"], expected["text"]))
    assert wrong == []


@pytest.mark.excel
@pytest.mark.parametrize("zoom", [100, 200])
def test_rules_sit_where_sheetpeek_puts_them_at_other_sizes(zoom):
    # Each rule's top below the glyphs' baseline (``rule_offset``, a strike's negative):
    # sheetpeek scales the 11 pt offsets with the size, rounded to whole points, and puts a
    # double's second line 2 screen points (2 sheet points at zoom 100, 1 at 200) below the
    # first. The baseline and each rule start on a device pixel, so the offset is within half a
    # point at zoom 100 and a quarter at 200 of where Excel puts it; sheetpeek's own offsets are
    # whole points, so half a point at both. Ends are compared within a point (a device pixel and
    # the 1 pt overrun of single and double underlines).
    captured = _group("rules", zoom)
    wrong = []
    for case in suite.cases("rules"):
        if case.cell not in captured or not any(
            case.name == f"{kind} {size}" for kind in suite.RULE_KINDS for size in suite.RULE_SIZES
        ):
            continue
        record, expected = captured[case.cell], suite.EXPECTED["rules"][case.cell]
        offsets = sorted(_floats(record["rule_offset"]))
        want = []
        for _, lines, offset, _, start, end in expected["rules"]:
            want += [offset, offset + 200 / zoom] if lines == 2 else [offset]
            starts, ends = _floats(record["rule_start"]), _floats(record["rule_end"])
            if any(abs(s - start) > 1 for s in starts) or any(abs(e - end) > 1 for e in ends):
                wrong.append((case.name, "ends", starts, ends, start, end))
        if len(offsets) != len(want) or any(
            abs(a - b) > 0.5 for a, b in zip(offsets, sorted(want), strict=True)
        ):
            wrong.append((case.name, "offsets", offsets, want))
    _pinned(f"test_rules_sit_where_sheetpeek_puts_them_at_other_sizes:{zoom}", wrong)


@pytest.mark.excel
def test_accounting_underlines_in_merges_wrapped_and_distributed_text():
    # A merge's rule spans the merge (from 3 pt after its left gridline to 2 pt before its right
    # one); wrapped text has one rule (two for double) a line; distributed words one rule under
    # the line. Ends within a point; tops within half a point of sheetpeek's.
    captured = _group("rules")
    wrong = []
    for case in suite.cases("rules"):
        if not any(word in case.name for word in ("merge", "wrapped", "distributed")):
            continue
        record, expected = captured[case.cell], suite.EXPECTED["rules"][case.cell]
        tops = _floats(record["rule_tops"])
        want = []
        for _, lines, _, top, _, _ in expected["rules"]:
            want += [top, top + 2] if lines == 2 else [top]
        if len(tops) != len(want) or any(
            abs(a - b) > 0.5 for a, b in zip(sorted(tops), sorted(want), strict=True)
        ):
            wrong.append((case.name, "tops", tops, want))
        _, _, _, _, start, end = expected["rules"][0]
        starts, ends = _floats(record["rule_start"]), _floats(record["rule_end"])
        if any(abs(s - start) > 1 for s in starts) or any(abs(e - end) > 1 for e in ends):
            wrong.append((case.name, "ends", starts, ends, start, end))
    _pinned("test_accounting_underlines_in_merges_wrapped_and_distributed_text", wrong)


@pytest.mark.excel
def test_an_accounting_baseline_in_a_fixed_row():
    # Bottom aligned "HHH" with a single accounting underline in rows fixed at 15, 20 and 30 pt:
    # sheetpeek puts its baseline 6 pt above the bottom, the plain text's 3 pt. Half a point.
    captured = _group("rules")
    wrong = []
    for case in suite.cases("rules"):
        if "fixed" not in case.name:
            continue
        record, expected = captured[case.cell], suite.EXPECTED["rules"][case.cell]
        key = "baseline" if "rules" in case.measure else "baselines"
        found = _floats(record[key])[-1]
        if abs(found - expected["baselines"][-1]) > 0.5:
            wrong.append((case.name, found, expected["baselines"]))
    _pinned("test_an_accounting_baseline_in_a_fixed_row", wrong)


KNOWN_STEPS: dict[str, int] = {}
"""Cells whose captured step differs from sheetpeek's, each with its reason."""


@pytest.mark.excel
def test_shrink_steps_at_other_sizes():
    # The step is the ink width over the unshrunk reference's, in sevenths: Calibri 11's
    # captures put that ratio within 0.04 of k/7 (0.8723 for 6/7 the worst), against 0.14
    # between steps, so the nearest seventh is the step.
    captured = _group("steps")
    wrong = {}
    for case in suite.cases("steps"):
        record = captured[case.cell]
        step = round(float(record["width_ratio"]) * 7)
        expected = suite.EXPECTED["steps"][case.cell]["scale"]
        if step != expected:
            wrong[case.cell] = step
    _pinned("test_shrink_steps_at_other_sizes", wrong)


@pytest.mark.excel
@pytest.mark.parametrize("zoom", [100, 200])
def test_shrunk_text_keeps_the_baseline_sheetpeek_gives_it(zoom):
    # The baseline is the bottom of the ink's left quarter (all "H"); half a point. Unshrunk
    # "HHHg" in the same row and alignment is the control; rows 10 to 16 separate a baseline
    # kept from the unshrunk size, one shared with a larger neighbor and one from the shrunk
    # size's own metrics.
    captured = _group("shrunk", zoom)
    wrong = []
    for case in suite.cases("shrunk"):
        if "baselines" not in case.measure or case.cell not in captured:
            continue
        found = _floats(captured[case.cell]["baselines"])
        expected = suite.EXPECTED["shrunk"][case.cell]["baselines"]
        if len(found) != 1 or abs(found[0] - expected[0]) > 0.5:
            wrong.append((case.name, found, expected))
    _pinned(f"test_shrunk_text_keeps_the_baseline_sheetpeek_gives_it:{zoom}", wrong)


@pytest.mark.excel
def test_when_nothing_fits_shrunk_text_is_clipped_at_its_cell():
    # sheetpeek draws 60 "H"s at 1/7, clipped at the cell's right edge. Excel draws no ink at
    # zoom 100 (a known deviation, pinned: nothing is drawn when even 1/7 does not fit).
    captured = _group("shrunk")
    for cell in ("C18", "O18"):
        record = captured[cell]
        assert record["ink_right"] == record["ink_left"] == "", cell


X_BAND = 5.5
"""Points from the baseline to the top of "aaa"'s ink at Calibri 11, as the capture measures it."""


@pytest.mark.excel
def test_lines_that_do_not_fit_their_block_start_where_sheetpeek_starts_them():
    # Three lines in 15 and 20 pt rows, centered, bottom and top aligned: the bands' bottoms
    # ("aaa", "bbb", "ccc" do not descend; "bbb"'s ascender only rises) against sheetpeek's
    # baselines of the lines that show, half a point. A line whose x-height band starts inside
    # the row shows; one cut by the row's bottom ends there, a half point up (the ink's pixel).
    captured = _group("shrunk")
    wrong = []
    for case in suite.cases("shrunk"):
        if not case.name.startswith("overfull"):
            continue
        record, expected = captured[case.cell], suite.EXPECTED["shrunk"][case.cell]
        bottoms = [
            float(start) + float(length)
            for start, length in (s.split("+") for s in record["line_tops"].split())
        ]
        height = expected["row_height"]
        shown = [min(b, height - 0.5) for b in expected["baselines"] if 0 < b < height + X_BAND]
        if len(bottoms) != len(shown) or any(
            abs(a - b) > 0.5 for a, b in zip(bottoms, shown, strict=True)
        ):
            wrong.append((case.name, bottoms, shown))
    _pinned("test_lines_that_do_not_fit_their_block_start_where_sheetpeek_starts_them", wrong)


@pytest.mark.excel
def test_an_indent_and_wrapped_lines_as_captured():
    # The number of lines tells whether an indent narrows the width a wrapped line may take
    # (sheetpeek breaks at the full width); the ink's left (left aligned) or right (right
    # aligned) edge, within a point (a glyph's side bearing and a device pixel), where the
    # first line starts or ends.
    captured = _group("wrap")
    wrong = []
    for case in suite.cases("wrap"):
        if not case.name.startswith("wrap "):
            continue
        record, expected = captured[case.cell], suite.EXPECTED["wrap"][case.cell]
        if int(record["lines"]) != expected["lines"]:
            wrong.append((case.name, "lines", record["lines"], expected["lines"]))
        if " left " in case.name and abs(float(record["ink_left"]) - expected["left"]) > 1:
            wrong.append((case.name, "left", record["ink_left"], expected["left"]))
    assert wrong == []


@pytest.mark.excel
def test_justify_and_the_vertical_spreads_in_fitting():
    # Excel's row height (its read) for each spread alone in a fitted row, against sheetpeek's.
    captured = _group("wrap")
    wrong = []
    for case in suite.cases("wrap"):
        if not case.name.startswith("fitting"):
            continue
        found = float(captured[case.cell]["row_height"])
        expected = suite.EXPECTED["wrap"][case.cell]["row_height"]
        if found != expected:
            wrong.append((case.name, found, expected))
    assert wrong == []


@pytest.mark.excel
def test_baselines_of_sizes_and_faces_xlformat_has_not_measured():
    # Each line's baseline (the bottom of its "H") against sheetpeek's, which takes the nearest
    # measured size and Calibri's metrics for other faces; Calibri 11 is the control. Half a
    # point, a device pixel at zoom 100.
    captured = _group("faces")
    wrong = []
    for case in suite.cases("faces"):
        found = _floats(captured[case.cell]["baselines"])
        expected = suite.EXPECTED["faces"][case.cell]["baselines"]
        if len(found) != len(expected) or any(
            abs(a - b) > 0.5 for a, b in zip(found, expected, strict=True)
        ):
            wrong.append((case.name, found, expected))
    _pinned("test_baselines_of_sizes_and_faces_xlformat_has_not_measured", wrong)


CAP_HEIGHT = 0.644
"""Calibri's cap height in ems: a line whose "H" starts above the row's bottom shows."""


@pytest.mark.excel
def test_lines_that_overflow_their_row_sit_where_sheetpeek_puts_them():
    # §3.6: three and two wrapped lines in rows of 15 to 32 pt, a line taller than its 15 pt row,
    # and vertical justify and distributed in a 20 pt row: each visible line's baseline (the
    # ink's bottom under its "H"), half a point. A line cut by the row's bottom reads there.
    captured = _group("overflow")
    wrong = []
    for case in suite.cases("overflow"):
        record, expected = captured[case.cell], suite.EXPECTED["overflow"][case.cell]
        height = expected["row_height"]
        size = 20 if case.name.startswith("tall 20") else 14 if case.name.startswith("tall") else 11
        cap = CAP_HEIGHT * size
        shown = [min(b, height - 0.5) for b in expected["baselines"] if b - cap < height and b > 0]
        found = [float(b) for b in record["baselines"].split()] if record["baselines"] else []
        if len(found) != len(shown) or any(
            abs(a - b) > 0.5 for a, b in zip(found, shown, strict=True)
        ):
            wrong.append((case.name, found, shown))
    _pinned("test_lines_that_overflow_their_row_sit_where_sheetpeek_puts_them", wrong)
