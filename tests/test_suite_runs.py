"""The screen suite ``runs`` (rich-text runs, ARCHITECTURE.md §4.16), without Excel.

Its workbooks are built and read back with sheetpeek as built, and sheetpeek's layout of each
case is the suite's ``EXPECTED``. Its measuring runs on a synthetic capture with known marks.
The comparison with ``tests/captures/runs.tsv`` is skipped until that capture exists.
"""

import zipfile

import openpyxl
import pytest
from capture import runs
from capture.screen.ink import nearest
from capture.screen.window import MARKER
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from tests.screen_canvas import SCALE, Canvas

BLACK, RED, BLUE, GREEN = (0, 0, 0), (255, 0, 0), (0, 0, 255), (0, 176, 80)
PALETTE = {"000000": BLACK, "FF0000": RED, "0000FF": BLUE, "00B050": GREEN}


def label(font):
    """A font as the suite's references name it: ``Calibri 11 bold superscript``."""
    found = f"{font.name} {font.size:g}"
    if font.bold:
        found += " bold"
    if font.vert_align:
        found += f" {font.vert_align}"
    return found


def drawn(path):
    """What sheetpeek draws for each case of the workbook at ``path``, as ``EXPECTED`` holds it."""
    group = path.stem.removeprefix("runs_")
    rendered = sheetpeek.read(path).layout("excel")
    columns = {c.index: c for c in rendered.columns}
    found = {}
    for case in runs.cases(group):
        letters, r = coordinate_from_string(case.cell)
        c = column_index_from_string(letters)
        row = next(x for x in rendered.rows if x.index == r)
        lines = [t for t in row.runs if t.source == (r, c)]
        words = []
        for line in lines:
            pieces = line.spans or (
                sheetpeek.TextSpan(line.text, line.font),  # one font: the run's color is drawn
            )
            for span in pieces:
                color = (span.font.color if line.spans else line.color) or "000000"
                words += [(label(span.font), color, span.shift)] * len(span.text.split())
        record = {
            "row_height": row.height_pt,
            "text": " ".join(line.text for line in lines),
            "baselines": tuple(line.y - row.y_pt for line in lines),
        }
        if "words" in case.measure:
            record["words"] = tuple(words)
        if "gridlines" in case.measure:
            (cell,) = [x for x in row.cells if x.column == c]
            column = columns[c]
            (line,) = lines
            record["reach"] = line.x + line.width - (column.x_pt + column.width_pt)
            record["spill"] = cell.spill[1]
        found[case.cell] = record
    return found


# --- the workbooks


def test_the_runs_workbooks_read_back_as_built(tmp_path):
    found = runs.build(tmp_path)
    assert found == runs.CASES
    for group in runs.GROUPS:
        path = tmp_path / f"runs_{group}.xlsx"
        sheet = sheetpeek.read(path)
        for case in runs.cases(group):
            letters, row = coordinate_from_string(case.cell)
            cell = sheet.cell(row, column_index_from_string(letters))
            assert cell.value == case.plain, case.name
            if case.runs and group != "text_and_runs":
                assert cell.runs is not None, case.name
        assert drawn(path) == runs.EXPECTED[group], group


def test_each_group_holds_its_rich_strings_as_excel_writes_them(tmp_path):
    runs.build(tmp_path)
    with zipfile.ZipFile(tmp_path / "runs_text_and_runs.xlsx") as package:
        strings = package.read("xl/sharedStrings.xml").decode()
        sheet = package.read("xl/worksheets/sheet1.xml").decode()
    assert '<si><t xml:space="preserve">HH </t><r><rPr><rFont val="Calibri"/><b/>' in strings
    assert '<c r="C4" s="1" t="s"><v>0</v></c>' in sheet  # plain references stay inline
    with zipfile.ZipFile(tmp_path / "runs_normal.xlsx") as package:
        styles = package.read("xl/styles.xml").decode()
        strings = package.read("xl/sharedStrings.xml").decode()
    assert styles.index('<font><sz val="14"/><name val="Arial"/>') < styles.index("</font>")
    assert "<r><t" in strings  # a run without properties
    assert "<r><rPr><b/></rPr>" in strings
    assert '<r><rPr><sz val="11"/></rPr>' in strings
    assert '<r><rPr><rFont val="Calibri"/></rPr>' in strings
    assert runs.normal_font(styles) == styles  # once is enough


def test_every_case_has_a_row_and_a_cell_of_its_own():
    for group in runs.GROUPS:
        found = runs.cases(group)
        rows = [coordinate_from_string(c.cell)[1] for c in found]
        assert len(set(rows)) == len(rows), group
        marker = runs.MARKERS[group][1]
        letters, row = coordinate_from_string(marker)
        assert all(r < row for r in rows), group
        last = max(
            column_index_from_string(coordinate_from_string(c)[0]) for c in runs._cells(group)
        )
        assert column_index_from_string(letters) > last, group


def test_no_case_draws_near_the_marker_color():
    marker = tuple(int(MARKER[k : k + 2], 16) for k in (0, 2, 4))
    for color in (runs.RED, runs.BLUE, runs.GREEN):
        rgb = tuple(int(color[k : k + 2], 16) for k in (0, 2, 4))
        assert sum((a - b) ** 2 for a, b in zip(rgb, marker, strict=True)) ** 0.5 > 200


# --- measuring, on a synthetic capture


def _canvas(group, text=(), heights=()):
    ws = openpyxl.Workbook().active
    runs.fill(ws, group)
    for row, height in heights:
        ws.row_dimensions[row].height = height
    reads = [(cell, "geometry.text", value) for cell, value in text]
    return Canvas(ws, runs._cells(group), reads)


def _case(cell, group="main"):
    return next(c for c in runs.cases(group) if c.cell == cell)


def test_words_are_split_at_the_widest_gaps_with_their_size_color_and_area():
    canvas = _canvas("main", [("J22", "HH HH HH")])
    x0, y0, _, _ = canvas.box("J22")
    # three words, each two stems 2 px apart; the words 6 px apart; the middle one raised
    for k, (rgb, top, bottom) in enumerate(((BLACK, 14, 28), (RED, 8, 18), (BLUE, 14, 28))):
        left = x0 + 6 + 20 * k
        for stem in (left, left + 6):
            canvas.paint(stem, y0 + top, stem + 4, y0 + bottom, rgb)
    measured = runs.measure(canvas.capture(), _case("J22"))
    assert measured["words"] == "3+5 13+5 23+5"
    assert measured["word_tops"] == "7 4 7"
    assert measured["word_bottoms"] == "14 9 14"
    assert measured["word_colors"] == "000000 FF0000 0000FF"
    assert measured["word_areas"] == "28 20 28"  # 2 stems of 4 by 14 px, over 2 px a point squared
    assert (measured["ink_left"], measured["ink_top"], measured["ink_bottom"]) == ("3", "4", "14")


def test_a_word_takes_the_color_of_its_darkest_ink():
    canvas = _canvas("main", [("J16", "HH")])
    x0, y0, _, _ = canvas.box("J16")
    canvas.paint(x0 + 6, y0 + 10, x0 + 30, y0 + 30, (255, 150, 150))  # light edges
    canvas.paint(x0 + 6, y0 + 10, x0 + 30, y0 + 12, RED)  # a core of a tenth
    canvas.paint(x0 + 6, y0 + 12, x0 + 30, y0 + 14, (200, 40, 40))
    measured = runs.measure(canvas.capture(), _case("J16"))
    assert measured["words"] == "3+12"
    assert measured["word_colors"] == "FF0000"


def test_wrapped_lines_are_bands_of_ink_from_the_cells_top():
    canvas = _canvas("main", heights=[(10, 59)])
    x0, y0, _, _ = canvas.box("C10")
    for top, height in ((16, 14), (54, 26), (98, 14)):
        canvas.paint(x0 + 6, y0 + top, x0 + 40, y0 + top + height, BLACK)
    measured = runs.measure(canvas.capture(), _case("C10"))
    assert measured["lines"] == "3"
    assert measured["line_tops"] == "8+7 27+13 49+7"


def test_a_gridline_shows_or_is_hidden_at_each_boundary_the_text_may_cross():
    canvas = _canvas("main")
    x0, y0, _, y1 = canvas.box("C4")
    gray = (212, 212, 212)
    for cell in ("F4", "G4", "H4"):  # gridlines at the right of F, G and H, then of E
        right = canvas.box(cell)[2]
        canvas.paint(right, y0, right + 1, y1, gray)
    canvas.paint(canvas.box("E4")[2], y0, canvas.box("E4")[2] + 1, y1, gray)
    canvas.paint(x0 + 6, y0 + 20, canvas.box("D4")[2] + 20, y0 + 34, BLACK)
    measured = runs.measure(canvas.capture(), _case("C4"))
    assert measured["hidden_gridlines"] == "110000"
    assert measured["ink_right"] == f"{(canvas.box('D4')[2] + 20 - canvas.box('I4')[2]) / SCALE:g}"


def test_every_case_is_measured_on_a_blank_capture():
    for group in runs.GROUPS:
        cap = _canvas(group).capture()
        for case in runs.cases(group):
            measured = runs.measure(cap, case)
            assert measured is not None, case.name
            assert set(runs.COLUMNS) >= set(measured)
            if "gridlines" in case.measure:
                assert measured["hidden_gridlines"] == "111111"  # a blank canvas has none


# --- the capture (suite runs)


def _captured():
    """The captured rows by group and cell, and each group's refusal, if any."""
    path = runs.OUTPUT
    if not path.exists():
        pytest.skip("not captured yet: run python -m capture.runs run")
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    found: dict[str, dict[str, dict[str, str]]] = {}
    for row in body:
        record = dict(zip(header, row, strict=False))
        if record["name"].startswith("refused:"):
            found[record["group"]] = {"refused": record}
        elif record["zoom"] == str(runs.ZOOM):
            found.setdefault(record["group"], {})[record["cell"]] = record
    return found


def _spans(text):
    return [tuple(map(float, part.split("+"))) for part in text.split()]


def _floats(text):
    return [float(part) for part in text.split()]


PAST_THE_TABLE = {"C12": {"row_height": 70.0, "baselines": (15.0, 47.0, 67.0)}}
"""A 28 pt run in a wrapped line, past the row table's 24 pt: sheetpeek gives it 24 pt's line and
baseline (row 64, baselines 15, 43, 61), where Excel's line is 4 pt taller with its baseline 4 pt
lower. A known deviation of sizes past 24 (ARCHITECTURE §4.7), pinned here at Excel's values so a
change on either side shows."""


SMALL_RUNS_BASELINE = {"C14": (18.0,)}
"""9 pt runs in a 20 pt cell: the row is the cell font's 26 pt in both, but Excel puts the runs'
baseline 8 pt above the bottom where sheetpeek puts it 3 pt above (23), a 9 pt line's place. Where
Excel places small runs in a taller row is not modeled, a known deviation (ARCHITECTURE §4.16)."""


@pytest.mark.excel
def test_runs_fit_rows_as_captured():
    # Excel's row height (its own read, exact) against sheetpeek's for every case of the main
    # group: each line takes the height of its largest run (C10: 16 + 27 + 16), a fitted row
    # counts the runs' fonts and not the cell's (C14), superscript and subscript runs (J22, J24).
    captured = _captured()["main"]
    for cell, expected in runs.EXPECTED["main"].items():
        if cell in PAST_THE_TABLE:
            assert float(captured[cell]["row_height"]) == PAST_THE_TABLE[cell]["row_height"]
            assert expected["row_height"] != PAST_THE_TABLE[cell]["row_height"], "unpin " + cell
            continue
        assert float(captured[cell]["row_height"]) == expected["row_height"], cell


@pytest.mark.excel
def test_wrapped_and_mixed_lines_sit_on_the_baselines_excel_draws():
    # No letter of these texts descends, so each ink band ends on its line's baseline. The ink of
    # a glyph ends on a device pixel: half a point at zoom 100 on a 2x display, the tolerance.
    captured = _captured()["main"]
    for cell in ("C8", "C10", "C12", "C14", "J18", "J20"):
        bands = _spans(captured[cell]["line_tops"])
        bottoms = [top + height for top, height in bands]
        expected = runs.EXPECTED["main"][cell]["baselines"]
        if cell in PAST_THE_TABLE:
            expected = PAST_THE_TABLE[cell]["baselines"]
        if cell in SMALL_RUNS_BASELINE:
            assert expected != SMALL_RUNS_BASELINE[cell], "unpin " + cell
            expected = SMALL_RUNS_BASELINE[cell]
        assert len(bottoms) == len(expected), (cell, bottoms, expected)
        assert bottoms == pytest.approx(expected, abs=0.5), cell
    # The small runs sit on the large run's baseline, as a cell of the large size draws them.
    for cell in ("J18", "J20"):
        bottoms = _floats(captured[cell]["word_bottoms"])
        assert bottoms == pytest.approx([runs.EXPECTED["main"][cell]["baselines"][0]] * 3, abs=0.5)


@pytest.mark.excel
def test_widths_add_per_run_as_captured():
    # The ink ends short of the last advance by the last glyph's right side bearing ("e", the
    # same in both texts, about half a point at 11 pt), on a device pixel: the larger run's
    # extra reach shows within a point. The gridlines the text crosses are hidden.
    captured = _captured()["main"]
    reach = {}
    for cell in ("C4", "C6"):
        expected = runs.EXPECTED["main"][cell]
        # ink_right is measured from I's right edge, six default columns past C's (53 pt each,
        # captured, ARCHITECTURE §4.1)
        ink = float(captured[cell]["ink_right"]) + 6 * 53
        assert expected["reach"] - 2 <= ink <= expected["reach"] + 0.5, (cell, ink)
        hidden = captured[cell]["hidden_gridlines"]
        assert hidden == "1" * expected["spill"] + "0" * (6 - expected["spill"]), (cell, hidden)
        reach[cell] = ink
    expected_extra = runs.EXPECTED["main"]["C4"]["reach"] - runs.EXPECTED["main"]["C6"]["reach"]
    assert reach["C4"] - reach["C6"] == pytest.approx(expected_extra, abs=1)


def _word(record, k):
    _, width = _spans(record["words"])[k]
    top, bottom = _floats(record["word_tops"])[k], _floats(record["word_bottoms"])[k]
    return {
        "width": width,
        "height": bottom - top,
        "bottom": bottom,
        "area": _floats(record["word_areas"])[k],
        "color": record["word_colors"].split()[k],
    }


UPSTREAM = {
    # Excel draws text under a one-section [Blue]General in blue, runs included; xlformat gives
    # text no section color there (only a text section's color), so sheetpeek keeps the runs'.
    ("main", "J28"): (("Calibri 11", "0000FF", 0.0), ("Calibri 11", "0000FF", 0.0)),
}
"""Words drawn as Excel draws them where a sibling package does not give sheetpeek the means yet:
compared at Excel's fonts and colors, and each must still differ from sheetpeek's (unpin it)."""


@pytest.mark.excel
@pytest.mark.parametrize("group", runs.GROUPS)
def test_each_word_is_drawn_in_the_font_and_color_sheetpeek_gives_it(group):
    # Each "HH" word against the reference cell holding "HH" in the font sheetpeek gives the
    # word: its ink width and height within a device pixel (half a point; glyphs are set on whole
    # points, so a word's raster is its reference's), its ink area within 12 % (bold is about half
    # again regular's), its color the nearest of the four the cases use, and a raised or lowered
    # word's bottom its shift from the first word's. Covers a run without properties (the
    # cell's font), a run with properties (nothing from the cell's), missing name and size (the
    # Normal font's), section colors, text sections and a <t> beside runs.
    found = _captured()
    if group not in found:
        pytest.skip(f"group {group} not captured yet")
    captured = found[group]
    if "refused" in captured:
        pytest.skip(f"Excel refused the file: {captured['refused']['name']}")
    references = {
        case.name.removeprefix("reference "): case.cell
        for case in runs.cases(group)
        if case.name.startswith("reference ")
    }
    for case in runs.cases(group):
        if case.name.startswith("reference ") or set(case.plain.split()) != {"HH"}:
            continue
        record = captured[case.cell]
        expected = runs.EXPECTED[group][case.cell]["words"]
        if (group, case.cell) in UPSTREAM:
            assert expected != UPSTREAM[group, case.cell], "unpin " + case.cell
            expected = UPSTREAM[group, case.cell]
        assert record["text"] == runs.EXPECTED[group][case.cell]["text"], case.name
        assert len(record["words"].split()) == len(expected), case.name
        first = _word(record, 0)
        for k, (font, color, shift) in enumerate(expected):
            word = _word(record, k)
            reference = _word(captured[references[font]], 0)
            where = (case.name, k, font)
            # Ink is counted by its distance from the background, and a colored glyph's edge
            # pixels lie nearer white than a black one's: the red 16 pt words lose half a point
            # at the left and a fifth of their area (17 pt and 74 against the black reference's
            # 18 and 93.5), with the same top and bottom. Their width is compared within a point
            # and their area not at all; black words as above.
            black = word["color"] == "000000"
            assert word["width"] == pytest.approx(reference["width"], abs=0.5 if black else 1), (
                where
            )
            assert word["height"] == pytest.approx(reference["height"], abs=0.5), where
            if black:
                assert word["area"] == pytest.approx(reference["area"], rel=0.12), where
            rgb = tuple(int(word["color"][i : i + 2], 16) for i in (0, 2, 4))
            assert nearest(rgb, PALETTE) == color, where
            assert word["bottom"] - first["bottom"] == pytest.approx(shift, abs=0.5), where
