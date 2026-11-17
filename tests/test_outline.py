"""Outlines: levels and settings read, groups and symbols laid out, the gutter in each output.

The expected symbols and sizes come from captures (Excel 16 for Mac, Office 2021,
dark appearance), unverified on Microsoft 365 (ARCHITECTURE.md §4.11).
"""

import re

import openpyxl
import pytest
from openpyxl.worksheet.properties import Outline as OutlineProperties

import sheetpeek
from sheetpeek import Outline, OutlineMark
from sheetpeek._pass_outline import groups
from sheetpeek.backends import HTMLBackend, MarkdownBackend, TextBackend
from sheetpeek.backends.html._outline import _centered
from tests.workbooks import SHEET1, replace_once, rewrite, save

# The captured sheet: row -> (level, hidden, collapsed); summaries below and right.
ROWS = {
    4: (1, False, False),
    5: (2, False, False),
    6: (3, False, False),
    7: (2, False, False),
    8: (1, False, False),
    9: (2, True, False),
    10: (2, True, False),
    11: (1, False, True),
    14: (1, True, False),
    15: (1, True, False),
    16: (1, True, False),
    17: (0, False, True),
}
COLUMNS = {
    "D": (1, False, False),
    "E": (2, False, False),
    "F": (1, False, False),
    "H": (1, True, False),
    "I": (1, True, False),
    "J": (0, False, True),
}


def captured(path, *, below=True, right=True):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in range(1, 21):
        ws.cell(row, 1, row)
    ws["J1"] = "x"
    for row, (level, hidden, collapsed) in ROWS.items():
        dimension = ws.row_dimensions[row]
        dimension.outlineLevel, dimension.hidden, dimension.collapsed = level, hidden, collapsed
    for letter, (level, hidden, collapsed) in COLUMNS.items():
        dimension = ws.column_dimensions[letter]
        dimension.outlineLevel, dimension.hidden, dimension.collapsed = level, hidden, collapsed
    if not (below and right):
        ws.sheet_properties.outlinePr = OutlineProperties(summaryBelow=below, summaryRight=right)
    return save(wb, path)


def by_index(marks):
    found = {}
    for mark in marks:
        found.setdefault(mark.index, []).append((mark.lane, mark.kind))
    return found


@pytest.fixture
def book(tmp_path):
    return captured(tmp_path / "outline.xlsx")


# --- reading


def test_levels_collapsed_flags_and_settings_are_read(book):
    sheet = sheetpeek.read(book)
    rows = {r.index: (r.outline_level, r.hidden, r.collapsed) for r in sheet.rows}
    assert {i: v for i, v in rows.items() if v != (0, False, False)} == ROWS
    columns = {c.letter: (c.outline_level, c.hidden, c.collapsed) for c in sheet.columns}
    assert {k: v for k, v in columns.items() if v != (0, False, False)} == COLUMNS
    assert sheet.outline == Outline(row_depth=3, column_depth=2)


def test_the_depth_is_the_deepest_level_in_the_whole_sheet(book):
    # Captured: Excel derives the depth from the rows; openpyxl writes no outlineLevelRow.
    assert "outlineLevelRow" not in _part(book)
    sheet = sheetpeek.read(book, ref="A1:B3")
    assert sheet.outline is not None
    assert (sheet.outline.row_depth, sheet.outline.column_depth) == (3, 2)
    assert sheet.outline.rows_beyond == (0, 1)  # row 4, below the range, is grouped


def test_a_sheet_without_levels_has_no_outline(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.active.row_dimensions[1].collapsed = True  # a flag without a group
    sheet = sheetpeek.read(save(wb, tmp_path / "plain.xlsx"))
    assert sheet.outline is None
    assert sheet.rows[0].collapsed
    assert sheet.layout("excel").outline is None


def _part(path):
    import zipfile

    with zipfile.ZipFile(path) as archive:
        return archive.read(SHEET1).decode()


@pytest.mark.parametrize(
    ("view", "properties", "shown"),
    [
        (None, None, True),
        (None, "0", False),
        ("0", None, False),
        ("1", "0", True),  # the sheet view wins
        ("0", "1", False),
        ("false", None, False),
        (None, "false", False),
        ("true", "0", True),
        ("yes", "0", True),  # stated, not a boolean: xlio's default, shown (unverified)
    ],
)
def test_show_outline_symbols(book, view, properties, shown):
    if properties is not None:
        rewrite(
            book,
            SHEET1,
            replace_once("<outlinePr ", f'<outlinePr showOutlineSymbols="{properties}" '),
        )
    if view is not None:
        rewrite(
            book,
            SHEET1,
            replace_once("<sheetView ", f'<sheetView showOutlineSymbols="{view}" '),
        )
    sheet = sheetpeek.read(book)
    assert sheet.outline is not None
    assert sheet.outline.show_symbols is shown
    rendered = sheet.layout("excel")
    assert (rendered.outline is not None) is shown
    assert ("outline" in rendered.features) is shown
    assert ("outline" in {n.code for n in sheetpeek.notices(sheet, "markdown")}) is shown


@pytest.mark.parametrize(
    ("text", "level"), [("8", 0), ("9", 1), ("255", 7), ("-1", 0), ("x", 0), ("1.5", 0)]
)
def test_a_level_outside_0_to_7_is_read_as_excel_reads_it(book, text, level):
    # Modulo 8; not decimal digits, no level (xlio's capture outline). Nothing to report.
    rewrite(
        book,
        SHEET1,
        replace_once('<row r="4" outlineLevel="1"', f'<row r="4" outlineLevel="{text}"'),
    )
    sheet = sheetpeek.read(book)
    assert sheet.rows[3].outline_level == level
    assert "outline.level" not in {n.code for n in sheet.notices}


# --- groups and symbols


def test_groups_from_levels():
    levels = [0, 1, 2, 3, 2, 1, 0]
    assert groups(levels) == [(1, 1, 5, 6), (2, 2, 4, 5), (3, 3, 3, 4)]
    assert groups(levels, summary_after=False) == [(1, 1, 5, 0), (2, 2, 4, 1), (3, 3, 3, 2)]
    assert groups([]) == []


def test_the_captured_symbols(book):
    # Excel 16.114 capture: buttons on rows 7, 8, 11, 12 and 17, + on the collapsed groups' rows 11
    # and 17; a dot on each visible detail row in the lane after its level unless a button is
    # there (rows 4, 5, 6); lane 3 holds the level-3 group.
    outline = sheetpeek.read(book).layout("excel").outline
    assert outline is not None
    assert (outline.row_lanes, outline.column_lanes) == (4, 3)
    assert by_index(outline.rows) == {
        4: [(1, "start"), (2, "dot")],
        5: [(1, "bar"), (2, "start"), (3, "dot")],
        6: [(1, "bar"), (2, "bar"), (3, "start"), (4, "dot")],
        7: [(1, "bar"), (2, "bar"), (3, "open")],
        8: [(1, "bar"), (2, "open")],
        11: [(1, "bar"), (2, "closed")],
        12: [(1, "open")],
        17: [(1, "closed")],
    }
    assert by_index(outline.columns) == {
        4: [(1, "start"), (2, "dot")],
        5: [(1, "bar"), (2, "start"), (3, "dot")],
        6: [(1, "bar"), (2, "open")],
        7: [(1, "open")],
        10: [(1, "closed")],
    }


def test_summaries_above_and_left_mirror_the_symbols(tmp_path):
    # Captured for columns (the button on the left, the tick at the far end); taken as the same
    # for rows (unverified).
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in range(1, 6):
        ws.cell(row, 1, row)
    for row in (3, 4):
        ws.row_dimensions[row].outlineLevel = 1
    ws.sheet_properties.outlinePr = OutlineProperties(summaryBelow=False, summaryRight=False)
    outline = sheetpeek.read(save(wb, tmp_path / "above.xlsx")).layout("excel").outline
    assert outline is not None
    assert by_index(outline.rows) == {
        2: [(1, "open")],
        3: [(1, "bar"), (2, "dot")],
        4: [(1, "start"), (2, "dot")],
    }


def test_hidden_rows_shown_get_their_bars_and_keep_the_files_button(book):
    outline = sheetpeek.read(book).layout("excel", show_hidden=True).outline
    assert outline is not None
    marks = by_index(outline.rows)
    assert marks[9] == [(1, "bar"), (2, "start"), (3, "dot")]
    assert marks[11] == [(1, "bar"), (2, "closed")]  # the file collapses the group


def test_a_range_that_cuts_a_group_draws_its_bars_and_reports_the_button(book):
    rendered = sheetpeek.read(book, ref="A1:B8").layout("excel")
    assert rendered.outline is not None
    marks = by_index(rendered.outline.rows)
    assert marks[8] == [(1, "bar"), (2, "open")]  # rows 4-11's button is on row 12
    cut = [n for n in rendered.notices if n.code == "outline.cut"]
    # rows 9 to 11 are outside too, so which row is the summary is not known here
    assert [dict(n.context) for n in cut] == [{"axis": "rows", "first": 4, "last": 8}]
    assert (cut[0].outcome, cut[0].reason) == ("omitted", "option")
    assert "outside the range" in cut[0].message


def test_a_group_running_in_from_above_has_no_tick(book):
    # Row 7 is the summary of the level-3 group on row 6, outside the range.
    rendered = sheetpeek.read(book, ref="A7:B12").layout("excel")
    assert rendered.outline is not None
    marks = by_index(rendered.outline.rows)
    assert marks[7] == [(1, "bar"), (2, "bar"), (3, "open")]
    assert not [n for n in rendered.notices if n.code == "outline.cut"]


def test_a_cut_reports_the_summary_when_it_is_the_next_row(book):
    rendered = sheetpeek.read(book, ref="A1:B11").layout("excel")
    cut = [dict(n.context) for n in rendered.notices if n.code == "outline.cut"]
    assert cut == [{"axis": "rows", "first": 4, "last": 11, "summary": 12}]


# --- HTML


def html(path, **options):
    return sheetpeek.render(path, HTMLBackend(**options), on_notice="ignore")


def gutter_cells(text):
    body = text[text.index("<tbody>") :]
    return re.findall(r'<tr style="height:(\d+)px"><td class="sp-ol">(.*?)</td>', body)


def test_the_gutter_is_sized_by_depth(book):
    # Captured: 2 + 12 x (depth + 1) + 3 screen points, one CSS px each: 53 for 3 levels, 41
    # for 2, the same at every zoom.
    text = html(book)
    assert text.count('<col style="width:53px">') == 1
    assert '<tr style="height:41px">' in text


def test_bars_ticks_dots_and_buttons_are_placed_as_captured(book):
    cells = {i: cell for i, (_, cell) in enumerate(gutter_cells(html(book)), 1)}
    # row 4: the bar 2 px wide, 3 into lane 1 (x = 5), from 2 px below the row's top, with a
    # 6 x 2 tick; a 2 x 2 dot in lane 2 (x = 17), centered on the 20 px row
    assert '<i style="left:5px;top:2px;width:2px;height:18px"></i>' in cells[4]
    assert '<i style="left:5px;top:2px;width:6px;height:2px"></i>' in cells[4]
    assert '<i style="left:17px;top:10px;width:2px;height:2px"></i>' in cells[4]
    # row 7: the level-3 group's minus button fills lane 3 (x = 26), centered on the row, the
    # bar running into it
    assert '<i style="left:29px;top:0px;width:2px;height:5px"></i>' in cells[7]
    assert '<b class="sp-olbutton" style="left:26px;top:5px">' in cells[7]
    # row 11 (the ninth drawn): a plus with no bar in lane 2
    plus = cells[9]
    assert plus.count("<i ") == 3
    assert "left:14px;top:5px" in plus
    assert "left:17px" not in plus


@pytest.mark.parametrize(("height", "top"), [(15, 2), (30, 10), (60, 25)])
def test_a_button_is_centered_on_its_row(height, top):
    # Captured: 2 into a 15 pt row, 10 into 30, 25 into 60 (screen points).
    assert _centered(height, 12) == top


def test_level_buttons_are_drawn_beside_both_headings(book):
    text = html(book, column_headings=True, row_headings=True)
    assert text.count('class="sp-ollevel"') == 4 + 3
    assert "<thead><tr" in text
    assert text.count("<thead>") == 1
    support = HTMLBackend(column_headings=True, row_headings=True).support["outline"]
    assert support.outcome is None
    without = HTMLBackend().support["outline"]
    assert (without.outcome, without.reason) == ("approximated", "option")
    assert 'class="sp-ollevel"' not in html(book)


def test_buttons_none_leaves_the_gutter_out_and_reports_it(book):
    text = html(book, buttons="none")
    assert "sp-ol" not in text
    found = [n for n in sheetpeek.notices(book, HTMLBackend(buttons="none")) if n.code == "outline"]
    assert [(n.outcome, n.reason) for n in found] == [("omitted", "option")]


@pytest.mark.parametrize("backend", [HTMLBackend, TextBackend, sheetpeek.backends.RichBackend])
def test_an_unknown_buttons_value_is_refused(backend):
    with pytest.raises(ValueError, match="buttons"):
        backend(buttons="interactive")


def test_a_page_without_an_outline_carries_none_of_its_markup(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    text = html(save(wb, tmp_path / "plain.xlsx"), column_headings=True, row_headings=True)
    assert "sp-ol" not in text
    assert "--sheetpeek-outline" not in text


def test_the_dark_gutter_colors_are_the_captured_ones(book):
    text = html(book, color_scheme="dark")
    for color in ("#555555", "#999999", "#333333"):  # gutter 85, bars 153, button border 51
        assert color in text


# --- the other outputs


def test_text_draws_a_gutter_of_characters(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in range(1, 5):
        ws.cell(row, 1, row)
    ws["B1"], ws["C1"] = "b", "c"
    for row in (2, 3):
        ws.row_dimensions[row].outlineLevel = 1
    path = save(wb, tmp_path / "chars.xlsx")
    # openpyxl would give a grouped column a width; the XML keeps the default
    columns = '<cols><col min="2" max="2" outlineLevel="1"/></cols><sheetData>'
    rewrite(path, SHEET1, replace_once("<sheetData>", columns))
    lines = sheetpeek.render(path, "text", on_notice="ignore").splitlines()
    pad = " " * 3  # two lanes and a space
    # columns A, B (start), C (summary, its button in the middle of 8): 8 + 1 + 8 + 1 + 8
    assert lines[0] == pad + " " * 9 + "┌" + "─" * 7 + "─" + "───-"
    assert lines[1] == pad + " " * 9 + "   ·"
    assert lines[2] == pad + "       1 b        c"
    assert lines[3] == "┌·" + " " * 8 + "2"
    assert lines[4] == "│·" + " " * 8 + "3"
    assert lines[5] == "-" + " " * 9 + "4"
    plain = sheetpeek.render(path, TextBackend(buttons="none"), on_notice="ignore")
    assert plain.splitlines()[0] == "       1 b        c"


def test_rich_draws_the_same_gutter(book):
    pytest.importorskip("rich")
    from sheetpeek.backends import RichBackend

    text = sheetpeek.render(book, RichBackend(color_system=None), on_notice="ignore")
    plain = sheetpeek.render(book, "text", on_notice="ignore")
    assert [line.rstrip() for line in text.splitlines()] == plain.splitlines()


@pytest.mark.parametrize("name", ["markdown", "csv", "tsv"])
def test_outputs_without_a_gutter_report_the_outline(book, name):
    found = [n for n in sheetpeek.notices(book, name) if n.code == "outline"]
    assert [(n.outcome, n.reason) for n in found] == [("omitted", "not_representable")]
    assert "┌" not in sheetpeek.render(book, name, on_notice="ignore")


def test_the_symbols_are_laid_out_in_every_mode(book):
    marks = sheetpeek.read(book).layout("none").outline
    assert marks is not None
    assert OutlineMark(4, 1, "start") in marks.rows
    assert sheetpeek.render(book, MarkdownBackend(), on_notice="ignore").startswith("|")
