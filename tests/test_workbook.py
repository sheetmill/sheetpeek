"""The sheet list, tab colors and the workbook page (ARCHITECTURE.md §4.15).

Which tabs Excel shows and how it draws a tab color are captures (Excel 16.114.4 for
Mac, Office 2021, dark appearance, zoom 100): visible sheets only, in file order; an inactive
colored tab filled solid with white text; the active one a gradient to a light tint of its color;
``auto`` drawn as an uncolored tab.
"""

import io
import re
import warnings

import openpyxl
import pytest
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import Color

import sheetpeek
from sheetpeek import LimitError, NotAWorksheetError, ReadLimits
from sheetpeek.cli import main
from tests.workbooks import rewrite, save

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WORKBOOK = "xl/workbook.xml"
RELS = "xl/_rels/workbook.xml.rels"


def tabs_book(path):
    """The capture's sheets: four colored tabs and a plain one, a hidden and a very hidden one."""
    wb = openpyxl.Workbook()
    first = wb.active
    first.title = "Active"
    first["A1"] = "active"
    first.sheet_properties.tabColor = "FF0000"
    for name, color in (
        ("Theme", Color(theme=4)),
        ("Index", Color(indexed=12)),
        ("Auto", Color(auto=True)),
        ("Plain", None),
        ("Hidden", Color(rgb="FF00B050")),
        ("Secret", Color(rgb="FFFFC000")),
    ):
        ws = wb.create_sheet(name)
        ws["A1"] = name.lower()
        if color is not None:
            ws.sheet_properties.tabColor = color
    wb["Hidden"].sheet_state = "hidden"
    wb["Secret"].sheet_state = "veryHidden"
    return save(wb, path)


@pytest.fixture
def tabs(tmp_path):
    return tabs_book(tmp_path / "tabs.xlsx")


def kinds_book(path):
    """A worksheet, a chartsheet, a dialog sheet and a macro sheet, in that order."""
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb.active["A1"] = 1
    for name in ("Dialog", "Macro"):
        wb.create_sheet(name)["A1"] = name
    chart = BarChart()
    chart.add_data(Reference(wb.active, min_col=1, min_row=1, max_row=1))
    wb.create_chartsheet("Chart").add_chart(chart)
    save(wb, path)
    return rewrite(
        path,
        RELS,
        lambda t: t.replace(
            f'{REL}/worksheet" Target="/xl/worksheets/sheet2.xml"',
            f'{REL}/dialogsheet" Target="/xl/worksheets/sheet2.xml"',
        ).replace(
            f'{REL}/worksheet" Target="/xl/worksheets/sheet3.xml"',
            'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" '
            'Target="/xl/worksheets/sheet3.xml"',
        ),
    )


@pytest.fixture
def kinds(tmp_path):
    return kinds_book(tmp_path / "kinds.xlsx")


def active_tab(path, index):
    """Make sheet ``index`` (from 0) the active one."""
    return rewrite(path, WORKBOOK, lambda t: re.sub(r'activeTab="\d+"', f'activeTab="{index}"', t))


def page(source, **options):
    options.setdefault("on_notice", "ignore")
    result = sheetpeek.render_workbook(source, **options)
    assert result is not None
    return result


def nav(markup):
    return re.search(r"<nav>(.*?)</nav>", markup, re.S)[1]


# --- the sheet list


def test_sheets_list_kind_state_and_tab_color(tabs):
    with sheetpeek.open(tabs) as book:
        found = [(s.name, s.kind, s.state, s.tab_color) for s in book.sheets]
    # Theme accent1 of the default theme is 4F81BD and indexed 12 is 0000FF, as the capture read
    # them on the tabs; auto draws an uncolored tab (capture on Excel 16.114).
    assert found == [
        ("Active", "worksheet", "visible", "FF0000"),
        ("Theme", "worksheet", "visible", "4F81BD"),
        ("Index", "worksheet", "visible", "0000FF"),
        ("Auto", "worksheet", "visible", None),
        ("Plain", "worksheet", "visible", None),
        ("Hidden", "worksheet", "hidden", "00B050"),
        ("Secret", "worksheet", "very_hidden", "FFC000"),
    ]


def test_sheets_tell_kinds_from_the_relationship_type(kinds):
    with sheetpeek.open(kinds) as book:
        assert [(s.name, s.kind, s.readable) for s in book.sheets] == [
            ("Data", "worksheet", True),
            ("Dialog", "dialogsheet", False),
            ("Macro", "macrosheet", True),
            ("Chart", "chartsheet", False),
        ]
        assert [s.name for s in book.sheets] == book.sheet_names


@pytest.mark.parametrize(("name", "words"), [("Chart", "chartsheet"), ("Dialog", "dialog sheet")])
def test_a_chartsheet_or_dialog_sheet_is_not_a_worksheet(kinds, name, words):
    with sheetpeek.open(kinds) as book, pytest.raises(NotAWorksheetError, match=words) as caught:
        book.sheet(name)
    assert isinstance(caught.value, ValueError)


def test_a_macro_sheet_renders_as_a_worksheet_with_its_notice(kinds):
    sheet = sheetpeek.read(kinds, "Macro")
    assert sheet.cell(1, 1).value == "Macro"
    assert [n.code for n in sheet.notices] == ["sheet.macrosheet"]


def test_the_default_sheet_skips_an_active_chartsheet(kinds):
    assert sheetpeek.read(active_tab(kinds, 3)).name == "Data"
    assert sheetpeek.read(active_tab(kinds, 2)).name == "Macro"
    assert sheetpeek.read(active_tab(kinds, 1)).name == "Data"  # a dialog sheet


def test_the_default_sheet_skips_a_hidden_active_sheet(tabs):
    assert sheetpeek.read(active_tab(tabs, 6)).name == "Active"
    assert sheetpeek.read(active_tab(tabs, 2)).name == "Index"


# openpyxl, which still reads the cells, warns about the sheet without a part.
@pytest.mark.filterwarnings("ignore::UserWarning:openpyxl")
def test_a_sheet_listed_without_a_part_still_counts_in_tab_and_name_indexes(tmp_path):
    # <sheet r:id=""> is listed in the workbook part but has no part: activeTab and a defined
    # name's localSheetId count it as written, so the sheets after it keep their own.
    workbook = openpyxl.Workbook()
    workbook.active.title = "First"
    for name, area in (("Second", "A1:B2"), ("Third", "C3:D4")):
        sheet = workbook.create_sheet(name)
        sheet["A1"] = name
        sheet.print_area = area
    path = save(workbook, tmp_path / "gone.xlsx")

    def edit(text):
        text = re.sub(
            r"(<sheets><sheet [^>]*/>)",
            r'\1<sheet name="Gone" sheetId="9" '
            r'state="veryHidden" r:id=""/>',
            text,
            count=1,
        )
        text = text.replace('localSheetId="2"', 'localSheetId="3"')
        return text.replace('localSheetId="1"', 'localSheetId="2"')

    rewrite(path, WORKBOOK, edit)
    active_tab(path, 2)
    assert sheetpeek.read(path).name == "Second"
    assert sheetpeek.read(path, "Second").print_areas == (sheetpeek.CellRange(1, 1, 2, 2),)
    assert sheetpeek.read(path, "Third").print_areas == (sheetpeek.CellRange(3, 3, 4, 4),)


def test_a_workbook_of_chartsheets_has_no_default_sheet(kinds):
    # with lxml installed, openpyxl writes the r: namespace before the name
    hide = r'(<sheet [^>]*?name="(?:Data|Macro)" sheetId="\d+") state="visible"'
    rewrite(kinds, WORKBOOK, lambda t: re.sub(hide, r'\1 state="hidden"', t))
    assert sheetpeek.read(kinds).name == "Data"  # hidden, but the only readable kind left
    rewrite(
        kinds,
        RELS,
        lambda t: t.replace(
            f'{REL}/worksheet" Target="/xl/worksheets/sheet1.xml"',
            f'{REL}/dialogsheet" Target="/xl/worksheets/sheet1.xml"',
        ).replace("xlMacrosheet", "dialogsheet"),
    )
    with pytest.raises(NotAWorksheetError, match="no worksheet"):
        sheetpeek.read(kinds)


def test_very_hidden_sheets_render_by_name(tabs):
    sheet = sheetpeek.read(tabs, "Secret")
    assert (sheet.state, sheet.cell(1, 1).value) == ("very_hidden", "secret")


def test_a_one_sheet_render_reports_the_tab_color_as_an_option(tabs):
    (notice,) = [n for n in sheetpeek.read(tabs).notices if n.code == "sheet.tab_color"]
    assert (notice.outcome, notice.reason) == ("omitted", "option")
    assert "render_workbook" in notice.message
    assert not [n for n in sheetpeek.read(tabs, "Auto").notices if n.code == "sheet.tab_color"]


# --- the workbook page


def test_the_page_has_a_tab_for_each_visible_sheet_in_file_order(tabs):
    markup = page(tabs)
    labels = re.findall(r">([^<]+)</a>", nav(markup))
    assert labels == ["Active", "Theme", "Index", "Auto", "Plain"]  # capture on Excel 16.114
    assert re.findall(r'<section id="(sp-\d+)"', markup) == [f"sp-{i}" for i in range(1, 6)]
    assert ">hidden<" not in markup
    assert ">secret<" not in markup


def test_hidden_sheets_are_drawn_on_request(tabs):
    markup = page(tabs, hidden_sheets=True)
    assert re.findall(r">([^<]+)</a>", nav(markup))[5:] == ["Hidden", "Secret"]
    assert '<a href="#sp-6" class="sp-hidden sp-t-f" title="hidden sheet">' in markup
    assert '<a href="#sp-7" class="sp-hidden sp-t-g" title="very hidden sheet">' in markup


def test_the_tab_strip_is_below_the_sheets_or_above_on_request(tabs):
    markup = page(tabs)
    assert markup.index("<nav>") > markup.rindex("</section>")  # as Excel draws it
    markup = page(tabs, tabs="top")
    assert markup.index("</nav>") < markup.index("<section")
    with pytest.raises(ValueError, match="tabs must be"):
        page(tabs, tabs="left")


def test_hidden_sheets_and_hidden_rows_are_separate_options(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"], wb.active["A2"] = "shown", "row hidden"
    wb.active.row_dimensions[2].hidden = True
    wb.create_sheet("Hidden").sheet_state = "hidden"
    path = tmp_path / "both.xlsx"
    wb.save(path)
    markup = page(path, hidden_sheets=True)
    assert ">Hidden</a>" in markup
    assert "row hidden" not in markup
    markup = page(path, show_hidden=True)
    assert ">Hidden</a>" not in markup
    assert "row hidden" in markup


def test_tab_colors_as_excel_draws_them(tabs):
    markup = page(tabs)
    # Inactive: filled solid with white text. Active: a gradient from near white to a light
    # tint, with Excel's green text. Auto: an uncolored tab (capture on Excel 16.114).
    assert ".sp-book > nav a.sp-t-b { background: #4F81BD; color: #ffffff; }" in markup
    assert ".sp-book > nav a.sp-t-c { background: #0000FF; color: #ffffff; }" in markup
    assert "sp-t-d" not in markup  # auto
    assert "sp-t-e" not in markup  # none
    active = "background: linear-gradient(#FFE6E6, #FF9999); color: #278452; font-weight: bold;"
    assert f".sp-book > nav a.sp-first {{ {active} }}" in markup
    assert '<a href="#sp-1" class="sp-first sp-t-a">Active</a>' in markup


def test_the_active_tab_is_the_first_shown(tabs):
    markup = page(active_tab(tabs, 2))
    assert '<a href="#sp-3" class="sp-first sp-t-c">Index</a>' in markup
    assert '<section id="sp-3" class="sp-first">' in markup
    markup = page(active_tab(tabs, 5))  # hidden: the first tab instead
    assert '<section id="sp-1" class="sp-first">' in markup


def test_chartsheets_and_dialog_sheets_get_a_tab_and_a_panel(kinds):
    markup = page(kinds)
    assert re.findall(r">([^<]+)</a>", nav(markup)) == ["Data", "Dialog", "Macro", "Chart"]
    assert "&#x27;Chart&#x27; is a chartsheet: charts are not shown." in markup
    assert "&#x27;Dialog&#x27; is a dialog sheet, a legacy form: it is not shown." in markup
    assert markup.count('<div class="sheetpeek') == 2


def linked_book(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "From"
    for cell, location in (
        ("A1", "'To sheet'!B2"),
        ("A2", "to SHEET!$C$3"),  # names compare ignoring case, as Excel's do
        ("A3", "From!A1"),
        ("A4", "Hidden!A1"),
        ("A5", "'To sheet'!Z99"),  # outside what is drawn
        ("A6", "Total"),  # a defined name
    ):
        ws[cell] = "link"
        ws[cell].hyperlink = f"#{location}"
        ws[cell].hyperlink.target = None
        ws[cell].hyperlink.location = location
    other = wb.create_sheet("To sheet")
    other["B2"], other["C3"] = "target", "merged"
    other.merge_cells("C3:D4")
    other["A1"] = "back"
    other["A1"].hyperlink = "#From!A2"
    other["A1"].hyperlink.target = None
    other["A1"].hyperlink.location = "From!A2"
    wb.create_sheet("Hidden").sheet_state = "hidden"
    return save(wb, path)


def test_links_to_another_sheet_on_the_page_go_there(tmp_path):
    markup = page(linked_book(tmp_path / "links.xlsx"))
    hrefs = re.findall(r'<a href="(#sp-\d+-[^"]+)">', markup)
    assert hrefs == ["#sp-2-B2", "#sp-2-C3", "#sp-1-A1", "#sp-1-A2"]  # the last one back
    for target in ("sp-2-B2", "sp-2-C3", "sp-1-A1", "sp-1-A2"):
        assert markup.count(f'id="{target}"') == 1, target
    # The hidden sheet, a cell not drawn and a defined name: anchors that go nowhere.
    assert markup.count("<a>") == 3


def test_a_one_sheet_render_keeps_its_links_and_ids(tmp_path):
    markup = sheetpeek.render(linked_book(tmp_path / "links.xlsx"), "html", on_notice="ignore")
    assert re.findall(r'<a href="(#[^"]+)">', markup) == ["#sp-A1"]
    assert markup.count("<a>") == 5


def test_links_followed_on_the_page_are_not_reported(tmp_path):
    found = []
    page(linked_book(tmp_path / "links.xlsx"), on_notice=found.append)
    by_code = {(n.sheet, n.code): n for n in found}
    assert by_code[("From", "link.elsewhere")].count == 3
    assert ("To sheet", "link.elsewhere") not in by_code
    one = sheetpeek.notices(linked_book(tmp_path / "again.xlsx"), "html", sheet="To sheet")
    assert [n.count for n in one if n.code == "link.elsewhere"] == [1]


def test_notices_are_delivered_per_sheet_and_for_the_workbook(tabs):
    found = []
    page(tabs, on_notice=found.append)
    assert "sheet.tab_color" not in {n.code for n in found}  # drawn
    (hidden,) = [n for n in found if n.code == "sheet.hidden"]
    assert (hidden.sheet, hidden.count, hidden.reason) == (None, 2, "option")
    assert hidden.context["very_hidden"] == 1
    assert "1 of them very hidden" in hidden.message
    found = []
    page(tabs, hidden_sheets=True, on_notice=found.append)
    assert found == []


def test_the_default_is_one_warning_per_sheet_and_one_for_the_workbook(kinds):
    with pytest.warns(sheetpeek.NotShownWarning) as caught:
        sheetpeek.render_workbook(kinds)
    messages = [str(w.message) for w in caught if w.category is sheetpeek.NotShownWarning]
    assert len(messages) == 2
    assert messages[0].startswith("sheet 'Macro' to html")
    assert messages[1].startswith("workbook to html: 2 kind(s)")
    assert all(w.filename == __file__ for w in caught)


def test_the_page_lists_its_notices_on_request(kinds):
    markup = page(kinds, notices=True)
    assert markup.count('<details class="sp-notices"') == 2
    assert "chartsheets not shown: &#x27;Chart&#x27;" in markup


def fills_book(path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "x"
    wb.active["A1"].fill = PatternFill("gray125", fgColor="FF0000")
    other = wb.create_sheet("Two")
    other["A1"] = "y"
    other["A1"].fill = PatternFill("darkGrid", fgColor="0000FF")
    return save(wb, path)


def test_generated_classes_are_unique_to_each_sheet(tmp_path):
    markup = page(fills_book(tmp_path / "fills.xlsx"))
    assert re.findall(r'<i class="(sp-f-[a-z-]+)"', markup) == ["sp-f-a-a", "sp-f-b-a"]
    assert markup.count(".sheetpeek .sp-f-a-a {") == markup.count(".sheetpeek .sp-f-b-a {") == 1
    single = sheetpeek.render(fills_book(tmp_path / "one.xlsx"), "html", on_notice="ignore")
    assert re.findall(r'<i class="(sp-f-[a-z-]+)"', single) == ["sp-f-a"]


def test_the_page_bounds_its_cells_together(tabs):
    # Five sheets of one cell each: four pass a limit of four per sheet, the page does not.
    page(tabs, limits=ReadLimits(max_cells=5))
    with pytest.raises(LimitError, match="workbook page is over 4 cells by sheet 'Plain'"):
        page(tabs, limits=ReadLimits(max_cells=4))


def test_a_fragment_and_options(tabs):
    markup = page(tabs, fragment=True, column_headings=True, color_scheme="dark")
    assert markup.startswith("<style>\n")
    assert '<div class="sp-book">' in markup
    assert "<body>" not in markup
    assert 'data-color-scheme="dark"' in markup
    assert ".sp-book > nav a { background: #3a3a3a; color: #bbbbbb; }" in markup
    with pytest.raises(TypeError):
        sheetpeek.render_workbook(tabs, bogus=True)


def test_an_open_book_is_left_open_and_out_is_written(tabs):
    out = io.StringIO()
    with sheetpeek.open(tabs) as book:
        assert sheetpeek.render_workbook(book, out=out, on_notice="ignore") is None
        assert book.sheet("Plain").cell(1, 1).value == "plain"
    assert "<title>tabs.xlsx</title>" in out.getvalue()
    with tabs.open("rb") as handle:
        assert "<title>workbook</title>" in page(io.BytesIO(handle.read()))


# --- the command line


def test_cli_all_sheets(tabs, capsys):
    assert main(["view", str(tabs), "--all-sheets", "-t", "html", "--notices", "none"]) == 0
    assert '<div class="sp-book">' in capsys.readouterr().out
    assert main(["view", str(tabs), "--all-sheets", "--hidden", "--notices", "none"]) == 0
    assert ">Secret</a>" not in capsys.readouterr().out  # --hidden is rows and columns
    assert main(["view", str(tabs), "--all-sheets", "--hidden-sheets", "--notices", "none"]) == 0
    assert ">Secret</a>" in capsys.readouterr().out


def test_cli_hidden_sheets_goes_with_all_sheets(tabs, capsys):
    assert main(["view", str(tabs), "--hidden-sheets"]) == 2
    assert "--hidden-sheets goes with --all-sheets" in capsys.readouterr().err


def test_cli_all_sheets_reports_and_is_strict(tabs, capsys):
    assert main(["view", str(tabs), "--all-sheets", "--strict"]) == 3
    assert "sheetpeek: sheet.hidden: hidden sheets not shown (2" in capsys.readouterr().err


@pytest.mark.parametrize(
    ("extra", "words"),
    [
        (["-t", "text"], "-t text has no workbook page"),
        (["-t", "csv"], "-t csv has no workbook page"),
        (["--sheet", "Plain"], "no --sheet or --ref"),
        (["--ref", "A1"], "no --sheet or --ref"),
        (["--list"], "--list cannot go together"),
    ],
)
def test_cli_all_sheets_refuses_what_it_cannot_do(tabs, capsys, extra, words):
    assert main(["view", str(tabs), "--all-sheets", *extra]) == 2
    captured = capsys.readouterr()
    assert captured.out == ""
    assert words in captured.err


def test_cli_list_shows_state_and_kind(tabs, capsys):
    assert main(["view", str(tabs), "--list"]) == 0
    assert capsys.readouterr().out.splitlines()[4:] == [
        "Plain",
        "Hidden  (hidden)",
        "Secret  (very hidden)",
    ]


def test_render_still_warns_once(tabs):
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        sheetpeek.render(tabs, "html")
    assert len(caught) == 1


@pytest.mark.parametrize(
    "color",
    ["FF'><b>x", "FF;x:url", "#FF0000", "FF0000;}", "red", "FFGGGGGG", "", "FF00FF0"],
)
def test_a_tab_color_reaches_css_only_as_hex(kinds, color):
    # Excel takes text that is not hex as opaque black, and fewer than eight hex digits as a
    # number whose low 24 bits are the color (xlio's capture attributes).
    from xml.sax.saxutils import quoteattr

    tab = f"<sheetPr><tabColor rgb={quoteattr(color)} />"
    rewrite(kinds, "xl/worksheets/sheet2.xml", lambda t: t.replace("<sheetPr>", tab, 1))
    markup = page(kinds)
    with sheetpeek.open(kinds) as book:
        expected = (
            f"{int(color, 16) & 0xFFFFFF:06X}"
            if re.fullmatch("[0-9A-Fa-f]{1,7}", color)
            else "000000"
        )
        assert book.sheets[1].tab_color == expected
    assert "<b>" not in markup
    assert "GGGG" not in markup
    assert "x:url" not in markup
    assert f"nav a.sp-t-b {{ background: #{expected};" in markup


def test_a_dialog_sheets_tab_color_is_drawn(kinds):
    rewrite(
        kinds,
        "xl/worksheets/sheet2.xml",
        lambda t: t.replace("<sheetPr>", '<sheetPr><tabColor rgb="FF00FF00" />', 1),
    )
    assert ".sp-book > nav a.sp-t-b { background: #00FF00; color: #ffffff; }" in page(kinds)


@pytest.mark.parametrize(
    ("tint", "expected"),
    [
        ("nan", "4F81BD"),
        ("INF", "4F81BD"),
        ("-inf", "4F81BD"),
        ("x", "4F81BD"),
        # Past 1 or -1, a tint wraps as a signed 16-bit count of 32,767ths, as Excel takes it
        # (xlio's capture attributes): 2 is -2/32,767, which leaves the color as it is.
        ("2", "4F81BD"),
        ("-2", "4F81BD"),
        ("-0.5", "244062"),
    ],
)
def test_a_tab_tint_is_read_as_excel_reads_it(kinds, tint, expected):
    tab = f'<sheetPr><tabColor theme="4" tint="{tint}" />'
    rewrite(kinds, "xl/worksheets/sheet2.xml", lambda t: t.replace("<sheetPr>", tab, 1))
    with sheetpeek.open(kinds) as book:
        assert book.sheets[1].tab_color == expected


def test_a_tab_color_is_read_only_from_sheet_properties(kinds):
    # sheetPr comes first; a tabColor anywhere else is not the tab's, and the read stops there.
    misplaced = '<customSheetViews><tabColor rgb="FF00FF00" /></customSheetViews><sheetPr>'
    rewrite(kinds, "xl/worksheets/sheet2.xml", lambda t: t.replace("<sheetPr>", misplaced, 1))
    with sheetpeek.open(kinds) as book:
        assert book.sheets[1].tab_color is None
