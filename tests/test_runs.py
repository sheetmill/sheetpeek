"""Rich-text runs: read from shared strings, measured and drawn per run (ARCHITECTURE.md §4.16).

The expected layouts (zoom 100) are the suite runs' (tests/test_suite_runs.py compares them).
Excel writes runs in shared strings and openpyxl writes them inline, so the workbooks here move
their runs into shared strings, as Excel writes them.
"""

import io
import itertools
import re

import openpyxl
import pytest
from openpyxl.styles import Alignment, Font

import sheetpeek
from sheetpeek import TextSpan
from sheetpeek._model import Font as ModelFont
from sheetpeek._runs import Text, wrap
from sheetpeek.backends.html import HTMLBackend
from sheetpeek.backends.markdown import MarkdownBackend
from tests.workbooks import save, shared_strings

CALIBRI = ModelFont("Calibri", 11.0)
LONG = "a long text that spills with a LARGE run inside"


def run(text, properties=None):
    """One ``<r>``: ``properties`` is the inner XML of its ``rPr``, or ``None`` for none."""
    rpr = "" if properties is None else f"<rPr>{properties}</rPr>"
    return f'<r>{rpr}<t xml:space="preserve">{text}</t></r>'


def calibri(size=11, extra=""):
    """A complete run font, as Excel writes one."""
    return f'{extra}<sz val="{size}"/><rFont val="Calibri"/>'


def book(tmp_path, build, strings, name="book.xlsx"):
    """A workbook built by ``build``, the cells in ``strings`` turned into shared strings."""
    wb = openpyxl.Workbook()
    for coordinate in strings:
        wb.active[coordinate] = "placeholder"
    build(wb.active)
    path = save(wb, tmp_path / name)
    return shared_strings(path, strings)


def nothing(ws):
    pass


def runs_of(rendered, row):
    return next(r for r in rendered.rows if r.index == row).runs


def row_of(rendered, row):
    return next(r for r in rendered.rows if r.index == row)


# Reading.


def test_runs_are_read_with_their_fonts(tmp_path):
    def build(ws):
        ws["A1"].font = Font(name="Calibri", size=16, color="FFFF0000")

    path = book(tmp_path, build, {"A1": run("plain ") + run("bold", "<b/>") + run("!", calibri())})
    cell = sheetpeek.read(path, platform="mac").cell(1, 1)
    assert cell.value == "plain bold!"
    red = ModelFont("Calibri", 16.0, color="FF0000")
    # Captured (runs): a run without properties takes the cell's 16 pt red; a run whose
    # properties hold only bold takes nothing from the cell, and is drawn 11 pt and black. A run
    # stating no name or size takes the Normal font's.
    assert cell.runs == (
        TextSpan("plain ", red),
        TextSpan("bold", ModelFont("Calibri", 11.0, bold=True)),
        TextSpan("!", CALIBRI),
    )


def test_runs_in_the_cells_font_are_no_runs(tmp_path):
    path = book(
        tmp_path,
        nothing,
        {
            "A1": run("one ", calibri()) + run("font", calibri()),
            "A2": run("a ") + run("b"),
            "A3": run("only"),
        },
    )
    sheet = sheetpeek.read(path, platform="mac")
    assert [(c.value, c.runs) for c in (sheet.cell(r, 1) for r in (1, 2, 3))] == [
        ("one font", None),
        ("a b", None),
        ("only", None),
    ]
    assert not [n for n in sheet.notices if n.code.startswith("text.runs")]


def test_adjacent_runs_in_one_font_are_joined_and_empty_ones_dropped(tmp_path):
    strings = {"A1": run("a", "<b/>") + run("", "<i/>") + run("b", "<b/>") + run("c")}
    cell = sheetpeek.read(book(tmp_path, nothing, strings), platform="mac").cell(1, 1)
    assert [span.text for span in cell.runs] == ["ab", "c"]


def test_escapes_are_decoded_in_each_run(tmp_path):
    cell = sheetpeek.read(
        book(tmp_path, nothing, {"A1": run("_x0041_", "<b/>") + run("_x0042_")}), platform="mac"
    ).cell(1, 1)
    assert cell.value == "AB"
    assert [span.text for span in cell.runs] == ["A", "B"]


def test_one_string_shared_by_cells_of_different_fonts(tmp_path):
    def build(ws):
        ws["A2"] = "placeholder"
        ws["A2"].font = Font(size=20)

    path = book(tmp_path, build, {"A1": run("x") + run("y", "<b/>")})
    # Both cells share the first string.
    path = shared_strings_again(path)
    sheet = sheetpeek.read(path, platform="mac")
    assert sheet.cell(1, 1).runs[0].font.size == 11
    assert sheet.cell(2, 1).runs[0].font.size == 20
    assert sheet.cell(1, 1).runs[1] == sheet.cell(2, 1).runs[1]


def shared_strings_again(path):
    from tests.workbooks import SHEET1, rewrite

    return rewrite(
        path,
        SHEET1,
        lambda t: re.sub(r'<c r="A2"([^>]*) t="inlineStr"([^>]*)><is>.*?</is></c>',
                         r'<c r="A2"\1 t="s"\2><v>0</v></c>', t),
    )  # fmt: skip


def test_text_before_runs_is_a_run_in_the_cell_font(tmp_path):
    # The schema allows a <t> before runs, which Excel does not write. Excel keeps it as a run in
    # the cell's font, first, and the runs with their own properties (captured, runs).
    cell = sheetpeek.read(
        book(tmp_path, nothing, {"A1": "<t>top</t>" + run("runs", "<b/>")}), platform="mac"
    )
    found = cell.cell(1, 1)
    assert found.value == "topruns"
    assert found.runs is not None
    assert [(span.text, span.font.bold) for span in found.runs] == [("top", False), ("runs", True)]


def test_text_after_runs_is_refused_as_excel_refuses_it(tmp_path):
    path = book(tmp_path, nothing, {"A1": run("runs", "<b/>") + "<t>end</t>"})
    with pytest.raises(sheetpeek.PackageError, match="text after its runs"):
        sheetpeek.read(path, platform="mac")


def test_runs_of_inline_strings_are_drawn(tmp_path):
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont

    wb = openpyxl.Workbook()
    wb.active["B2"] = CellRichText([TextBlock(InlineFont(b=True), "rich"), " text"])
    sheet = sheetpeek.read(save(wb, tmp_path / "inline.xlsx"), platform="mac")
    cell = sheet.cell(2, 2)
    assert cell.value == "rich text"
    assert [(span.text, span.font.bold) for span in cell.runs] == [("rich", True), (" text", False)]
    assert sheet.notices == ()


# Layout in excel mode.


def runs_sheet(tmp_path):
    """The cells of the runs case, in column C; D to H are empty."""
    head, tail = LONG.split("LARGE")

    def build(ws):
        ws["C6"] = "aa BIG aa"
        ws["C6"].font = Font(name="Calibri", size=20)
        ws["C10"] = LONG
        ws["C12"].alignment = Alignment(wrap_text=True)

    strings = {
        "C4": run("aa ", calibri()) + run("BIG", calibri(20)) + run(" aa", calibri()),
        "C8": run(head, calibri()) + run("LARGE", calibri(16, "<b/>")) + run(tail, calibri()),
        "C12": run("aaa ", calibri()) + run("BBB", calibri(20)) + run(" ccc ddd", calibri()),
        "C16": run("E=mc", calibri()) + run("2", calibri(11, '<vertAlign val="superscript"/>')),
        "C18": run("H", calibri())
        + run("2", calibri(11, '<vertAlign val="subscript"/>'))
        + run("O", calibri()),
    }
    return sheetpeek.read(book(tmp_path, build, strings), platform="mac").layout("excel")


@pytest.fixture
def excel(tmp_path):
    return runs_sheet(tmp_path)


def test_a_larger_run_fits_the_row_as_a_cell_of_its_size(excel):
    # Captured (runs): one 20 pt run in an 11 pt cell gives a 26 pt row, the same as a
    # 20 pt cell, and the small runs sit on the large run's baseline.
    small, large = row_of(excel, 4), row_of(excel, 6)
    assert small.height_pt == large.height_pt == 26
    ((spanned,), (plain,)) = small.runs, large.runs
    assert spanned.y - small.y_pt == plain.y - large.y_pt
    assert [s.font.size for s in spanned.spans] == [11, 20, 11]


def test_widths_add_per_run_for_spill(excel):
    # Captured (runs): the bold 16 pt run makes the text reach 15 pt further (ink ends
    # 166 pt past C's right edge, against 151 pt), one column more; the row fits Calibri 16.
    ((spanned,), (plain,)) = runs_of(excel, 8), runs_of(excel, 10)
    assert spanned.width - plain.width == pytest.approx(15, abs=2)
    assert spanned.width == sum(s.width for s in spanned.spans)
    assert excel.cell(8, 3).spill == (0, 4)
    assert excel.cell(10, 3).spill == (0, 3)
    assert row_of(excel, 8).height_pt == 21


def test_each_wrapped_line_takes_its_largest_runs_height(excel):
    # Captured (runs): "aaa" / "BBB" / "ccc ddd", the row 59 pt = 16 + 27 + 16.
    row = row_of(excel, 12)
    assert [r.text for r in row.runs] == ["aaa", "BBB", "ccc ddd"]
    assert row.height_pt == 59
    # Lines are their largest run's line pitch apart (15 and 26 pt), each baseline its own font's
    # distance below the line's top (12 and 21 pt), as captured (runs).
    baselines = [r.y - row.y_pt for r in row.runs]
    assert [b - a for a, b in itertools.pairwise(baselines)] == [
        15 - 12 + 21,
        26 - 21 + 12,
    ]


def test_superscript_and_subscript_runs(excel):
    # Captured (runs): rows of 17 pt, as whole cells in superscript or subscript.
    assert row_of(excel, 16).height_pt == row_of(excel, 18).height_pt == 17
    (formula,) = runs_of(excel, 16)
    plain, raised = formula.spans
    assert (plain.scale, plain.shift) == (1, 0)
    assert raised.scale == pytest.approx(7 / 11)  # drawn at round(2 * 11 / 3) = 7 pt
    assert raised.shift == -5  # as a superscript cell at 11 pt (§4.13)
    (water,) = runs_of(excel, 18)
    assert [s.shift for s in water.spans] == [0, 2, 0]
    assert formula.scale == 1


def test_characters_and_none_modes_show_the_text(excel, tmp_path):
    sheet = excel.sheet
    assert "aaa BBB ccc ddd" in sheetpeek.render(sheet, "csv", on_notice="ignore").splitlines()
    assert "E=mc2" in sheetpeek.render(sheet, "text", on_notice="ignore")


def test_a_text_section_draws_in_the_cells_font(tmp_path):
    def build(ws):
        ws["A1"].number_format = '@" more"'

    rendered = sheetpeek.read(
        book(tmp_path, build, {"A1": run("bold", "<b/>") + run("x")}), platform="mac"
    ).layout("excel")
    (line,) = runs_of(rendered, 1)
    assert (line.text, line.spans) == ("boldx more", ())


def test_shrink_measures_the_runs(tmp_path):
    def build(ws):
        ws["A1"].alignment = Alignment(shrink_to_fit=True)

    big = run("HHHHH", calibri(20)) + run("HHHHH", calibri())
    rendered = sheetpeek.read(book(tmp_path, build, {"A1": big}), platform="mac").layout("excel")
    (line,) = runs_of(rendered, 1)
    assert line.scale < 1
    assert [s.scale for s in line.spans] == [line.scale, line.scale]
    assert line.width <= 53 - 3 - 2


def test_fill_repeats_the_runs(tmp_path):
    def build(ws):
        ws["A1"].alignment = Alignment(horizontal="fill")

    rendered = sheetpeek.read(
        book(tmp_path, build, {"A1": run("a", "<b/>") + run("b")}), platform="mac"
    ).layout("excel")
    (line,) = runs_of(rendered, 1)
    assert line.text == "ab" * (len(line.text) // 2)
    assert [s.text for s in line.spans] == ["a", "b"] * (len(line.text) // 2)


def test_distributed_words_keep_their_runs(tmp_path):
    def build(ws):
        ws["A1"].alignment = Alignment(horizontal="distributed")
        ws.column_dimensions["A"].width = 30

    rendered = sheetpeek.read(
        book(tmp_path, build, {"A1": run("one tw", "<b/>") + run("o three")}), platform="mac"
    ).layout("excel")
    words = runs_of(rendered, 1)
    assert [w.text for w in words] == ["one", "two", "three"]
    assert [[s.text for s in w.spans] for w in words] == [["one"], ["tw", "o"], ["three"]]
    assert words[0].spans[0].font.bold
    assert not words[2].spans[0].font.bold


def test_underlines_and_strikethrough_follow_their_runs(tmp_path):
    strings = {"A1": run("under", '<u/><color rgb="FF0000FF"/>') + run(" struck", "<strike/>")}
    rendered = sheetpeek.read(book(tmp_path, nothing, strings), platform="mac").layout("excel")
    (line,) = runs_of(rendered, 1)
    under, struck = line.rules
    first, _ = line.spans
    # 1 pt past the run's text (Excel 16.115, text_style suite)
    assert (under.kind, under.x0, under.x1) == ("underline", line.x, line.x + first.width + 1)
    assert (struck.kind, struck.x0) == ("strike", line.x + first.width)
    assert struck.x1 == pytest.approx(line.x + line.width)
    page = sheetpeek.render(rendered, "html")
    assert 'class="sp-rule" style="' in page
    assert re.search(r'<i class="sp-rule" style="[^"]*color:#0000FF"', page)


def test_a_section_color_replaces_each_runs_color(tmp_path):
    def build(ws):
        ws["A1"].number_format = "[Blue]@"

    strings = {"A1": run("red", '<color rgb="FFFF0000"/>') + run("bold", "<b/>")}
    (line,) = runs_of(
        sheetpeek.read(book(tmp_path, build, strings), platform="mac").layout("excel"), 1
    )
    assert {s.font.color for s in line.spans} == {"0000FF"}


# HTML.


def test_html_nests_a_span_per_run_with_no_line_height(excel):
    out = io.StringIO()
    HTMLBackend(fragment=True).render(excel, out)
    page = out.getvalue()
    line = re.search(
        r'<span class="sp-run" style="([^"]*)">((?:<span [^>]*>[^<]*</span>)+)</span>', page
    )
    assert line is not None
    # The positioned run takes the largest run's font, so that no run grows its line box.
    assert "font:20pt" in line[1]
    inner = re.findall(r'<span style="([^"]*)">([^<]*)</span>', line[2])
    assert [text for _, text in inner] == ["aa ", "BIG", " aa"]
    assert all(style.startswith("line-height:0") for style, _ in inner)
    assert "font:11pt" in inner[0][0]
    assert "font:" not in inner[1][0]
    assert all("letter-spacing:" in style for style, _ in inner)


def test_html_raises_a_superscript_run_by_offset(excel):
    page = sheetpeek.render(excel, HTMLBackend(fragment=True)).split("</style>")[1]
    assert "vertical-align" not in page
    raised = re.search(r'<span style="([^"]*)">2</span>', page)[1]
    assert "position:relative;top:-7px" in raised  # 5 pt
    assert "font:7pt" in raised


def test_html_escapes_each_run(tmp_path):
    rendered = sheetpeek.read(
        book(tmp_path, nothing, {"A1": run("&lt;b&gt;", "<b/>") + run("&amp;")}), platform="mac"
    ).layout("excel")
    page = sheetpeek.render(rendered, HTMLBackend(fragment=True))
    assert ">&lt;b&gt;</span><span" in page
    assert ">&amp;</span>" in page


# Other outputs.


def test_markdown_marks_bold_and_italic_runs_with_emphasis(tmp_path):
    strings = {
        "A1": run("plain ") + run("bold ", "<b/>") + run("both", "<b/><i/>") + run(" it", "<i/>"),
        "A2": run("a*b", "<b/>") + run(" c"),
    }
    sheet = sheetpeek.read(book(tmp_path, nothing, strings), platform="mac")
    rendered = sheetpeek.render(
        sheet, MarkdownBackend("letters", emphasis=True), on_notice="ignore"
    )
    assert rendered.splitlines()[2:] == [
        "| plain **bold** ***both*** *it* |",
        "| **a\\*b** c |",
    ]
    plain = sheetpeek.render(sheet, MarkdownBackend("letters"), on_notice="ignore")
    assert plain.splitlines()[2] == "| plain bold both it |"


@pytest.mark.parametrize(
    ("output", "outcome"),
    [
        ("html", None),
        ("rich", "approximated"),
        ("text", "omitted"),
        ("csv", "omitted"),
        ("tsv", "omitted"),
        ("markdown", "omitted"),
    ],
)
def test_each_output_reports_what_it_does_with_runs(tmp_path, output, outcome):
    path = book(tmp_path, nothing, {"A1": run("bold", "<b/>") + run(" plain")})
    found = {n.code: n for n in sheetpeek.notices(path, to=output)}
    if outcome is None:
        assert not found
    else:
        assert found["text.runs"].outcome == outcome
    if output in ("text", "csv", "tsv"):
        assert found["font.bold"].count == 1  # a bold run is bold text
    markdown = {n.code for n in sheetpeek.notices(path, to=MarkdownBackend(emphasis=True))}
    assert not {"text.runs", "font.bold"} & markdown


# Breaking lines by offsets.


@pytest.mark.parametrize(
    ("text", "lines"),
    [
        ("aaa bbb ccc", ["aaa", "bbb ccc"]),  # "aaa bbb " is 8: the space counts
        ("aaa bb ccc", ["aaa bb", "ccc"]),  # "aaa bb " is 7
        ("aaa bbb", ["aaa bbb"]),  # nothing follows the last word
        ("  lead", ["lead"]),
        ("a  b", ["a  b"]),
        ("x\nyy\n\nz", ["x", "yy", "", "z"]),
        ("abcdefghij", ["abcdefg", "hij"]),
        ("aaa bbbbbbbbbb", ["aaa", "bbbbbbb", "bbb"]),
        ("", [""]),
    ],
)
def test_wrap_breaks_at_spaces_and_inside_long_words(text, lines):
    # A character is 1 wide here and the budget 7.
    found = wrap(text, 7, lambda a, b: b - a)
    assert [text[a:b] for a, b in found] == lines


def test_text_measures_each_piece_in_its_font():
    import xlformat

    big = ModelFont("Calibri", 20.0)
    text = Text.of_spans(
        (TextSpan("ab\r", CALIBRI), TextSpan("\ncd", big)),
        lambda f: xlformat.Font(f.name, f.size),
        breaks=False,
    )
    assert text.text == "abcd"
    assert text.fonts(1, 3) == [CALIBRI, big]
    assert text.fonts(0, 0) == [CALIBRI]  # an empty line takes the font where it is
    assert text.fonts(4, 4) == [big]
    assert text.width(0, 4) == xlformat.text_width("ab", font=xlformat.Font("Calibri", 11)) + (
        xlformat.text_width("cd", font=xlformat.Font("Calibri", 20))
    )
