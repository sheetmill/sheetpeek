"""Notices: nothing a workbook uses is dropped silently, and nothing it merely contains is reported.

The every-feature workbook holds one use of every registered feature; each output must show each
feature it is given or report it. Plain workbooks from openpyxl, XlsxWriter and Excel must give
no notice at all.
"""

import re
import warnings
from pathlib import Path

import openpyxl
import pytest
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, GradientFill, PatternFill, Side
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.views import Pane

import sheetpeek
from sheetpeek import cli
from sheetpeek._notices import FEATURES, CellSet, visible
from sheetpeek.backends import _support, backend_names, get_backend
from tests.workbooks import replace_once, replace_part, rewrite, save

FILES = Path(__file__).parent / "files"
SHEET1 = "xl/worksheets/sheet1.xml"
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CF_URI = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Registered codes the every-feature workbook cannot hold, each tested where it can arise.
ELSEWHERE = {
    "print.area_unusable": "only with render(print_area=True); tests/test_view_drawn.py",
    "format.refused": "tests/test_layout.py and tests/test_cli.py",
    "value.unshowable": "tests/test_layout.py",
    "outline.cut": "only with a range that cuts a group; tests/test_outline.py",
    "font.cut": "only past 999 fonts; tests/test_font_cut.py",
    "grid.color": "only where gridlines are shown, and All hides them; tests/test_theme.py",
    "sheet.hidden": "only on the workbook page, which leaves hidden sheets out; "
    "tests/test_workbook.py",
}

DRAWING = (
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
    f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" {R}>'
    "<xdr:twoCellAnchor><xdr:from/><xdr:to/>"
    '<xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic>'
    "<xdr:clientData/></xdr:twoCellAnchor>"
    "<xdr:twoCellAnchor><xdr:from/><xdr:to/>"
    '<xdr:pic><xdr:blipFill><a:blip r:link="rId2"/></xdr:blipFill></xdr:pic>'
    "<xdr:clientData/></xdr:twoCellAnchor>"
    "<xdr:oneCellAnchor><xdr:from/><xdr:ext/><xdr:graphicFrame><a:graphic>"
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>'
    "</a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>"
    "<xdr:oneCellAnchor><xdr:from/><xdr:ext/><xdr:graphicFrame><a:graphic>"
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>'
    "</a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>"
    "<xdr:absoluteAnchor><xdr:pos/><xdr:ext/><xdr:sp/><xdr:clientData/></xdr:absoluteAnchor>"
    "</xdr:wsDr>"
)
THREADED = (
    '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/'
    'threadedcomments"><threadedComment ref="G2" id="{1}"/></ThreadedComments>'
)
STRINGS = (
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" '
    'uniqueCount="2"><si><t>kanji</t><rPh sb="0" eb="2"><t>kana</t></rPh></si>'
    "<si><r><rPr><b/></rPr><t>bold</t></r><r><t> plain</t></r></si></sst>"
)


def extensions(*uris: str) -> str:
    body = "".join(
        f'<ext uri="{uri}"><x14:conditionalFormattings xmlns:x14="urn:x14">'
        "<x14:conditionalFormatting><x14:cfRule/></x14:conditionalFormatting>"
        "</x14:conditionalFormattings></ext>"
        if uri.upper() == CF_URI.upper()
        else f'<ext uri="{uri}"/>'
        for uri in uris
    )
    return f"<extLst>{body}</extLst>"


def build_every_feature(path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "All"
    # What the model carries (the census).
    ws["A1"] = "names"
    ws["A1"].font = Font(
        name="Segoe UI",
        size=14,
        bold=True,
        italic=True,
        underline="single",
        strike=True,
        color="FF0000",
    )
    ws["A2"] = "accounting"
    ws["A2"].font = Font(underline="singleAccounting")
    ws["A3"] = -1
    ws["A3"].number_format = "0;[Red]-0"
    ws["A4"].fill = PatternFill("solid", fgColor="FFFF00")
    ws["A5"] = "aligned"
    ws["A5"].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True, indent=1)
    ws["A6"] = "link"
    ws["A6"].hyperlink = "https://example.com/"
    ws["A7"] = "here"
    ws["A7"].hyperlink = Hyperlink(ref="A7", location="'All'!C1")
    ws["A8"] = "elsewhere"
    ws["A8"].hyperlink = Hyperlink(ref="A8", location="Dialog!A1")
    # What the reader reports of links.
    ws["A9"] = "refused"
    ws["A9"].hyperlink = "javascript:alert(1)"
    ws["A15"] = "local"
    ws["A15"].hyperlink = "../budget.xlsx"
    ws["A11"] = '=HYPERLINK(B1&"x")'
    ws["A12"] = '=HYPERLINK("https://example.org/")'
    ws["A12"].hyperlink = "https://example.com/"
    ws["C1"] = "merged"
    ws.merge_cells("C1:D1")
    ws.column_dimensions["B"].width = 20
    ws.row_dimensions[3].height = 30
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = False
    # What the reader reports.
    ws["B2"] = "border"
    ws["B2"].border = Border(left=Side("thin"))
    ws["B3"] = "diagonal"
    ws["B3"].border = Border(diagonal=Side("thin"), diagonalUp=True)
    ws["B4"].fill = PatternFill("gray125", fgColor="000000")
    ws["B5"].fill = GradientFill(stop=("FFFFFF", "0000FF"))
    ws["M2"].border = Border(bottom=Side("slantDashDot"))
    ws["B6"] = "superscript"
    ws["B6"].font = Font(vertAlign="superscript")
    ws["B7"] = "scheme"
    ws["B7"].font = Font(name="Verdana", scheme="minor")
    for row, alignment in enumerate(
        (
            Alignment(textRotation=45, wrap_text=True),  # drawn in lines
            Alignment(textRotation=255, wrap_text=True),  # stacked: drawn on one line
            Alignment(shrink_to_fit=True),
            Alignment(readingOrder=2),
            Alignment(vertical="justify"),
        ),
        8,
    ):
        ws.cell(row, 2, "text").alignment = alignment
    ws["B8"].border = Border(top=Side("thin"))  # Excel slants it with the text, as drawn
    ws["B16"] = "merged"  # a merge's borders stay upright
    ws["B16"].alignment = Alignment(textRotation=45)
    ws["B16"].border = Border(left=Side("thin"))
    ws.merge_cells("B16:C16")
    ws["B13"] = CellRichText([TextBlock(InlineFont(b=True), "rich"), " text"])
    ws["B14"] = '=CUBEVALUE("x")'  # not computed, no cached value
    # A [ColorN] the workbook palette recolors (index 17, [Color10]).
    wb._colors = [*wb._colors[:17], "00804000", *wb._colors[18:]]
    ws["O2"] = 1
    ws["O2"].number_format = "[Color10]0"
    ws["A30"] = "row format"
    ws.row_dimensions[30].fill = PatternFill("solid", fgColor="00FF00")
    ws.sheet_view.showZeros = False
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showFormulas = True
    ws.sheet_view.view = "pageLayout"
    ws.sheet_view.defaultGridColor = False
    ws.sheet_view.colorId = 10
    ws.sheet_view.pane = Pane(xSplit=2400, ySplit=1800, topLeftCell="C4", state="split")
    ws.row_dimensions[5].outlineLevel = 1
    ws.column_dimensions["Z"].outlineLevel = 9  # past Excel's 7
    ws.conditional_formatting.add(
        "A1:A3", CellIsRule(operator="equal", formula=["1"], fill=PatternFill("solid", "FF0000"))
    )
    ws.auto_filter.ref = "A1:B14"
    ws["E1"], ws["F1"], ws["E2"], ws["F2"] = "h1", "h2", 1, 2
    ws.add_table(Table(displayName="T1", ref="E1:F2"))
    ws["G1"] = "noted"
    ws["G1"].comment = Comment("a note", "author")
    ws["G2"] = "threaded"
    ws.print_area = "A1:B2"
    ws.print_title_rows = "1:1"
    ws.row_breaks.append(Break(id=5))
    ws.sheet_properties.tabColor = "FF0000"
    for name in ("Dialog", "Macro"):
        wb.create_sheet(name)["A1"] = "x"
    wb.create_chartsheet("Chart")
    save(wb, path)

    rewrite(
        path,
        SHEET1,
        replace_once(
            "</worksheet>",
            f'<drawing {R} r:id="rIdDrw" /><picture {R} r:id="rIdPic" />'
            f'<oleObjects><oleObject {R} progId="Package" shapeId="1" r:id="rIdOle" /></oleObjects>'
            f'<controls><control {R} shapeId="2" r:id="rIdCtl" name="c" /></controls>'
            "<bogusElement />"
            + extensions(
                CF_URI,
                "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}",
                "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}",
                "{7E03D99C-DC04-49D9-9315-930204A7B6E9}",
                "{00000000-0000-0000-0000-000000000000}",
            )
            + "</worksheet>",
        ),
    )
    rewrite(
        path,
        SHEET1,
        lambda t: re.sub(
            r'(<row r="1"[^>]*>.*?)</row>',
            r'\1<c r="H1" t="s" ph="1"><v>0</v></c><c r="I1" t="s"><v>1</v></c></row>',
            t,
            count=1,
            flags=re.S,
        ),
    )
    # The two HYPERLINK() cells with cached values, as Excel saves them.
    rewrite(
        path,
        SHEET1,
        lambda t: re.sub(
            r'<c r="(A1[12])"[^>]*>(<f>.*?</f>).*?</c>',
            r'<c r="\1" t="str">\2<v>shown</v></c>',
            t,
        ).replace('<hyperlink ref="A7"', '<hyperlink ref="A10" location="A1"/><hyperlink ref="A7"'),
    )  # and a link on the empty A10, which openpyxl does not write
    rels = "xl/worksheets/_rels/sheet1.xml.rels"
    rewrite(
        path,
        rels,
        replace_once(
            "</Relationships>",
            f'<Relationship Id="rIdDrw" Type="{REL}/drawing" Target="../drawings/drawing9.xml" />'
            f'<Relationship Id="rIdPic" Type="{REL}/image" Target="../media/background.png" />'
            f'<Relationship Id="rIdThr" Type="http://schemas.microsoft.com/office/2017/10/'
            'relationships/threadedComment" Target="../threadedComments/thread1.xml" />'
            "</Relationships>",
        ),
    )
    replace_part(path, "xl/drawings/drawing9.xml", DRAWING.encode())
    replace_part(path, "xl/media/background.png", b"\x89PNG\r\n\x1a\n")
    # Excel gives a chartsheet a drawing; openpyxl cannot read one without its relationships.
    replace_part(
        path,
        "xl/chartsheets/_rels/sheet1.xml.rels",
        (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL}/drawing" Target="../drawings/drawing8.xml"/>'
            "</Relationships>"
        ).encode(),
    )
    replace_part(
        path,
        "xl/drawings/drawing8.xml",
        b'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
    )
    replace_part(path, "xl/threadedComments/thread1.xml", THREADED.encode())
    replace_part(path, "xl/sharedStrings.xml", STRINGS.encode())
    rewrite(
        path,
        "[Content_Types].xml",
        replace_once(
            "</Types>",
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" /></Types>',
        ),
    )
    rewrite(
        path,
        "xl/_rels/workbook.xml.rels",
        lambda t: (
            t.replace(
                "</Relationships>",
                f'<Relationship Id="rIdSst" Type="{REL}/sharedStrings" '
                'Target="sharedStrings.xml" /></Relationships>',
            )
            .replace(
                f'{REL}/worksheet" Target="/xl/worksheets/sheet2.xml"',
                f'{REL}/dialogsheet" Target="/xl/worksheets/sheet2.xml"',
            )
            .replace(
                f'{REL}/worksheet" Target="/xl/worksheets/sheet3.xml"',
                'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" '
                'Target="/xl/worksheets/sheet3.xml"',
            )
        ),
    )
    vml = next(n for n in __import__("zipfile").ZipFile(path).namelist() if n.endswith(".vml"))
    rewrite(
        path,
        vml,
        replace_once(
            "</xml>",
            '<v:shape><x:ClientData ObjectType="Checkbox"></x:ClientData></v:shape></xml>',
        ),
    )
    return path


@pytest.fixture(scope="module")
def every(tmp_path_factory):
    return build_every_feature(tmp_path_factory.mktemp("notices") / "every.xlsx")


def codes(notices):
    return {n.code for n in notices}


def test_every_reader_code_is_reported(every):
    with sheetpeek.open(every, platform="mac") as book:
        found = codes(book.notices)
        for name in ("All", "Macro"):
            found |= codes(book.sheet(name).notices)
    expected = {c for c, f in FEATURES.items() if f.source == "reader"} - set(ELSEWHERE)
    assert expected - found == set()


def test_every_layout_code_is_reported(every):
    sheet = sheetpeek.read(every, "All", platform="mac")
    found = codes(sheet.layout("excel").notices) | codes(
        sheet.layout("characters", show_hidden=True).notices
    )
    hidden = _with_hidden(every.parent / "hidden.xlsx")
    found |= codes(sheetpeek.read(hidden, platform="mac").layout("none", show_hidden=True).notices)
    expected = {c for c, f in FEATURES.items() if f.source == "layout"} - set(ELSEWHERE)
    assert expected - found == set()


def _with_hidden(path):
    wb = openpyxl.Workbook()
    wb.active["A1"], wb.active["B2"] = 1, 2
    wb.active.row_dimensions[1].hidden = True
    wb.active.column_dimensions["A"].hidden = True
    return save(wb, path)


def test_every_census_code_is_present(every):
    found = set(sheetpeek.read(every, "All", platform="mac").layout("excel").features)
    expected = {c for c, f in FEATURES.items() if f.source == "census"} - set(ELSEWHERE)
    assert expected - found == set()


@pytest.mark.parametrize("name", backend_names())
def test_every_census_code_is_declared(name):
    census = {c for c, f in FEATURES.items() if f.source == "census"}
    assert set(get_backend(name).support) == census


@pytest.mark.parametrize("name", backend_names())
def test_every_feature_is_shown_or_reported(every, name):
    backend = get_backend(name)
    rendered = sheetpeek.read(every, "All", platform="mac").layout(backend.mode)
    reported = {n.code for n in sheetpeek.notices(rendered, backend) if n.output == backend.name}
    for code in rendered.features:
        support = backend.support[code]
        assert support.outcome is None or code in reported, code
        if support.outcome is None:
            assert code not in reported, code


def test_an_undeclared_feature_is_reported(every):
    class Bare:
        name = "bare"
        mode = "none"

        def render(self, sheet, out):
            out.write("")

    rendered = sheetpeek.read(every, "All", platform="mac").layout("none")
    found = [n for n in sheetpeek.notices(rendered, Bare()) if n.output == "bare"]
    assert codes(found) == set(rendered.features)
    assert all(n.context.get("undeclared") == 1 and n.outcome == "omitted" for n in found)


def test_a_removed_declaration_is_reported(every, monkeypatch):
    # The mutation the completeness test guards against: an output forgets a feature.
    original = _support.text

    def without_fill(**options):
        table = original(**options)
        del table["fill"]
        return table

    monkeypatch.setattr(_support, "text", without_fill)
    found = [n for n in sheetpeek.notices(every, "text", sheet="All") if n.code == "fill"]
    assert [n.context.get("undeclared") for n in found] == [1]


# Plain workbooks: values, number formats and default styles, from three writers.


def _plain_rows(write):
    for row, values in enumerate(
        (("Name", "Amount", "Date"), ("a", 1.5, 45000), ("b", -2, 45001)), 1
    ):
        for column, value in enumerate(values, 1):
            write(row, column, value)


@pytest.mark.parametrize("name", backend_names())
def test_a_plain_openpyxl_workbook_gives_no_notice(tmp_path, name):
    wb = openpyxl.Workbook()
    _plain_rows(lambda r, c, v: wb.active.cell(r, c, v))
    for row in (2, 3):
        wb.active.cell(row, 2).number_format = "#,##0.00"
        wb.active.cell(row, 3).number_format = "yyyy-mm-dd"
    path = save(wb, tmp_path / "plain.xlsx")
    assert sheetpeek.notices(path, name) == ()


@pytest.mark.parametrize("name", backend_names())
def test_a_plain_xlsxwriter_workbook_gives_no_notice(tmp_path, name):
    xlsxwriter = pytest.importorskip("xlsxwriter")
    path = tmp_path / "plain.xlsx"
    workbook = xlsxwriter.Workbook(str(path))
    sheet = workbook.add_worksheet("Data")
    money = workbook.add_format({"num_format": "#,##0.00"})
    _plain_rows(lambda r, c, v: sheet.write(r - 1, c - 1, v, money if c == 2 and r > 1 else None))
    workbook.close()
    assert sheetpeek.notices(path, name) == ()


@pytest.mark.parametrize("name", backend_names())
def test_workbooks_saved_by_excel_give_no_reader_notice(name):
    files = sorted(FILES.glob("*.xlsx"))
    assert files
    for path in files:
        with sheetpeek.open(path, platform="mac") as book:
            assert book.notices == ()
            for sheet_name in book.sheet_names:
                found = sheetpeek.notices(book.sheet(sheet_name), name)
                assert [n for n in found if n.output is None] == [], path.name


# The used-not-present rules.


def test_boilerplate_is_not_reported(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    path = save(wb, tmp_path / "boilerplate.xlsx")
    rewrite(
        path,
        SHEET1,
        replace_once(
            "</worksheet>",
            '<sheetProtection sheet="1" /><autoFilter /><dataValidations count="1">'
            '<dataValidation type="whole" sqref="A1" /></dataValidations>'
            '<printOptions gridLines="1" />'
            "<headerFooter><oddHeader>&amp;CPage</oddHeader></headerFooter>"
            '<colBreaks count="1"><brk id="3" /></colBreaks>'
            '<tableParts count="0" /></worksheet>',
        ),
    )
    assert sheetpeek.read(path, platform="mac").notices == ()


def test_extension_uris_are_compared_as_excel_compares_them(tmp_path):
    def sheet_with(*uris):
        wb = openpyxl.Workbook()
        wb.active["A1"] = 1
        path = save(wb, tmp_path / f"ext{len(list(tmp_path.iterdir()))}.xlsx")
        rewrite(path, SHEET1, replace_once("</worksheet>", extensions(*uris) + "</worksheet>"))
        return codes(sheetpeek.read(path, platform="mac").notices)

    assert sheet_with(CF_URI) == {"cf"}
    # Excel ignores the upper-case form that [MS-XLSX] prints (captured): nothing is drawn.
    assert sheet_with(CF_URI.upper()) == set()
    # Not yet captured: matched ignoring case, so that nothing is missed.
    assert sheet_with("{05c60535-1f16-4fd2-b633-f4f36f0b64e0}") == {"sparkline"}
    assert sheet_with("{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}") == set()  # validation
    assert sheet_with("{11111111-2222-3333-4444-555555555555}") == {"sheet.unknown_extension"}


def test_a_linked_picture_is_refused_as_unsafe(every):
    found = [
        n
        for n in sheetpeek.read(every, "All", platform="mac").notices
        if n.code == "drawing.picture_linked"
    ]
    assert [(n.outcome, n.reason) for n in found] == [("omitted", "unsafe")]
    assert "never fetches" in found[0].message


def test_objects_hidden_by_the_workbook_are_not_reported(every, tmp_path):
    path = tmp_path / "hidden_objects.xlsx"
    path.write_bytes(every.read_bytes())
    rewrite(
        path,
        "xl/workbook.xml",
        lambda t: re.sub(r"<workbookPr\b", '<workbookPr showObjects="none"', t, count=1),
    )
    found = codes(sheetpeek.read(path, "All", platform="mac").notices)
    assert not {c for c in found if c.startswith("drawing") or c == "control"}


def test_a_dialog_sheet_is_reported_not_shown_blank(every):
    # openpyxl loads a dialog sheet as an empty worksheet; it is not drawn as a blank sheet.
    with sheetpeek.open(every, platform="mac") as book:
        found = [n for n in book.notices if n.code == "sheet.dialogsheet"]
        assert [(n.outcome, n.context["sheet"], n.level) for n in found] == [
            ("omitted", "Dialog", "workbook")
        ]
        with pytest.raises(sheetpeek.NotAWorksheetError, match="dialog sheet"):
            book.sheet("Dialog")


def test_chartsheets_and_dialog_sheets_are_listed_as_not_shown(every, capsys):
    with sheetpeek.open(every, platform="mac") as book:
        assert [(n.code, n.context["sheet"]) for n in book.notices] == [
            ("sheet.dialogsheet", "Dialog"),
            ("sheet.chartsheet", "Chart"),
        ]
    assert cli.main(["view", str(every), "--list"]) == 0
    assert capsys.readouterr().out.splitlines() == [
        "All",
        "Dialog  (dialog sheet, not shown)",
        "Macro  (macro sheet)",
        "Chart  (chartsheet, not shown)",
    ]


def test_cells_hidden_from_view_are_not_reported(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in (1, 2, 3):
        ws.cell(row, 1, "x").comment = Comment("a note", "author")
    ws.row_dimensions[2].hidden = True
    path = save(wb, tmp_path / "hidden_note.xlsx")
    sheet = sheetpeek.read(path, platform="mac")
    (whole,) = [n for n in sheet.notices if n.code == "comment"]
    assert whole.count == 3
    (shown,) = [n for n in sheet.layout("excel").notices if n.code == "comment"]
    assert (shown.count, [str(r) for r in shown.cells]) == (2, ["A1", "A3"])
    assert "(2 cells)" in shown.message


# The API.


def test_render_warns_once(every):
    with pytest.warns(sheetpeek.NotShownWarning) as caught:
        sheetpeek.render(every, "csv", sheet="All")
    ours = [w for w in caught if w.category is sheetpeek.NotShownWarning]
    assert len(ours) == 1
    assert "sheet 'All' to csv" in str(ours[0].message)


def test_on_notice_takes_every_notice_instead_of_the_warning(every):
    received = []
    with warnings.catch_warnings():
        warnings.simplefilter("error", sheetpeek.NotShownWarning)
        sheetpeek.render(every, "csv", sheet="All", on_notice=received.append)
        sheetpeek.render(every, "csv", sheet="All", on_notice="ignore")
    assert tuple(received) == sheetpeek.notices(every, "csv", sheet="All")


def test_nothing_to_report_issues_no_warning(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    path = save(wb, tmp_path / "one.xlsx")
    with warnings.catch_warnings():
        warnings.simplefilter("error", sheetpeek.NotShownWarning)
        sheetpeek.render(path, "html")


def test_a_bad_handler_is_refused(every):
    with pytest.raises(TypeError):
        sheetpeek.render(every, "csv", sheet="All", on_notice="loud")  # type: ignore[arg-type]


def test_notices_are_ordered_reader_layout_output(every):
    found = sheetpeek.notices(every, "html", sheet="All")
    stages = [
        "output" if n.output else FEATURES[n.code].source for n in found if n.code in FEATURES
    ]
    order = {"reader": 0, "census": 0, "layout": 1, "output": 2}
    assert stages == sorted(stages, key=lambda s: order[s])


def test_matches_takes_a_code_and_the_codes_under_it():
    notice = sheetpeek.Notice("border.diagonal", "omitted", "not_implemented", "m", "S")
    assert notice.matches("border")
    assert notice.matches("border.diagonal")
    assert not notice.matches("bord")
    assert not notice.matches("border.diagonal.x")
    assert notice.level == "cells"
    assert sheetpeek.Notice("cf", "omitted", "not_implemented", "m", "S").level == "sheet"
    assert (
        sheetpeek.Notice("sheet.chartsheet", "omitted", "not_implemented", "m").level == "workbook"
    )


def test_cells_are_coalesced_into_ranges():
    cells = CellSet()
    for row in (1, 2, 3):
        for column in (2, 3):
            cells.add(row, column)
    cells.add(5, 1)
    ranges, truncated = cells.ranges()
    assert ([str(r) for r in ranges], truncated, cells.count) == (["B1:C3", "A5"], False, 7)


def test_visible_splits_ranges_around_hidden_rows_and_columns():
    notice = sheetpeek.Notice(
        "border",
        "omitted",
        "not_implemented",
        "borders not shown (9 cells): x",
        "S",
        (sheetpeek.CellRange.parse("A1:C3"),),
        9,
    )
    (kept,) = visible([notice], {2}, {2})
    assert [str(r) for r in kept.cells] == ["A1", "C1", "A3", "C3"]
    assert kept.count == 4
    assert visible([notice], {1, 2, 3}, set()) == []


def test_the_warnings_view_keeps_only_the_old_warnings(every):
    sheet = sheetpeek.read(every, "All", platform="mac")
    assert len(sheet.notices) > 20
    assert [w[:9] for w in sheet.warnings] == ["1 formula"]


def test_an_html_page_can_list_what_it_does_not_show(every, tmp_path):
    path = tmp_path / "hostile.xlsx"
    path.write_bytes(every.read_bytes())
    rewrite(
        path,
        SHEET1,
        lambda t: t.replace("{00000000-0000-0000-0000-000000000000}", "&lt;script&gt;x"),
    )
    page = sheetpeek.render(
        path, sheetpeek.backends.HTMLBackend(notices=True), sheet="All", on_notice="ignore"
    )
    assert '<details class="sp-notices"' in page
    assert "unrecognized sheet extensions not shown" in page
    assert "<script" not in page
    plain = sheetpeek.render(path, "html", sheet="All", on_notice="ignore")
    assert "sp-notices" not in plain


# The command line.


def test_the_cli_reports_after_the_output(every, capsys):
    assert cli.main(["view", str(every), "-s", "All", "-t", "csv"]) == 0
    err = capsys.readouterr().err
    assert "sheetpeek: border: borders not shown" in err
    assert "sheetpeek: font.bold: bold text not shown in csv" in err
    assert "  cells:" not in err


def test_the_cli_notice_options(every, capsys):
    assert cli.main(["view", str(every), "-s", "All", "-t", "csv", "--notices", "none"]) == 0
    assert capsys.readouterr().err == ""
    assert cli.main(["view", str(every), "-s", "All", "-t", "csv", "--notices", "full"]) == 0
    assert "  cells: B2" in capsys.readouterr().err
    assert cli.main(["view", str(every), "-s", "All", "-t", "csv", "--strict"]) == 3
    err = capsys.readouterr().err
    assert "font.bold" in err
    assert cli.main(["view", str(every), "-s", "All", "-t", "csv", "--suppress", "font"]) == 0
    err = capsys.readouterr().err
    assert "sheetpeek: font." not in err
    assert "sheetpeek: border:" in err


def test_strict_passes_a_plain_workbook(tmp_path, capsys):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    path = save(wb, tmp_path / "one.xlsx")
    assert cli.main(["view", str(path), "-t", "text", "--strict"]) == 0
    assert capsys.readouterr().err == ""


def test_only_notices_are_issued_for_sheet_extensions(tmp_path):
    # openpyxl warns that it drops these extensions; sheetpeek reports them as notices instead.
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    path = save(wb, tmp_path / "extensions.xlsx")
    uris = (
        CF_URI,
        "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}",
        "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}",
        "{11111111-2222-3333-4444-555555555555}",
    )
    rewrite(path, SHEET1, replace_once("</worksheet>", extensions(*uris) + "</worksheet>"))
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        # garbage collection may emit an earlier test's unclosed file here
        warnings.simplefilter("ignore", ResourceWarning)
        sheetpeek.render(path, "html")
    assert [w.category for w in caught] == [sheetpeek.NotShownWarning]
    assert {"cf", "sparkline", "slicer", "sheet.unknown_extension"} <= codes(
        sheetpeek.read(path, platform="mac").notices
    )


def _row_column_notice(tmp_path, style_row=None, style_column=None, edit=None):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for ref in ("A1", "C3"):
        sheet[ref] = 1
    if style_row is not None:
        for key, value in style_row.items():
            setattr(sheet.row_dimensions[2], key, value)
    if style_column is not None:
        for key, value in style_column.items():
            setattr(sheet.column_dimensions["B"], key, value)
    path = save(workbook, tmp_path / "rows.xlsx")
    if edit is not None:
        rewrite(path, "xl/styles.xml", edit)
    found = [
        n for n in sheetpeek.read(path, platform="mac").notices if n.code == "style.row_column"
    ]
    return found[0].context if found else None


def test_a_row_fill_and_a_column_border_are_reported(tmp_path):
    fill = PatternFill("solid", fgColor="FF0000")
    assert _row_column_notice(tmp_path, style_row={"fill": fill}) == {"rows": 1, "columns": 0}
    side = Border(left=Side(style="thin"))
    assert _row_column_notice(tmp_path, style_column={"border": side}) == {
        "rows": 0,
        "columns": 1,
    }


def test_border_sides_of_style_none_draw_nothing_on_rows(tmp_path):
    side = Border(left=Side(style="thin"))

    def edit(text):
        return text.replace('style="thin"', 'style="none"')

    assert _row_column_notice(tmp_path, style_row={"border": side}, edit=edit) is None


def test_black_text_and_white_fills_are_not_reported_where_an_output_draws_them_so(tmp_path):
    # An output without colors draws black text on white paper: an explicit black font color or
    # a white solid fill loses nothing there. Any other color is reported, in runs too.
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "black"
    ws["A1"].font = Font(color="FF000000")
    ws["A2"] = "white"
    ws["A2"].fill = PatternFill("solid", fgColor="FFFFFFFF")
    ws["A3"] = CellRichText(["a", TextBlock(InlineFont(color="FF000000", b=True), "b")])
    quiet = save(wb, tmp_path / "quiet.xlsx")
    assert {"font.color", "fill"}.isdisjoint(codes(sheetpeek.notices(quiet, "text")))
    ws["B1"] = "red"
    ws["B1"].font = Font(color="FFFF0000")
    ws["B2"] = "yellow"
    ws["B2"].fill = PatternFill("solid", fgColor="FFFFFF00")
    ws["B3"] = CellRichText(["a", TextBlock(InlineFont(color="FF0000FF"), "b")])
    loud = save(wb, tmp_path / "loud.xlsx")
    found = {n.code: n for n in sheetpeek.notices(loud, "text")}
    assert found["font.color"].cells == (
        sheetpeek.CellRange.parse("B1"),
        sheetpeek.CellRange.parse("B3"),
    )
    assert found["fill"].cells == (sheetpeek.CellRange.parse("B2"),)
