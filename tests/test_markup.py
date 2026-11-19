"""No string a workbook can supply becomes markup, in any backend that takes markup.

Every source below puts the same payload in one string a file controls. Every output is parsed,
never searched: HTML with ``html.parser``, Markdown through a CommonMark renderer that allows HTML
and then ``html.parser``. What the parser finds must be only what sheetpeek writes itself. The
renderer is a development dependency: without it the Markdown tests skip, and the others run.
"""

import http.client
import re
import threading
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urlsplit
from xml.sax.saxutils import escape, quoteattr

import openpyxl
import pytest
from openpyxl.styles import Alignment, Border, Font, GradientFill, PatternFill, Side
from openpyxl.styles.colors import Color

import sheetpeek
from sheetpeek import PackageError
from sheetpeek._serve import make_server
from sheetpeek.backends import HTMLBackend, MarkdownBackend
from tests.workbooks import (
    NS,
    R_NS,
    SHEET1,
    link_shared_strings,
    replace_once,
    replace_part,
    replace_sheet_data,
    rewrite,
    save,
    shared_strings,
)

MARK = "PWN"
PAYLOAD = (
    f"\"'></title></style></textarea><{MARK} ONERROR=x>&lt;{MARK}&gt;</td><script>x</script>"
    f"[l](javascript:{MARK})![i](x:{MARK})<http://{MARK}.example>`c`*e*_f_~~s~~&amp;"
    f"'; background:url({MARK}); x:'\\"
)
HTML_TAGS = {
    "html", "head", "meta", "title", "style", "body", "div", "table", "colgroup", "col", "thead",
    "tbody", "tr", "td", "th", "span", "details", "summary", "ul", "li", "a", "i",
}  # fmt: skip
HTML_ATTRIBUTES = {
    "data-color-scheme",
    "lang",
    "charset",
    "name",
    "content",
    "class",
    "style",
    "colspan",
    "rowspan",
    "scope",
    "href",
    "rel",
    "title",
    "id",
}
STYLE_PROPERTIES = {
    "width", "height", "left", "right", "top", "font", "color", "background", "border-color",
    "margin-top", "border-left-color", "border-top-color", "letter-spacing", "text-decoration",
    "background-position", "line-height", "position",
}  # fmt: skip
MARKDOWN_TAGS = {"table", "thead", "tbody", "tr", "th", "td", "br", "p", "strong", "em", "a"}
LINK_SCHEMES = ("http", "https", "mailto")


# A cell's id on a one-sheet page, and a sheet's or its cell's on the workbook page.
IDS = r"sp-[A-Z]+[0-9]+|sp-[0-9]+(?:-[A-Z]+[0-9]+)?"


def safe_href(value: str) -> bool:
    """Whether a link goes somewhere sheetpeek allows: a page of serve, a cell, or a URL."""
    if value.startswith(("/", "#sp-")):
        return value.startswith("/") or re.fullmatch(f"#(?:{IDS})", value) is not None
    return urlsplit(value).scheme.lower() in LINK_SCHEMES


def markdown_renderer():
    """A CommonMark renderer with tables and bare links, that allows HTML; or the test skips."""
    markdown_it = pytest.importorskip("markdown_it")
    pytest.importorskip("linkify_it")  # rich brings markdown_it without it
    return markdown_it.MarkdownIt("gfm-like", {"html": True, "linkify": True})


@dataclass
class Found:
    tags: list[str] = field(default_factory=list)
    attributes: list[tuple[str, str]] = field(default_factory=list)
    style_elements: list[str] = field(default_factory=list)
    text: list[str] = field(default_factory=list)


def parse(markup: str) -> Found:
    found = Found()

    class Parser(HTMLParser):
        in_style = False

        def handle_starttag(self, tag, attrs):
            found.tags.append(tag)
            found.attributes += [(name, value or "") for name, value in attrs]
            self.in_style = tag == "style"

        def handle_endtag(self, tag):
            self.in_style = False

        def handle_data(self, data):
            (found.style_elements if self.in_style else found.text).append(data)

        def handle_comment(self, data):
            raise AssertionError(f"a comment in the output: {data!r}")

        def handle_decl(self, decl):
            assert decl.lower() == "doctype html"

        def unknown_decl(self, data):
            raise AssertionError(f"a declaration in the output: {data!r}")

        def handle_pi(self, data):
            raise AssertionError(f"a processing instruction in the output: {data!r}")

    parser = Parser(convert_charrefs=True)
    parser.feed(markup)
    parser.close()
    return found


def declarations(style: str) -> list[tuple[str, str]]:
    """A style attribute's declarations, split as CSS splits them: quotes hide what is in them."""
    result, current, quote, skip = [], "", None, False
    for char in style:
        if skip:
            current, skip = current + char, False
        elif quote:
            current += char
            skip, quote = char == "\\", None if char == quote else quote
        elif char in "'\"":
            current, quote = current + char, char
        elif char == ";":
            result.append(current)
            current = ""
        else:
            current += char
    assert quote is None, f"an open string in {style!r}"
    result.append(current)
    pairs = []
    for declaration in filter(None, map(str.strip, result)):
        name, _, value = declaration.partition(":")
        pairs.append((name.strip(), value.strip()))
    return pairs


def unquoted(value: str) -> str:
    return re.sub(r"'(?:[^'\\]|\\.)*'|\"(?:[^\"\\]|\\.)*\"", "''", value)


def assert_html_is_ours(
    markup: str, *, extra_tags=frozenset(), extra_attributes=frozenset()
) -> Found:
    found = parse(markup)
    assert set(found.tags) <= HTML_TAGS | extra_tags, set(found.tags) - HTML_TAGS
    for name, value in found.attributes:
        assert name in HTML_ATTRIBUTES | extra_attributes, (name, value)
        if name == "style":
            for prop, css in declarations(value):
                assert prop in STYLE_PROPERTIES, (prop, css)
                assert MARK not in unquoted(css), (prop, css)
                assert "url(" not in unquoted(css).lower(), (prop, css)
        elif name == "class":
            assert re.fullmatch(r"[a-z -]+", value), value
        elif name == "href":
            assert safe_href(value), value
        elif name == "rel":
            assert value == "noopener noreferrer", value
        elif name == "id":
            assert re.fullmatch(IDS, value), value
        elif name not in (
            "title",  # a tooltip from the file, which the parser shows is only an attribute
            "content",
            "charset",
            "lang",
            "colspan",
            "rowspan",
            "scope",
            "data-color-scheme",
        ):
            assert MARK not in value, (name, value)
    for sheet in found.style_elements:
        assert MARK not in sheet
    return found


# --- every string a file controls


def xml_text(tmp_path: Path, name: str = "book.xlsx", **cells) -> Path:
    wb = openpyxl.Workbook()
    wb.active["A1"] = "first"
    wb.active["B1"] = "header"
    for ref, value in cells.items():
        wb.active[ref] = value
    return save(wb, tmp_path / name)


def cell_xml(tmp_path, xml):
    path = xml_text(tmp_path)
    data = f'<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2">{xml}</row></sheetData>'
    return rewrite(path, SHEET1, replace_sheet_data(data))


def inline(tmp_path):
    return cell_xml(tmp_path, f'<c r="A2" t="inlineStr"><is><t>{escape(PAYLOAD)}</t></is></c>')


def cached(tmp_path):
    # A function xlrecalc does not compute keeps the value the file holds.
    formula = "CUBEVALUE(&quot;x&quot;)"
    return cell_xml(tmp_path, f'<c r="A2" t="str"><f>{formula}</f><v>{escape(PAYLOAD)}</v></c>')


def computed(tmp_path):
    # A string a formula computes on opening (the file holds no value; xlrecalc computes it).
    literal = escape('"' + PAYLOAD.replace('"', '""') + '"')
    return cell_xml(tmp_path, f'<c r="A2" t="str"><f>{literal}</f></c>')


def error_text(tmp_path):
    return cell_xml(tmp_path, f'<c r="A2" t="e"><v>{escape(PAYLOAD)}</v></c>')


def error_shaped(tmp_path):
    return cell_xml(tmp_path, f'<c r="A2" t="e"><v>#{MARK}&lt;{MARK}&gt;!</v></c>')


def date_text(tmp_path):
    return cell_xml(tmp_path, f'<c r="A2" t="d"><v>{escape(PAYLOAD)}</v></c>')


def shared(tmp_path):
    path = link_shared_strings(cell_xml(tmp_path, '<c r="A2" t="s"><v>0</v></c>'))
    strings = f"<sst {NS}><si><t>{escape(PAYLOAD)}</t></si></sst>"
    return replace_part(path, "xl/sharedStrings.xml", strings.encode())


def escaped_in_file(tmp_path):
    # _x003C_ decodes to <, after the file's XML has been parsed.
    text = PAYLOAD.replace("<", "_x003C_").replace(">", "_x003E_").replace("&", "_x0026_")
    return cell_xml(tmp_path, f'<c r="A2" t="inlineStr"><is><t>{escape(text)}</t></is></c>')


def line_breaks(tmp_path):
    text = f"a_x000A_{PAYLOAD}_x000D_{PAYLOAD}_x000D__x000A_{PAYLOAD}"
    return cell_xml(tmp_path, f'<c r="A2" t="inlineStr"><is><t>{escape(text)}</t></is></c>')


def _styled(tmp_path, **options):
    path = xml_text(tmp_path, A2=PAYLOAD, B2="neighbor", C3="right")
    wb = openpyxl.load_workbook(path)
    for name, value in options.items():
        setattr(wb.active["A2"], name, value)
    return save(wb, path)


def wrapped(tmp_path):
    return _styled(tmp_path, alignment=Alignment(wrap_text=True))


def filled_alignment(tmp_path):
    return _styled(tmp_path, alignment=Alignment(horizontal="fill"))


def justified(tmp_path):
    return _styled(tmp_path, alignment=Alignment(horizontal="distributed"))


def spilled_left(tmp_path):
    path = xml_text(tmp_path, D2=PAYLOAD)
    wb = openpyxl.load_workbook(path)
    wb.active["D2"].alignment = Alignment(horizontal="right")
    return save(wb, path)


def merged(tmp_path):
    path = xml_text(tmp_path, A2=PAYLOAD)
    wb = openpyxl.load_workbook(path)
    wb.active.merge_cells("A2:C3")
    return save(wb, path)


def filled_and_bordered(tmp_path):
    path = _styled(
        tmp_path,
        fill=PatternFill("gray0625", fgColor="FF0000FF"),
        border=Border(
            left=Side("thick", color="FFFF0000"),
            bottom=Side("dashDot"),
            diagonal=Side("double"),
            diagonalUp=True,
        ),
    )
    wb = openpyxl.load_workbook(path)
    wb.active["B2"].fill = GradientFill(degree=30, stop=("FFFFFFFF", "FF0000FF"))
    wb.active["C3"].fill = GradientFill("path", left=0.2, right=0.8, stop=("FFFFFFFF", "FF0000FF"))
    return save(wb, path)


def border_rgb(tmp_path):
    path = _styled(tmp_path, border=Border(left=Side("thin", color="00ABCDEF")))
    return rewrite(path, "xl/styles.xml", lambda t: t.replace('"00ABCDEF"', quoteattr(PAYLOAD)))


def gradient_rgb(tmp_path):
    path = _styled(tmp_path, fill=GradientFill(stop=("00ABCDEF", "FF0000FF")))
    return rewrite(path, "xl/styles.xml", lambda t: t.replace('"00ABCDEF"', quoteattr(PAYLOAD)))


def hidden(tmp_path):
    path = xml_text(tmp_path, A2=PAYLOAD, C2=PAYLOAD)
    wb = openpyxl.load_workbook(path)
    wb.active.column_dimensions["C"].hidden = True
    wb.active.row_dimensions[2].hidden = True
    return save(wb, path)


def _number_format(tmp_path, code, value=1):
    path = xml_text(tmp_path)
    wb = openpyxl.load_workbook(path)
    wb.active["A2"] = value
    wb.active["A2"].number_format = "0.000"
    wb.active.column_dimensions["A"].width = 200
    save(wb, path)
    return rewrite(path, "xl/styles.xml", lambda t: t.replace('"0.000"', quoteattr(code)))


def format_quoted(tmp_path):
    text = PAYLOAD.replace('"', "")
    return _number_format(tmp_path, f'"{text}"0')


def format_text_section(tmp_path):
    text = PAYLOAD.replace('"', "")
    return _number_format(tmp_path, f'@"{text}"', "value")


def format_escaped(tmp_path):
    return _number_format(tmp_path, "0" + "".join(f"\\{c}" for c in f"<{MARK}>"))


def format_currency(tmp_path):
    return _number_format(tmp_path, f"[$<{MARK}>-409]0")


def format_refused(tmp_path):
    return _number_format(tmp_path, f"0 n<{MARK}>")


def sheet_name(tmp_path):
    path = xml_text(tmp_path, A2="x")
    return rewrite(
        path,
        "xl/workbook.xml",
        lambda t: t.replace('name="Sheet"', f"name={quoteattr(PAYLOAD)}"),
    )


def font_name(tmp_path):
    path = _styled(tmp_path, font=Font(name="Zed"))
    return rewrite(
        path, "xl/styles.xml", lambda t: t.replace('val="Zed"', f"val={quoteattr(PAYLOAD)}")
    )


def normal_font_name(tmp_path):
    path = xml_text(tmp_path, A2="x")
    return rewrite(
        path,
        "xl/styles.xml",
        # Without its scheme: a theme font is drawn in the theme's typeface.
        lambda t: t.replace('val="Calibri"', f"val={quoteattr(PAYLOAD)}").replace(
            '<scheme val="minor" />', ""
        ),
    )


def theme_font_name(tmp_path):
    path = xml_text(tmp_path, A2="x")
    return rewrite(
        path,
        "xl/theme/theme1.xml",
        lambda t: re.sub(
            r'(<a:minorFont>\s*<a:latin typeface=)"Calibri"',
            lambda m: m[1] + quoteattr(PAYLOAD),
            t,
            count=1,
        ),
    )


def _themed(tmp_path, element):
    wb = openpyxl.Workbook()
    for ref in ("A1", "A2"):
        wb.active[ref] = "x"
        wb.active[ref].fill = PatternFill("solid", fgColor=Color(theme=4))
        wb.active[ref].font = Font(color=Color(theme=4))
    path = save(wb, tmp_path / "book.xlsx")
    new = f"<a:accent1><a:{element}={quoteattr(PAYLOAD)}/></a:accent1>"
    return rewrite(
        path,
        "xl/theme/theme1.xml",
        lambda t: re.sub(r"<a:accent1>.*?</a:accent1>", lambda _: new, t, flags=re.S),
    )


def theme_srgb(tmp_path):
    return _themed(tmp_path, "srgbClr val")


def theme_system(tmp_path):
    return _themed(tmp_path, "sysClr lastClr")


def style_rgb(tmp_path):
    path = _styled(
        tmp_path, font=Font(color="00ABCDEF"), fill=PatternFill("solid", fgColor="00ABCDEF")
    )
    return rewrite(path, "xl/styles.xml", lambda t: t.replace('"00ABCDEF"', quoteattr(PAYLOAD)))


def palette(tmp_path):
    path = _styled(tmp_path, font=Font(color=Color(indexed=8)))
    colors = f"<colors><indexedColors><rgbColor rgb={quoteattr(PAYLOAD)}/></indexedColors></colors>"
    return rewrite(path, "xl/styles.xml", replace_once("</styleSheet>", colors + "</styleSheet>"))


def hyperlink(tmp_path):
    path = xml_text(tmp_path, A2="x")
    link = f'<hyperlinks><hyperlink ref="A2" location={quoteattr(PAYLOAD)}/></hyperlinks>'
    return rewrite(path, SHEET1, replace_once("</sheetData>", "</sheetData>" + link))


def _linked(tmp_path, target, **attributes):
    """A2 linked to an external ``target`` through a relationship, with other attributes."""
    path = xml_text(tmp_path, A2="x")
    extra = "".join(f" {name}={quoteattr(value)}" for name, value in attributes.items())
    link = f'<hyperlinks><hyperlink ref="A2" r:id="rIdL"{extra}/></hyperlinks>'
    rewrite(
        path,
        SHEET1,
        lambda t: t.replace("</sheetData>", "</sheetData>" + link, 1).replace(
            "<worksheet ", f"<worksheet {R_NS} ", 1
        ),
    )
    rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        f'relationships/hyperlink" Target={quoteattr(target)} TargetMode="External"/>'
        "</Relationships>"
    )
    return replace_part(path, "xl/worksheets/_rels/sheet1.xml.rels", rels.encode())


def hyperlink_target(tmp_path):
    return _linked(tmp_path, "https://example.com/" + PAYLOAD)


def hyperlink_mailto(tmp_path):
    return _linked(tmp_path, "mailto:" + PAYLOAD)


def hyperlink_refused(tmp_path):
    return _linked(tmp_path, "javascript:" + PAYLOAD, tooltip="t")


def hyperlink_tab(tmp_path):
    return _linked(tmp_path, f"java\tscript:{MARK}", location=PAYLOAD)


def hyperlink_tooltip(tmp_path):
    return _linked(tmp_path, "https://example.com/", tooltip=PAYLOAD, display=PAYLOAD)


def hyperlink_formula(tmp_path):
    target = escape("https://example.com/" + PAYLOAD.replace('"', '""'))
    return cell_xml(tmp_path, f'<c r="A2" t="str"><f>HYPERLINK("{target}")</f><v>x</v></c>')


def _rich(tmp_path, runs, **options):
    """A2 holding a shared string of rich-text runs, as Excel writes them."""
    path = _styled(tmp_path, **options) if options else xml_text(tmp_path, A2="x")
    return shared_strings(path, {"A2": runs})


_THIRD = len(PAYLOAD) // 3
RUNS = (
    f"<r><rPr><b/></rPr><t>{escape(PAYLOAD[:_THIRD])}</t></r>"
    f'<r><rPr><i/><vertAlign val="superscript"/><sz val="20"/></rPr>'
    f'<t xml:space="preserve">{escape(PAYLOAD[_THIRD : 2 * _THIRD])}</t></r>'
    f'<r><rPr><u val="doubleAccounting"/><strike/><color rgb="FF00FF00"/></rPr>'
    f"<t>{escape(PAYLOAD[2 * _THIRD :])}</t></r>"
)


def runs(tmp_path):
    return _rich(tmp_path, RUNS)


def runs_wrapped(tmp_path):
    return _rich(tmp_path, RUNS, alignment=Alignment(wrap_text=True))


def runs_filled(tmp_path):
    return _rich(tmp_path, RUNS, alignment=Alignment(horizontal="fill", shrink_to_fit=True))


def run_font_name(tmp_path):
    return _rich(
        tmp_path, f"<r><rPr><rFont val={quoteattr(PAYLOAD)}/></rPr><t>x</t></r><r><t>y</t></r>"
    )


def run_rgb(tmp_path):
    return _rich(tmp_path, f"<r><rPr><color rgb={quoteattr(PAYLOAD)}/></rPr><t>x</t></r>")


def inline_runs(tmp_path):
    runs = f"<r><rPr><b/></rPr><t>{escape(PAYLOAD)}</t></r><r><t>y</t></r>"
    return cell_xml(tmp_path, f'<c r="A2" t="inlineStr"><is>{runs}</is></c>')


def properties(tmp_path):
    path = xml_text(tmp_path, A2="x")
    return rewrite(path, "docProps/core.xml", replace_once("openpyxl", escape(PAYLOAD)))


def other_sheet_name(tmp_path):
    """A second sheet, hidden, named by the payload, with a tab color, and a link to it."""
    wb = openpyxl.Workbook()
    wb.active["A1"] = "first"
    wb.active["A2"] = "x"
    other = wb.create_sheet("Other")
    other["B2"] = "there"
    other.sheet_properties.tabColor = "00FF00"
    path = save(wb, tmp_path / "book.xlsx")
    quoted = "'" + PAYLOAD.replace("'", "''") + "'!B2"
    link = f'<hyperlinks><hyperlink ref="A2" location={quoteattr(quoted)}/></hyperlinks>'
    rewrite(path, SHEET1, replace_once("</sheetData>", "</sheetData>" + link))
    return rewrite(
        path,
        "xl/workbook.xml",
        replace_once(
            'name="Other" sheetId="2" state="visible"',
            f'name={quoteattr(PAYLOAD)} sheetId="2" state="hidden"',
        ),
    )


def tab_color(tmp_path):
    path = xml_text(tmp_path, A2="x")
    color = f"<sheetPr><tabColor rgb={quoteattr(PAYLOAD)}/></sheetPr>"
    return rewrite(
        path,
        SHEET1,
        lambda t: re.sub(r"<sheetPr>.*?</sheetPr>|(?=<dimension)", color, t, count=1, flags=re.S),
    )


def defined_name(tmp_path):
    path = xml_text(tmp_path, A2="x")
    name = f'<definedNames><definedName name="n">{escape(PAYLOAD)}</definedName></definedNames>'
    return rewrite(path, "xl/workbook.xml", replace_once("</sheets>", "</sheets>" + name))


def theme_dark(tmp_path):
    """The theme's dark 1, which every default font's color now resolves to."""
    path = xml_text(tmp_path, A2="x")
    new = f'<a:dk1><a:sysClr val="windowText" lastClr={quoteattr(PAYLOAD)}/></a:dk1>'
    return rewrite(
        path,
        "xl/theme/theme1.xml",
        lambda t: re.sub(r"<a:dk1>.*?</a:dk1>", lambda _: new, t, flags=re.S),
    )


def _table(tmp_path, edit):
    """A2:B4 a table in TableStyleMedium2 with a totals row, its part edited by ``edit``."""
    from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A2"], ws["B2"], ws["A3"], ws["B3"], ws["A4"] = "a", "b", 1, 2, "Total"
    table = Table(displayName="Sales", ref="A2:B4", totalsRowCount=1)
    table.tableColumns = [TableColumn(id=1, name="a", totalsRowLabel="Total"),
                          TableColumn(id=2, name="b")]  # fmt: skip
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    path = save(wb, tmp_path / "book.xlsx")
    return rewrite(path, "xl/tables/table1.xml", edit)


def table_name(tmp_path):
    return _table(
        tmp_path, lambda t: t.replace('displayName="Sales"', f"displayName={quoteattr(PAYLOAD)}")
    )


def table_style_name(tmp_path):
    return _table(tmp_path, lambda t: t.replace('"TableStyleMedium2"', quoteattr(PAYLOAD)))


def table_column_name(tmp_path):
    return _table(tmp_path, lambda t: t.replace('name="b"', f"name={quoteattr(PAYLOAD)}"))


def table_totals_label(tmp_path):
    return _table(
        tmp_path,
        lambda t: t.replace('totalsRowLabel="Total"', f"totalsRowLabel={quoteattr(PAYLOAD)}"),
    )


def _own_table_style(tmp_path):
    """A2:B3 a table in the workbook's own style ``Mine``: a red header fill and font color."""
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.styles.table import TableStyle, TableStyleElement
    from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A2"], ws["B2"], ws["A3"] = "a", "b", 1
    dxf = DifferentialStyle(font=Font(color="FFFF0000"), fill=PatternFill(bgColor="FFFF0000"))
    element = TableStyleElement(type="headerRow", dxfId=wb._differential_styles.add(dxf))
    wb._table_styles.tableStyle.append(
        TableStyle(name="Mine", pivot=False, tableStyleElement=[element])
    )
    table = Table(displayName="Sales", ref="A2:B3")
    table.tableColumns = [TableColumn(id=1, name="a"), TableColumn(id=2, name="b")]
    table.tableStyleInfo = TableStyleInfo(name="Mine")
    ws.add_table(table)
    return save(wb, tmp_path / "book.xlsx")


def table_dxf_rgb(tmp_path):
    path = _own_table_style(tmp_path)
    return rewrite(path, "xl/styles.xml", lambda t: t.replace('"FFFF0000"', quoteattr(PAYLOAD)))


def table_own_style_name(tmp_path):
    path = _own_table_style(tmp_path)
    path = rewrite(
        path, "xl/styles.xml", lambda t: t.replace('name="Mine"', f"name={quoteattr(PAYLOAD)}")
    )
    return rewrite(path, "xl/tables/table1.xml", lambda t: t.replace('"Mine"', quoteattr(PAYLOAD)))


SOURCES = [
    inline, cached, computed, error_text, error_shaped, date_text, shared, escaped_in_file,
    line_breaks, wrapped, filled_alignment, justified, spilled_left, merged, filled_and_bordered,
    border_rgb, gradient_rgb, hidden, format_quoted,
    format_text_section, format_escaped, format_currency, format_refused, sheet_name, font_name,
    normal_font_name, theme_srgb, theme_system, style_rgb, palette, hyperlink, hyperlink_target,
    hyperlink_mailto, hyperlink_refused, hyperlink_tab, hyperlink_tooltip, hyperlink_formula,
    properties, defined_name,
    theme_font_name,
    runs, runs_wrapped, runs_filled, run_font_name, run_rgb,
    inline_runs,
    other_sheet_name, tab_color,
    theme_dark, table_name, table_style_name, table_column_name, table_totals_label,
    table_dxf_rgb, table_own_style_name,
]  # fmt: skip
# Whichever of these openpyxl validates, it refuses when it loads the file: nothing is shown.
BACKENDS = {
    "page": HTMLBackend(),
    "page, dark": HTMLBackend(color_scheme="dark"),
    "page, headings": HTMLBackend(column_headings=True, row_headings=True),
    "fragment": HTMLBackend(fragment=True),
    "fragment, headings, as serve": HTMLBackend(
        fragment=True, column_headings=True, row_headings=True
    ),
}
MARKDOWN = {
    "first": MarkdownBackend(),
    "letters": MarkdownBackend("letters"),
    "first, emphasis": MarkdownBackend(emphasis=True),
    "letters, emphasis": MarkdownBackend("letters", emphasis=True),
}
LAYOUTS = {
    "from the sheet": lambda s: s,
    "excel": lambda s: s.layout("excel"),
    "excel, hidden shown": lambda s: s.layout("excel", show_hidden=True),
    "characters": lambda s: s.layout("characters"),
    "characters, hidden shown": lambda s: s.layout("characters", show_hidden=True),
    "none": lambda s: s.layout("none"),
    "none, hidden shown": lambda s: s.layout("none", show_hidden=True),
}


def read_or_refused(path):
    try:
        return sheetpeek.read(path, limits=sheetpeek.ReadLimits(max_cells=1_000_000))
    except PackageError:
        return None


# Where each source's payload ends up. A source that stops reaching the output tests nothing, so
# this table is checked: "refused" raises on load, "absent" is by design never shown (a theme
# without valid colors, a hidden cell, a refused format's General, what no backend draws).
REACHES = {
    "refused": {
        "date_text", "error_text", "error_shaped",
    },
    "absent": {
        "hidden", "format_refused", "theme_srgb", "theme_system", "hyperlink", "properties",
        "tab_color", "defined_name", "run_rgb", "hyperlink_refused", "hyperlink_tab",
        "other_sheet_name", "theme_dark", "table_name", "table_style_name", "table_column_name",
        "table_totals_label", "table_dxf_rgb", "table_own_style_name",
    },
    "attribute": {
        "normal_font_name", "hyperlink_target", "hyperlink_mailto", "hyperlink_tooltip",
        "hyperlink_formula", "theme_font_name", "run_font_name",
    },
}  # fmt: skip


@pytest.mark.parametrize("build", SOURCES, ids=lambda f: f.__name__)
def test_each_source_reaches_the_output_as_expected(build, tmp_path):
    sheet = read_or_refused(build(tmp_path))
    if build.__name__ in REACHES["refused"]:
        assert sheet is None
        return
    page = parse(sheetpeek.render(sheet, HTMLBackend()))
    text = MARK in "".join(page.text)
    attribute = any(MARK in value for _, value in page.attributes)
    if build.__name__ in REACHES["absent"]:
        assert not text
        assert not attribute
    elif build.__name__ in REACHES["attribute"]:
        assert attribute
    else:
        assert text


@pytest.fixture(params=SOURCES, ids=lambda f: f.__name__)
def source(request, tmp_path):
    return request.param(tmp_path)


@pytest.mark.parametrize("layout", LAYOUTS.values(), ids=LAYOUTS.keys())
@pytest.mark.parametrize("backend", BACKENDS.values(), ids=BACKENDS.keys())
def test_html_holds_no_markup_from_the_file(source, backend, layout):
    sheet = read_or_refused(source)
    if sheet is not None:
        assert_html_is_ours(sheetpeek.render(layout(sheet), backend))


@pytest.mark.parametrize("layout", LAYOUTS.values(), ids=LAYOUTS.keys())
@pytest.mark.parametrize("backend", MARKDOWN.values(), ids=MARKDOWN.keys())
def test_markdown_holds_no_markup_from_the_file(source, backend, layout):
    renderer = markdown_renderer()
    sheet = read_or_refused(source)
    if sheet is None:
        return
    rendered = sheetpeek.render(layout(sheet), backend)
    found = parse(renderer.render(rendered))
    assert set(found.tags) <= MARKDOWN_TAGS, set(found.tags) - MARKDOWN_TAGS
    assert all(
        (name == "style" and re.fullmatch(r"text-align:\w+", value))
        or (name == "href" and safe_href(value))
        for name, value in found.attributes
    ), found.attributes
    assert rendered.count("\n") == sheet_lines(rendered), rendered


def sheet_lines(rendered):
    return sum(1 for line in rendered.splitlines() if line.startswith("| ") and line.endswith(" |"))


@pytest.mark.parametrize("layout", LAYOUTS.values(), ids=LAYOUTS.keys())
def test_a_cell_is_shown_as_its_literal_text_in_html(tmp_path, layout):
    # The other half of correct escaping: nothing is lost, in the text the reader sees.
    sheet = read_or_refused(inline(tmp_path))
    page = assert_html_is_ours(sheetpeek.render(layout(sheet), HTMLBackend()))
    assert PAYLOAD in "".join(page.text)


@pytest.mark.parametrize("layout", LAYOUTS.values(), ids=LAYOUTS.keys())
def test_a_cell_is_shown_as_its_literal_text_in_markdown(tmp_path, layout):
    renderer = markdown_renderer()
    sheet = read_or_refused(inline(tmp_path))
    rendered = sheetpeek.render(layout(sheet), "markdown")
    assert PAYLOAD in "".join(parse(renderer.render(rendered)).text)


WORKBOOK_PAGES = {
    "page": {},
    "hidden": {"hidden_sheets": True, "show_hidden": True},
    "tabs on top": {"tabs": "top"},
    "fragment, headings, notices": {
        "fragment": True,
        "column_headings": True,
        "row_headings": True,
        "notices": True,
        "hidden_sheets": True,
        "show_hidden": True,
    },
    "dark": {"color_scheme": "dark", "hidden_sheets": True},
}
BOOK_TAGS = frozenset({"nav", "section", "p"})


@pytest.mark.parametrize("options", WORKBOOK_PAGES.values(), ids=WORKBOOK_PAGES.keys())
def test_the_workbook_page_holds_no_markup_from_the_file(source, options):
    try:
        page = sheetpeek.render_workbook(
            source, limits=sheetpeek.ReadLimits(max_cells=1_000_000), on_notice="ignore", **options
        )
    except PackageError:
        return
    assert_html_is_ours(page, extra_tags=BOOK_TAGS)


def test_the_workbook_page_shows_sheet_names_as_text(tmp_path):
    path = other_sheet_name(tmp_path)
    page = sheetpeek.render_workbook(path, hidden_sheets=True, on_notice="ignore")
    found = assert_html_is_ours(page, extra_tags=BOOK_TAGS)
    assert PAYLOAD in "".join(found.text)  # the tab's label
    assert ("href", "#sp-2-B2") in found.attributes  # the link, by the sheet's position
    assert ("id", "sp-2-B2") in found.attributes
    assert not any(MARK in value for _, value in found.attributes)
    hidden = sheetpeek.render_workbook(path, on_notice="ignore")
    assert MARK not in hidden


def test_the_payload_would_be_markup_in_html_if_it_were_not_escaped():
    # So that the tests above can fail. The closers matter: a parser, like a browser, reads what
    # follows <title> as text until the element is closed.
    assert MARK.lower() in parse(f"<title>{PAYLOAD}</title>").tags
    found = parse(f"<td>{PAYLOAD}</td>")
    assert MARK.lower() in found.tags
    assert "script" in found.tags
    with pytest.raises(AssertionError):
        assert_html_is_ours(f"<td>{PAYLOAD}</td>")
    with pytest.raises(AssertionError):
        assert_html_is_ours('<td style="background:url(x)"></td>')
    with pytest.raises(AssertionError):
        assert_html_is_ours("<td style=\"font:11pt 'a'; x:y\"></td>")


def test_the_payload_would_be_markup_in_markdown_if_it_were_not_escaped():
    rendered = markdown_renderer().render(f"| h |\n| --- |\n| {PAYLOAD} |\n")
    assert set(parse(rendered).tags) - MARKDOWN_TAGS


# --- serve, as a client sees it


FILE_NAME = f"\"'><{MARK} a=b>.xlsx"


@pytest.fixture
def served(tmp_path):
    folder = tmp_path / f"<{MARK}>folder"
    folder.mkdir()
    for build in (
        inline, sheet_name, font_name, format_quoted, theme_srgb, palette, merged,
        hyperlink_target, hyperlink_tooltip, hyperlink_refused,
    ):  # fmt: skip
        place = tmp_path / build.__name__
        place.mkdir()
        (folder / f"{build.__name__}.xlsx").write_bytes(build(place).read_bytes())
    save(openpyxl.Workbook(), folder / FILE_NAME)
    replace_part(
        save(openpyxl.Workbook(), folder / "refused.xlsx"), f"xl/<{MARK}>.xml", b"<!DOCTYPE x><x/>"
    )
    server = make_server(folder, port=0)
    thread = threading.Thread(target=server.serve_forever, args=(0.01,), daemon=True)
    thread.start()
    yield server
    server.shutdown()
    server.server_close()


SERVE_TAGS = frozenset({"a"})
SERVE_ATTRIBUTES = frozenset({"href"})


def page_of(server, path):
    connection = http.client.HTTPConnection(*server.server_address[:2], timeout=5)
    connection.request("GET", path)
    response = connection.getresponse()
    body = response.read().decode()
    connection.close()
    return response.status, body


NAME = "%22'%3E%3CPWN%20a=b%3E.xlsx"
ENCODED = "%22%3E%3CPWN%20ONERROR=x%3E"
REQUESTS = [
    "/",
    f"/view/{NAME}",
    f"/view/{ENCODED}.xlsx",
    f"/{ENCODED}",
    "/view/refused.xlsx",
    f"/view/inline.xlsx?sheet={ENCODED}",
    f"/view/inline.xlsx?ref={ENCODED}",
    "/view/inline.xlsx",
    "/view/sheet_name.xlsx",
    "/view/font_name.xlsx",
    "/view/format_quoted.xlsx",
    "/view/theme_srgb.xlsx",
    "/view/palette.xlsx",
    "/view/merged.xlsx",
    "/view/hyperlink_target.xlsx",
    "/view/hyperlink_tooltip.xlsx",
    "/view/hyperlink_refused.xlsx",
]


@pytest.mark.parametrize("path", REQUESTS)
def test_a_page_of_serve_holds_no_markup_from_anyone(served, path):
    status, body = page_of(served, path)
    assert status in (200, 404, 422), (status, body[:200])
    assert_html_is_ours(body, extra_tags=SERVE_TAGS, extra_attributes=SERVE_ATTRIBUTES)


def test_serve_shows_a_payload_as_text(served):
    status, body = page_of(served, "/view/inline.xlsx")
    assert status == 200
    assert PAYLOAD in "".join(parse(body).text)
    status, body = page_of(served, "/")
    assert status == 200
    assert FILE_NAME in "".join(parse(body).text)
