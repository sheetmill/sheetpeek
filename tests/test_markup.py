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
from xml.sax.saxutils import escape, quoteattr

import openpyxl
import pytest
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.colors import Color

import sheetpeek
from sheetpeek import PackageError
from sheetpeek._serve import make_server
from sheetpeek.backends import HTMLBackend, MarkdownBackend
from tests.workbooks import (
    NS,
    SHEET1,
    replace_once,
    replace_part,
    replace_sheet_data,
    rewrite,
    save,
)

MARK = "PWN"
PAYLOAD = (
    f"\"'></title></style></textarea><{MARK} ONERROR=x>&lt;{MARK}&gt;</td><script>x</script>"
    f"[l](javascript:{MARK})![i](x:{MARK})<http://{MARK}.example>`c`*e*_f_~~s~~&amp;"
    f"'; background:url({MARK}); x:'\\"
)
HTML_TAGS = {
    "html", "head", "meta", "title", "style", "body", "div", "table", "colgroup", "col", "thead",
    "tbody", "tr", "td", "th", "span", "details", "summary", "ul", "li",
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
}
STYLE_PROPERTIES = {
    "width", "height", "left", "right", "top", "font", "color", "background", "border-color",
    "margin-top", "border-left-color", "border-top-color", "letter-spacing", "text-decoration",
}  # fmt: skip
MARKDOWN_TAGS = {"table", "thead", "tbody", "tr", "th", "td", "br", "p", "strong", "em"}


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
            assert value.startswith("/"), value
        elif name not in (
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
    return cell_xml(tmp_path, f'<c r="A2" t="str"><f>1</f><v>{escape(PAYLOAD)}</v></c>')


def error_text(tmp_path):
    return cell_xml(tmp_path, f'<c r="A2" t="e"><v>{escape(PAYLOAD)}</v></c>')


def error_shaped(tmp_path):
    return cell_xml(tmp_path, f'<c r="A2" t="e"><v>#{MARK}&lt;{MARK}&gt;!</v></c>')


def date_text(tmp_path):
    return cell_xml(tmp_path, f'<c r="A2" t="d"><v>{escape(PAYLOAD)}</v></c>')


def shared(tmp_path):
    path = cell_xml(tmp_path, '<c r="A2" t="s"><v>0</v></c>')
    rewrite(
        path,
        "[Content_Types].xml",
        replace_once(
            "</Types>",
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.'
            'openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
        ),
    )
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
        path, "xl/styles.xml", lambda t: t.replace('val="Calibri"', f"val={quoteattr(PAYLOAD)}")
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


def properties(tmp_path):
    path = xml_text(tmp_path, A2="x")
    return rewrite(path, "docProps/core.xml", replace_once("openpyxl", escape(PAYLOAD)))


def defined_name(tmp_path):
    path = xml_text(tmp_path, A2="x")
    name = f'<definedNames><definedName name="n">{escape(PAYLOAD)}</definedName></definedNames>'
    return rewrite(path, "xl/workbook.xml", replace_once("</sheets>", "</sheets>" + name))


SOURCES = [
    inline, cached, error_text, error_shaped, date_text, shared, escaped_in_file, line_breaks,
    wrapped, filled_alignment, justified, spilled_left, merged, hidden, format_quoted,
    format_text_section, format_escaped, format_currency, format_refused, sheet_name, font_name,
    normal_font_name, theme_srgb, theme_system, style_rgb, palette, hyperlink, properties,
    defined_name,
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
    "refused": {"date_text", "style_rgb", "palette"},
    "absent": {
        "hidden", "format_refused", "theme_srgb", "theme_system", "hyperlink", "properties",
        "defined_name",
    },
    "attribute": {"normal_font_name"},
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
        name == "style" and re.fullmatch(r"text-align:\w+", value)
        for name, value in found.attributes
    )
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
    for build in (inline, sheet_name, font_name, format_quoted, theme_srgb, palette, merged):
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
