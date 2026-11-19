"""Hyperlinks: what the reader takes from a file, what each output draws, and what it refuses.

The look of a link is the cell's format alone, and ``display`` is never drawn (captured,
screen suite ``links``).
"""

import re
from html.parser import HTMLParser
from xml.sax.saxutils import quoteattr

import openpyxl
import pytest

import sheetpeek
from sheetpeek import CellRange, Hyperlink
from sheetpeek._hyperlink import check_target, location_cell
from sheetpeek._read_links import literal_target
from sheetpeek.backends import HTMLBackend, MarkdownBackend
from tests.watchdog import bounded
from tests.workbooks import raw_sheet, replace_part, save

RELS = "xl/worksheets/_rels/sheet1.xml.rels"
HYPERLINK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


def linked_sheet(tmp_path, data, links, rels=(), merges="", columns=""):
    """A sheet named S with ``data`` (the rows), ``<hyperlink>`` elements and relationships.

    ``rels`` holds (id, target, external) triples.
    """
    body = f"{columns}<sheetData>{data}</sheetData>{merges}"
    if links:
        body += f"<hyperlinks>{links}</hyperlinks>"
    path = raw_sheet(tmp_path / "links.xlsx", body)
    if rels:
        entries = "".join(
            f'<Relationship Id="{rid}" Type="{HYPERLINK_REL}" Target={quoteattr(target)}'
            + (' TargetMode="External"' if external else "")
            + "/>"
            for rid, target, external in rels
        )
        replace_part(
            path,
            RELS,
            (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                f'relationships">{entries}</Relationships>'
            ).encode(),
        )
    return path


def text_cell(ref, text):
    return f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>'


def row(index, *cells):
    return f'<row r="{index}">{"".join(cells)}</row>'


def link_notices(sheet):
    return {n.code: n for n in sheet.notices if n.matches("link")}


# --- reading


def test_an_element_gives_target_location_tooltip_and_display(tmp_path):
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", "site"), text_cell("B1", "place"), text_cell("C1", "lost"))
        + row(2, text_cell("B2", "range")),
        '<hyperlink ref="A1" r:id="rId1" location="top" tooltip="tip" display="shown?"/>'
        '<hyperlink ref="B1:B2" location="\'S\'!A1"/>'
        '<hyperlink ref="C1" r:id="rId9"/>',  # an unknown relationship and no location
        [("rId1", "https://example.com/", True)],
    )
    sheet = sheetpeek.read(path)
    assert sheet.cell(1, 1).hyperlink == Hyperlink(
        "https://example.com/", "top", "tip", "shown?", "element"
    )
    assert sheet.cell(1, 2).hyperlink == sheet.cell(2, 2).hyperlink == Hyperlink(None, "'S'!A1")
    assert sheet.cell(1, 3).hyperlink is None
    assert link_notices(sheet) == {}


def test_a_range_is_cut_to_what_is_read(tmp_path):
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", "a"), text_cell("B1", "b"), text_cell("C1", "c")),
        '<hyperlink ref="A1:C1" location="A1"/>',
    )
    sheet = sheetpeek.read(path, ref="B1:C1")
    assert [c.hyperlink for c in sheet.cells[0]] == [Hyperlink(None, "A1")] * 2


def test_a_relationship_inside_the_package_is_refused(tmp_path):
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", "part")),
        '<hyperlink ref="A1" r:id="rId1"/>',
        [("rId1", "../media/x.bin", False)],
    )
    sheet = sheetpeek.read(path)
    assert sheet.cell(1, 1).hyperlink is None
    notice = link_notices(sheet)["link.scheme"]
    assert (notice.outcome, notice.reason, notice.context) == (
        "omitted",
        "unsafe",
        {"scheme": "part"},
    )


def formula_cell(ref, formula, value="shown", attributes=""):
    return f'<c r="{ref}" t="str"><f{attributes}>{formula}</f><v>{value}</v></c>'


def test_a_hyperlink_formula_with_a_literal_target_links(tmp_path):
    path = linked_sheet(
        tmp_path,
        row(1, formula_cell("A1", 'HYPERLINK("https://example.com/?q=&quot;&quot;x", "go")')),
        "",
    )
    # Excel recalculates a file it did not save: the friendly name; with the cached values
    # kept, the cached one.
    cell = sheetpeek.read(path).cell(1, 1)
    assert cell.value == "go"
    assert cell.hyperlink == Hyperlink('https://example.com/?q="x', source="formula")
    assert sheetpeek.read(path, recalc_on_load="never").cell(1, 1).value == "shown"


def test_a_hyperlink_formula_that_cannot_be_read_is_reported(tmp_path):
    path = linked_sheet(
        tmp_path,
        row(
            1,
            formula_cell("A1", 'HYPERLINK(B9&amp;"x")'),
            formula_cell(
                "B1", 'HYPERLINK("https://e.example/")', attributes=' t="shared" si="0" ref="B1:C1"'
            ),
            '<c r="C1" t="str"><f t="shared" si="0"/><v>shown</v></c>',
            formula_cell(
                "D1", 'HYPERLINK("https://e.example/")', attributes=' t="array" ref="D1:D2"'
            ),
            formula_cell("E1", 'IF(1,HYPERLINK("https://e.example/"))'),
            formula_cell("F1", '"HYPERLINK(" &amp; 1'),  # only in a string: not a call
        ),
        "",
    )
    sheet = sheetpeek.read(path)
    assert all(c.hyperlink is None for c in sheet.cells[0])
    found = {
        n.context["form"]: [str(c) for c in n.cells]
        for n in sheet.notices
        if n.code == "link.computed"
    }
    # The array's range is cut to what is read, as every range is.
    assert found == {"computed": ["A1", "E1"], "shared": ["B1:C1"], "array": ["D1"]}


def test_an_element_wins_over_a_formula_and_the_conflict_is_reported(tmp_path):
    path = linked_sheet(
        tmp_path,
        row(1, formula_cell("A1", 'HYPERLINK("https://formula.example/")')),
        '<hyperlink ref="A1" r:id="rId1"/>',
        [("rId1", "https://element.example/", True)],
    )
    sheet = sheetpeek.read(path)
    assert sheet.cell(1, 1).hyperlink == Hyperlink("https://element.example/")
    notice = link_notices(sheet)["link.conflict"]
    assert (notice.outcome, notice.reason, notice.count) == ("approximated", "file", 1)


def test_an_empty_linked_cell_is_reported_but_not_a_merges_covered_cells(tmp_path):
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", "merged"), text_cell("D1", "x")) + row(3, text_cell("A3", "x")),
        '<hyperlink ref="A1:C1" location="A3"/><hyperlink ref="E1" location="A3"/>'
        '<hyperlink ref="F1" location="A3"/>',
        merges='<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>',
    )
    sheet = sheetpeek.read(path, ref="A1:F3")
    notices = link_notices(sheet)
    assert list(notices) == ["link.empty_cell"]
    assert [str(c) for c in notices["link.empty_cell"].cells] == ["E1:F1"]
    # E1 has no <c>: the empty cell still carries its link.
    assert sheet.cell(1, 5).hyperlink is not None


@pytest.mark.parametrize(
    ("formula", "target"),
    [
        ('HYPERLINK("https://e.example/")', "https://e.example/"),
        (' hyperlink ( "a""b" , "c" ) ', 'a"b'),
        ('HYPERLINK("x(")', "x("),
        ('HYPERLINK("a")&"b"', None),
        ('HYPERLINK("a" & "b")', None),
        ("HYPERLINK(A1)", None),
        ('IF(1,HYPERLINK("a"))', None),
        ('MYHYPERLINK("a")', None),
    ],
)
def test_literal_target(formula, target):
    assert literal_target(formula) == target


# --- the scheme check, shared by every output


@pytest.mark.parametrize(
    ("target", "url", "scheme"),
    [
        ("https://example.com/a b", "https://example.com/a b", "https"),
        ("HTTP://EXAMPLE.COM/", "HTTP://EXAMPLE.COM/", "http"),
        ("mailto:someone@example.com", "mailto:someone@example.com", "mailto"),
        (" \x01https://example.com/\n", "https://example.com/", "https"),
        ("https://exam\tple.com/\r\n", "https://example.com/", "https"),
        ("https://e.example/\x9b31m", "https://e.example/%C2%9B31m", "https"),
        ("javascript:alert(1)", None, "javascript"),
        ("java\tscript:alert(1)", None, "javascript"),
        ("java\nscript:alert(1)", None, "javascript"),
        ("  JavaScript:alert(1)", None, "javascript"),
        ("data:text/html,x", None, "data"),
        ("vbscript:x", None, "vbscript"),
        ("file:///etc/passwd", None, "file"),
        ("\\\\server\\share\\x.xlsx", None, "relative"),
        ("..\\budget.xls", None, "relative"),
        ("#Sheet2!A1", None, "relative"),
        ("C:\\x.xlsx", None, "path"),
        ("http://[::1", None, "unparseable"),
    ],
)
def test_only_allowed_schemes_are_linked(target, url, scheme):
    assert check_target(target) == (url, scheme)


@pytest.mark.parametrize(
    ("location", "cell"),
    [
        ("A1", (1, 1)),
        ("$B$2", (2, 2)),
        ("S!C3:D4", (3, 3)),
        ("s!C3", (3, 3)),
        ("'S'!A1", (1, 1)),
        ("'It''s'!A1", None),
        ("Other!A1", None),
        ("Total", None),
        ("S!Total", None),
        ("'S", None),
        ("", None),
    ],
)
def test_location_cell(location, cell):
    assert location_cell(location, "S") == cell


def test_a_quoted_sheet_name_with_a_quote():
    assert location_cell("'It''s'!B2", "It's") == (2, 2)


# --- HTML


class Anchors(HTMLParser):
    """Each anchor's attributes and the text inside it; and each element's id."""

    def __init__(self):
        super().__init__()
        self.anchors = []
        self.ids = {}
        self.depth = 0

    def handle_starttag(self, tag, attrs):
        attributes = dict(attrs)
        if "id" in attributes:
            self.ids[attributes["id"]] = tag
        if tag == "a":
            self.anchors.append([attributes, ""])
            self.depth = 1
        elif self.depth:
            self.depth += 1

    def handle_endtag(self, tag):
        if tag == "a":
            self.depth = 0
        elif self.depth:
            self.depth -= 1

    def handle_data(self, data):
        if self.depth:
            self.anchors[-1][1] += data


def anchors(markup):
    parser = Anchors()
    parser.feed(markup)
    return parser


def html_of(path, backend=None, **read):
    return sheetpeek.render(
        sheetpeek.read(path, **read), backend or HTMLBackend(), on_notice="ignore"
    )


def test_an_anchor_wraps_the_cells_text_and_adds_no_look(tmp_path):
    long = "a long text that spills over its neighbors"
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", long)),
        '<hyperlink ref="A1" r:id="rId1" tooltip="&quot;tip&quot; &lt;b&gt;" display="other"/>',
        [("rId1", "https://example.com/?a=1&b=<2>", True)],
    )
    page = html_of(path)
    found = anchors(page)
    assert found.anchors == [
        [
            {
                "href": "https://example.com/?a=1&b=<2>",
                "rel": "noopener noreferrer",
                "title": '"tip" <b>',
            },
            long,
        ]
    ]
    # display is never drawn, and the anchor adds no look: the cell's format draws a link
    # (captured, links).
    assert "other" not in page
    # The anchor holds the run whole, clip window and all, which spills over the neighbors.
    assert re.search(
        r'<a [^>]*>(<span class="sp-clip"[^>]*>)?<span class="sp-run"[^>]*>a long', page
    )
    assert '</a></td><td class="sp-crossed">' in page
    assert ".sheetpeek a { color: inherit; text-decoration: inherit; }" in page


def test_new_tab_is_an_option_for_external_links_only(tmp_path):
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", "site"), text_cell("B1", "here")),
        '<hyperlink ref="A1" r:id="rId1"/><hyperlink ref="B1" location="A1"/>',
        [("rId1", "https://example.com/", True)],
    )
    found = anchors(html_of(path, HTMLBackend(new_tab=True)))
    assert [a.get("target") for a, _ in found.anchors] == ["_blank", None]
    assert all("target" not in a for a, _ in anchors(html_of(path)).anchors)


def test_a_link_to_a_shown_cell_goes_to_its_id(tmp_path):
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", "to C2"), text_cell("B1", "to merged"), text_cell("D1", "to hidden"))
        + row(2, text_cell("C2", "target"), text_cell("D2", "merged"), text_cell("F2", "hidden")),
        '<hyperlink ref="A1" location="S!C2"/><hyperlink ref="B1" location="E2"/>'
        '<hyperlink ref="D1" location="F2"/>',
        merges='<mergeCells count="1"><mergeCell ref="D2:E2"/></mergeCells>',
        columns='<cols><col min="6" max="6" hidden="1"/></cols>',
    )
    sheet = sheetpeek.read(path)
    found = anchors(sheetpeek.render(sheet, "html", on_notice="ignore"))
    hrefs = {text: a.get("href") for a, text in found.anchors}
    assert hrefs == {"to C2": "#sp-C2", "to merged": "#sp-D2", "to hidden": None}
    assert found.ids == {"sp-C2": "td", "sp-D2": "td"}
    # Shown with the hidden column, the link goes to it.
    found = anchors(
        sheetpeek.render(sheet.layout("excel", show_hidden=True), "html", on_notice="ignore")
    )
    assert {text: a.get("href") for a, text in found.anchors}["to hidden"] == "#sp-F2"


def test_a_link_elsewhere_is_shown_but_goes_nowhere(tmp_path):
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", "other sheet"), text_cell("B1", "a name")),
        '<hyperlink ref="A1" location="Other!A1" tooltip="tip"/>'
        '<hyperlink ref="B1" location="Total"/>',
    )
    sheet = sheetpeek.read(path)
    found = anchors(sheetpeek.render(sheet, "html", on_notice="ignore"))
    assert found.anchors == [[{"title": "tip"}, "other sheet"], [{}, "a name"]]
    (notice,) = [n for n in sheetpeek.notices(sheet, "html") if n.code == "link.elsewhere"]
    assert (notice.outcome, notice.count, notice.output) == ("approximated", 2, "html")


def test_a_refused_target_is_text_and_never_echoed(tmp_path):
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", "click"), text_cell("B1", "tab")),
        '<hyperlink ref="A1" r:id="rId1" tooltip="tip"/><hyperlink ref="B1" r:id="rId2"/>',
        [("rId1", "javascript:alert('PWN')", True), ("rId2", "java\tscript:alert('PWN')", True)],
    )
    sheet = sheetpeek.read(path)
    page = sheetpeek.render(sheet, "html", on_notice="ignore")
    assert "<a" not in page
    assert "PWN" not in page
    assert "tip" not in page
    assert "click" in page
    (notice,) = [n for n in sheetpeek.notices(sheet, "html") if n.code == "link.scheme"]
    assert notice.output is None  # decided once, for every output
    assert notice.count == 2
    assert notice.context == {"scheme": "javascript"}
    for to in ("markdown", "text", "csv"):
        assert [n.code for n in sheetpeek.notices(sheet, to) if n.matches("link")] == [
            "link.scheme"
        ]


def test_an_empty_linked_cell_gets_no_anchor(tmp_path):
    # Excel shows nothing in it (captured, links).
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", "x")),
        '<hyperlink ref="B1" r:id="rId1"/>',
        [("rId1", "https://example.com/", True)],
    )
    assert "<a" not in html_of(path, ref="A1:B1")


# --- the other outputs


def test_markdown_writes_a_link_with_its_destination_encoded(tmp_path):
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", "head"))
        + row(2, text_cell("A2", "a [b] | c"))
        + row(3, text_cell("A3", "here")),
        '<hyperlink ref="A2" r:id="rId1"/><hyperlink ref="A3" location="A1"/>',
        [("rId1", "https://example.com/a b(c)|d?x=1&y=2", True)],
    )
    out = sheetpeek.render(sheetpeek.read(path), "markdown", on_notice="ignore")
    assert out.splitlines()[2:] == [
        r"| [a \[b\] \| c](https://example.com/a%20b%28c%29%7Cd?x=1&amp;y=2) |",
        "| here |",
    ]


def test_markdown_puts_the_link_inside_emphasis(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "head"
    wb.active["A2"] = "bold"
    wb.active["A2"].hyperlink = "https://example.com/"
    wb.active["A2"].font = openpyxl.styles.Font(bold=True)
    path = save(wb, tmp_path / "bold.xlsx")
    out = sheetpeek.render(path, sheetpeek.backends.MarkdownBackend(emphasis=True))
    assert out.splitlines()[2] == "| **[bold](https://example.com/)** |"


@pytest.mark.parametrize("to", ["text", "csv", "tsv"])
def test_text_and_delimited_keep_the_text_without_the_target(tmp_path, to):
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", "site")),
        '<hyperlink ref="A1" r:id="rId1"/>',
        [("rId1", "https://example.com/", True)],
    )
    sheet = sheetpeek.read(path)
    out = sheetpeek.render(sheet, to, on_notice="ignore")
    assert "site" in out
    assert "example" not in out
    (notice,) = [n for n in sheetpeek.notices(sheet, to) if n.code == "link"]
    assert notice.outcome == "omitted"


def test_rich_draws_a_terminal_hyperlink_over_the_text_only(tmp_path):
    pytest.importorskip("rich")
    path = linked_sheet(
        tmp_path,
        row(1, text_cell("A1", "site"), text_cell("B1", "next")),
        '<hyperlink ref="A1" r:id="rId1"/>',
        [("rId1", "https://example.com/\u009b", True)],
    )
    backend = sheetpeek.backends.RichBackend(force_terminal=True, color_system="truecolor")
    out = sheetpeek.render(sheetpeek.read(path), backend, on_notice="ignore")
    linked = re.findall(r"\x1b\]8;id=[^;]*;([^\x1b]*)\x1b\\(.*?)\x1b\]8;;\x1b\\", out)
    assert linked == [("https://example.com/%C2%9B", "site")]
    assert "\u009b" not in out


# --- crafted input


@bounded
def test_a_huge_link_range_costs_nothing_past_the_range_read(tmp_path):
    links = '<hyperlink ref="A1:XFD1048576" location="A1"/>' * 1000
    path = linked_sheet(tmp_path, row(1, text_cell("A1", "a"), text_cell("B1", "b")), links)
    sheet = sheetpeek.read(path)
    assert sheet.ref == CellRange.parse("A1:B1")
    assert sheet.cell(1, 2).hyperlink == Hyperlink(None, "A1")
    for mode, to in (("excel", "html"), ("characters", "text"), ("none", "markdown")):
        sheetpeek.render(sheet.layout(mode), to, on_notice="ignore")


LOCAL = ("file:///C:/data/budget.xlsx", "..\\budget.xlsx", "C:\\data\\budget.xlsx", "budget.xlsx")


@pytest.mark.parametrize("target", LOCAL)
def test_local_targets_are_linked_only_when_asked(tmp_path, target):
    # Excel follows file: and relative targets; sheetpeek links them only with local_links.
    wb = openpyxl.Workbook()
    wb.active["A1"] = "local"
    wb.active["A1"].hyperlink = target
    wb.active["A2"] = "script"
    wb.active["A2"].hyperlink = "javascript:alert(1)"
    path = save(wb, tmp_path / "local.xlsx")
    plain = sheetpeek.render(path, "html", on_notice="ignore")
    assert "href" not in plain
    found = {n.code: n for n in sheetpeek.notices(path, "html")}
    assert (found["link.local"].outcome, found["link.local"].reason) == ("omitted", "option")
    assert found["link.scheme"].count == 1  # the javascript: target only
    linked = sheetpeek.render(path, HTMLBackend(local_links=True), on_notice="ignore")
    assert linked.count("href=") == 1
    assert "javascript" not in linked
    assert "link.local" not in {
        n.code for n in sheetpeek.notices(path, HTMLBackend(local_links=True))
    }
    markdown = sheetpeek.render(path, MarkdownBackend(local_links=True), on_notice="ignore")
    assert "](" in markdown
    assert "](" not in sheetpeek.render(path, "markdown", on_notice="ignore")
