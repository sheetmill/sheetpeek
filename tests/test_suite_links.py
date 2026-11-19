"""The screen suite ``links`` (hyperlinks, ARCHITECTURE.md §4.10), without Excel.

Its workbook is built and read back with sheetpeek as built, and what sheetpeek draws for each
case is the suite's ``EXPECTED`` (``FOLLOW_EXPECTED`` for where a link goes). Its measuring runs
on a synthetic capture with known marks, and its reading of the follow step on a synthetic
``formats.tsv``. The comparison with ``tests/captures/links.tsv`` is skipped until that capture
exists.
"""

import zipfile

import openpyxl
import pytest
from capture import links
from capture.screen.window import MARKER
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from sheetpeek._hyperlink import location_place
from tests.screen_canvas import SCALE, Canvas

GREEN = (0, 128, 0)


def _rgb(text):
    return tuple(int(text[k : k + 2], 16) for k in (0, 2, 4))


def _distance(a, b):
    return sum((x - y) ** 2 for x, y in zip(_rgb(a), _rgb(b), strict=True)) ** 0.5


def _cell(sheet, address):
    letters, row = coordinate_from_string(address)
    return sheet.cell(row, column_index_from_string(letters))


def _case(name):
    return next(c for c in links.CASES if c.name == name)


# --- the workbook


def test_the_links_workbook_reads_back_as_built(tmp_path):
    path = tmp_path / "links.xlsx"
    assert links.build(path) == links.CASES
    assert len({c.name for c in links.CASES}) == len(links.CASES)
    assert set(links.EXPECTED) == {c.name for c in links.CASES}
    sheet = sheetpeek.read(path)
    assert sheet.name == "Links"
    drawn = {}
    for case in links.CASES:
        cell = _cell(sheet, case.cell)
        font = cell.style.font
        text = "" if cell.value is None else str(cell.value)
        drawn[case.name] = (text, font.color or "000000", font.underline or "none")
        if "no link" in case.name:
            assert cell.hyperlink is None, case.name
        else:
            assert cell.hyperlink is not None, case.name
    assert drawn == links.EXPECTED
    assert _cell(sheet, "C8").hyperlink.display == "display text much longer than the value"
    assert _cell(sheet, "G8").hyperlink.display == "ghost"
    assert _cell(sheet, "G10").hyperlink.source == "element"  # the element wins
    codes = {n.code for n in sheetpeek.notices(sheet, "html")}
    assert {"link.conflict", "link.empty_cell"} <= codes


def test_each_link_goes_where_follow_expected_says(tmp_path):
    path = tmp_path / "links.xlsx"
    links.build(path)
    assert [c.name for c in links.FOLLOWS] == list(links.FOLLOW_EXPECTED)
    for case in links.FOLLOWS:
        link = _cell(sheetpeek.read(path, case.sheet), case.cell).hyperlink
        place = None
        if link.location is not None:
            name, row, column = location_place(link.location, case.sheet)
            place = f"{name}!{openpyxl.utils.get_column_letter(column)}{row}"
        assert place == links.FOLLOW_EXPECTED[case.name], case.name


def test_the_theme_and_the_empty_links_are_written_as_stated(tmp_path):
    path = tmp_path / "links.xlsx"
    links.build(path)
    with zipfile.ZipFile(path) as package:
        theme = package.read("xl/theme/theme1.xml").decode()
        sheet = package.read("xl/worksheets/sheet1.xml").decode()
        styles = package.read("xl/styles.xml").decode()
        book = package.read("xl/workbook.xml").decode()
    assert f'<a:srgbClr val="{links.HLINK}"/>' in theme
    assert f'<a:srgbClr val="{links.FOLHLINK}"/>' in theme
    assert links.EMPTY_LINKS in sheet
    assert '<c r="G8"' not in sheet  # the empty cells hold nothing
    assert "<v />" in sheet  # formulas carry no cached value
    assert 'fullCalcOnLoad="1"' in book
    assert 'builtinId="8"' in styles
    assert 'builtinId="9"' in styles
    assert '<color theme="10" /><sz val="11" /><u val="single" />' in styles
    with pytest.raises(ValueError, match="hlink"):
        links.theme("<a:clrScheme/>")
    with pytest.raises(ValueError, match="hyperlinks"):
        links.sheet("<worksheet/>")


def test_every_case_has_a_cell_of_its_own_and_the_markers_lie_beyond():
    cells = [coordinate_from_string(c) for c in links._cells()]
    assert len(set(cells)) == len(cells)
    letters, row = coordinate_from_string(links.MARKERS[1])
    assert row > max(r for _, r in cells)
    assert column_index_from_string(letters) > max(column_index_from_string(c) for c, _ in cells)
    for case in links.CASES:
        assert coordinate_from_string(case.cell)[0] in links.SLOTS


def test_no_case_draws_near_the_marker_color():
    for color in (links.HLINK, links.FOLHLINK, "000000", "FFFFFF"):
        assert _distance(color, MARKER) > 200


# --- measuring, on a synthetic capture


def _canvas():
    ws = openpyxl.Workbook().active
    links.fill(ws)
    return Canvas(ws, links._cells())


def test_an_underlined_text_reads_as_its_ink_and_one_rule():
    canvas = _canvas()
    case = _case("Hyperlink style, linked")
    x0, y0, _, y1 = canvas.box(case.cell)
    for x in range(x0 + 6, x0 + 80, 8):  # the glyphs' stems: 7 pt high
        canvas.paint(x, y1 - 20, x + 2, y1 - 6, GREEN)
    canvas.paint(x0 + 6, y1 - 4, x0 + 80, y1 - 2, GREEN)  # a 1 pt rule 1 pt below
    measured = links.measure(canvas.capture(), case)
    assert measured["ink_color"] == "008000"
    assert (measured["ink_left"], measured["ink_right"]) == ("3", "40")
    assert (measured["ink_top"], measured["ink_bottom"]) == (
        f"{(y1 - 20 - y0) / SCALE:g}",
        f"{(y1 - 2 - y0) / SCALE:g}",
    )
    assert measured["rules"] == "1"
    assert (measured["rule_top"], measured["rule_thickness"]) == (
        f"{(y1 - 4 - y0) / SCALE:g}",
        "1",
    )


def test_text_without_a_rule_reads_no_rule():
    canvas = _canvas()
    case = _case("Normal, linked")
    x0, _, _, y1 = canvas.box(case.cell)
    for x in range(x0 + 6, x0 + 80, 8):
        canvas.paint(x, y1 - 20, x + 2, y1 - 6, (0, 0, 0))
    measured = links.measure(canvas.capture(), case)
    assert (measured["ink_color"], measured["rules"]) == ("000000", "0")
    assert "rule_top" not in measured


def test_gridlines_are_neither_ink_nor_a_rule():
    canvas = _canvas()
    case = _case("Normal, linked")
    x0, y0, x1, y1 = canvas.box(case.cell)
    canvas.paint(x0 - 4, y0 - 1, x1 + 4, y0 + 1, (212, 212, 212))
    canvas.paint(x0 - 4, y1 - 1, x1 + 4, y1 + 1, (212, 212, 212))
    assert links.measure(canvas.capture(), case)["ink_pixels"] == "0"


def test_a_merge_is_measured_across_its_cells():
    canvas = _canvas()
    case = _case("empty merge, linked with a display")
    _, y0, x1, y1 = canvas.box(case.last)
    canvas.paint(x1 - 30, y0 + 8, x1 - 20, y1 - 8, (0, 0, 0))  # ink in the merge's last cell
    measured = links.measure(canvas.capture(), case)
    assert int(measured["ink_pixels"]) == 10 * (y1 - y0 - 16)
    left = canvas.geometry[case.cell]["left"]
    assert float(measured["ink_left"]) == pytest.approx((x1 - 30) / SCALE - left)


def test_every_case_is_measured_on_a_blank_capture():
    cap = _canvas().capture()
    for case in links.CASES:
        assert links.measure(cap, case) == {"ink_pixels": "0", "ink_color": "none", "rules": "0"}


def test_the_follow_reads_become_one_row_a_case(tmp_path, monkeypatch):
    monkeypatch.setattr(links, "BUILD_DIR", tmp_path / "build")
    monkeypatch.setattr(links, "OUTPUT", tmp_path / "links.tsv")
    out = tmp_path / "build" / "follow"
    out.mkdir(parents=True)
    case = links.FOLLOWS[1]
    reads = [
        ("active.before", "Second"), ("links.count", "1"), ("formula", "to B5"),
        ("link.address", ""), ("link.sub_address", "B5"), ("followed", "yes"),
        ("after.sheet", "Second"), ("after.cell", "$B$5"),
    ]  # fmt: skip
    lines = ["# excel: 16.115", "# suite: links_follow", "cell\tproperty\tvalue"]
    lines += [f"{case.name}\t{key}\t{value}" for key, value in reads]
    (out / "formats.tsv").write_text("\n".join(lines) + "\n", encoding="utf-8")
    links.write()
    text = (tmp_path / "links.tsv").read_text(encoding="utf-8").splitlines()
    assert text[:3] == ["# excel: 16.115", "# suite: links", "\t".join(links.COLUMNS)]
    rows = [dict(zip(links.COLUMNS, line.split("\t"), strict=True)) for line in text[3:]]
    assert [r["name"] for r in rows] == [f.name for f in links.FOLLOWS]
    found = rows[1]
    assert (found["kind"], found["sheet"], found["cell"]) == ("follow", "Second", "B2")
    assert (found["active"], found["links_count"], found["link_sub_address"]) == (
        "Second",
        "1",
        "B5",
    )
    assert (found["followed"], found["after_sheet"], found["after_cell"]) == (
        "yes",
        "Second",
        "$B$5",
    )
    assert rows[0]["followed"] == ""  # a case Excel was not asked about stays empty


# --- the capture (suite links)


def _captured(kind):
    path = links.OUTPUT
    if not path.exists():
        pytest.skip("not captured yet: run python -m capture.links run")
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    found = {}
    for row in body:
        record = dict(zip(header, row, strict=False))
        if record["kind"] == kind and (kind == "follow" or record["zoom"] == str(links.ZOOM)):
            found[record["name"]] = record
    if not found:
        pytest.skip(f"no {kind} case captured")
    return found


# Tolerances. A font's ink is anti-aliased; the tenth of its pixels farthest from the cell's
# color is the stems' core, whole device pixels at 2 a point: 40 takes the color profile and
# the hinting (the auto_colors suite's figure) and still tells black, green 008000 and orange
# C05000 apart (more than 120 from each other). Excel's display format reads its colors as
# integers, so they compare exactly. "v" in Calibri 11 is about 5 pt wide and the display text
# about 190 pt: an ink 12 pt wide or less is the value, whatever the hinting.
INK = 40
ONE_GLYPH = 12.0


def _read_color(text):
    return "".join(f"{int(part):02X}" for part in text.split(","))


@pytest.mark.excel
def test_a_link_has_no_look_of_its_own():
    for name, record in _captured("look").items():
        text, color, underline = links.EXPECTED[name]
        assert record["text"] == text, (name, record["text"])
        if not text:
            # an empty linked cell shows nothing: display is not drawn
            assert record["ink_pixels"] == "0", (name, record["ink_pixels"])
            continue
        assert record["ink_color"] != "none", (name, "no text drawn")
        assert _distance(record["ink_color"], color) <= INK, (name, record["ink_color"])
        assert _read_color(record["display_font_color"]) == color, name
        assert record["display_underline"] == f"underline style {underline}", name
        if len(text) == 1:
            # the rule detector takes a lone glyph's strokes for rules ("v": one rule at 9.5 pt
            # in C8, which Excel reads without underline): Excel's read above decides
            continue
        assert record["rules"] == ("1" if underline == "single" else "0"), (name, record["rules"])
        if underline == "single":
            assert float(record["rule_top"]) >= float(record["ink_top"]), name


@pytest.mark.excel
def test_display_is_not_drawn():
    for name, record in _captured("look").items():
        if links.EXPECTED[name][0] == "v":
            width = float(record["ink_right"]) - float(record["ink_left"])
            assert width <= ONE_GLYPH, (name, width)


@pytest.mark.excel
def test_links_go_where_sheetpeek_takes_them():
    for name, record in _captured("follow").items():
        expected = links.FOLLOW_EXPECTED[name]
        if expected is None:
            continue  # sheetpeek takes it to no place in the workbook; the record informs
        assert record["followed"] == "yes", (name, record["followed"])
        place = f"{record['after_sheet']}!{record['after_cell'].replace('$', '')}"
        assert place == expected, (name, place)


def _agreeing():
    lines = ["# suite: links", "\t".join(links.COLUMNS)]
    for case in links.CASES:
        text, color, underline = links.EXPECTED[case.name]
        record = {"kind": "look", "name": case.name, "cell": case.cell, "zoom": "100"}
        record["text"] = text
        if text:
            record |= {
                "ink_pixels": "300", "ink_color": color, "ink_left": "2",
                "ink_right": "6" if text == "v" else "70", "ink_top": "4", "ink_bottom": "12",
                "rules": "1" if underline == "single" else "0",
                "display_font_color": ",".join(str(c) for c in _rgb(color)),
                "display_underline": f"underline style {underline}",
            }  # fmt: skip
            if underline == "single":
                record |= {"rule_top": "13", "rule_thickness": "1"}
        else:
            record |= {"ink_pixels": "0", "ink_color": "none", "rules": "0"}
        lines.append("\t".join(record.get(c, "") for c in links.COLUMNS))
    for case in links.FOLLOWS:
        place = links.FOLLOW_EXPECTED[case.name] or "Third!B7"
        sheet, _, cell = place.partition("!")
        record = {"kind": "follow", "name": case.name, "cell": case.cell, "sheet": case.sheet}
        record |= {"followed": "yes", "after_sheet": sheet, "after_cell": f"${cell[0]}${cell[1:]}"}
        lines.append("\t".join(record.get(c, "") for c in links.COLUMNS))
    return lines


def test_the_comparison_accepts_a_capture_that_agrees_with_sheetpeek(tmp_path, monkeypatch):
    path = tmp_path / "links.tsv"
    monkeypatch.setattr(links, "OUTPUT", path)
    lines = _agreeing()
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    test_a_link_has_no_look_of_its_own()
    test_display_is_not_drawn()
    test_links_go_where_sheetpeek_takes_them()
    # and refuses one that does not: a linked Normal cell drawn in the link color, display
    # drawn, a location without a sheet taken on the active sheet
    mutations = {
        "Normal, linked\tC4": ("\t000000\t", "\t008000\t", test_a_link_has_no_look_of_its_own),
        "Normal, display unlike the value": ("\t6\t", "\t190\t", test_display_is_not_drawn),
        "the first sheet active": (
            "\tSecond\t$B$5",
            "\tLinks\t$B$5",
            test_links_go_where_sheetpeek_takes_them,
        ),
    }
    for name, (old, new, check) in mutations.items():
        changed = [line.replace(old, new) if name in line else line for line in lines]
        assert changed != lines, name
        path.write_text("\n".join(changed) + "\n", encoding="utf-8")
        with pytest.raises(AssertionError):
            check()
