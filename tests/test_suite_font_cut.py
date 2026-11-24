"""The screen suite ``font_cut`` (where Excel stops honoring a workbook's fonts), without Excel.

Its workbooks are built and their font tables read back from the files; its ink measure runs on
a synthetic capture with known marks.
"""

import re
import zipfile

import openpyxl
import pytest
from capture import font_cut as suite

from tests.screen_canvas import Canvas


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    return suite.build(tmp_path_factory.mktemp("font_cut"))


def _parts(path):
    with zipfile.ZipFile(path) as z:
        return z.read("xl/styles.xml").decode(), z.read("xl/worksheets/sheet1.xml").decode()


def test_each_list_probe_is_the_font_table_entry_its_case_names(built):
    styles, sheet = _parts(built["list"])
    fonts = re.search(r"<fonts[^>]*>(.*?)</fonts>", styles, re.S).group(1)
    assert len(re.findall(r"<font[ >]", fonts)) == suite.LIST_FIRST + suite.LIST_PROBES
    xfs = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", styles, re.S).group(1)
    font_of = [int(f) for f in re.findall(r'<xf [^>]*fontId="(\d+)"', xfs)]
    for cell, entry in suite.probes("list"):
        style = re.search(rf'<c r="{cell}"[^>]* s="(\d+)"', sheet)
        assert font_of[int(style.group(1))] == entry, cell


def test_the_rich_runs_follow_the_list_in_cell_order(built):
    styles, sheet = _parts(built["rich"])
    fonts = re.search(r"<fonts[^>]*>(.*?)</fonts>", styles, re.S).group(1)
    assert len(re.findall(r"<font[ >]", fonts)) == suite.RICH_FIRST
    runs = re.findall(r'<c r="(C\d+)" t="inlineStr"><is><r><rPr>', sheet)
    assert runs == [cell for cell, _ in suite.probes("rich")]
    assert [entry for _, entry in suite.probes("rich")] == list(range(995, 1005))


def test_the_ink_measure_tells_the_probes_red_from_the_default_fonts_black():
    ws = openpyxl.Workbook().active
    suite.fill(ws, "list")
    canvas = Canvas(ws, ["C4", "C5"])
    x0, y0, _, _ = canvas.box("C4")
    canvas.paint(x0 + 6, y0 + 8, x0 + 20, y0 + 30, (230, 0, 0))
    found = suite._ink(canvas.capture(), "C4")
    assert (found["red"], found["dark"]) == (str(14 * 22), "0")
    assert (found["ink_top"], found["ink_bottom"]) == ("4", "15")
    x0, y0, _, _ = canvas.box("C5")
    canvas.paint(x0 + 6, y0 + 10, x0 + 12, y0 + 24, (0, 0, 0))
    found = suite._ink(canvas.capture(), "C5")
    assert (found["red"], found["dark"]) == ("0", str(6 * 14))


def test_the_suite_is_a_screen_suite():
    from capture.screen.__main__ import SUITES

    assert "font_cut" in SUITES


# --- the comparison with Excel's capture


@pytest.mark.excel
def test_excel_for_mac_cuts_the_font_table_at_entry_1000():
    # Captured, Excel 16.115 for Mac, Calibri 11 Normal font: entries 990 to 999 of the font list,
    # and rich-text runs 995 to 999 after it, keep their Arial 24 (rows fitted at 30 pt, red ink);
    # from entry 1,000 the row fits at 16 pt and the ink is "Hxg" in Calibri 12, Excel for Mac's
    # standard size; a cut list font is drawn black, a cut run in its own red. Excel's object
    # model reports each list probe's own font, cut or not.
    if not suite.OUTPUT.exists():
        pytest.skip("not captured yet: run python -m capture.screen run font_cut")
    lines = suite.OUTPUT.read_text(encoding="utf-8").split("\n")
    rows = [line.split("\t") for line in lines if line and not line.startswith("#")]
    records = [dict(zip(rows[0], row, strict=True)) for row in rows[1:]]
    records = [r for r in records if r["group"] in ("list", "rich")]
    assert len(records) == suite.LIST_PROBES + suite.RICH_PROBES
    for r in records:
        cut = int(r["entry"]) >= 1000
        assert float(r["row_height"]) == (16.0 if cut else 30.0), r
        assert (r["ink_top"], r["ink_bottom"]) == (("5", "15") if cut else ("7.5", "30")), r
        if r["group"] == "list":
            assert (r["font_name"], r["font_size"]) == ("Arial", "24"), r
            assert (int(r["red"]) > 0, int(r["dark"]) > 0) == (
                (False, True) if cut else (True, False)
            )
        else:
            assert int(r["red"]) > 0, r


def test_the_theme_groups_rewrite_the_theme_and_set_the_normal_font(built):
    for group, (name, size) in suite.NORMALS.items():
        with zipfile.ZipFile(built[group]) as z:
            theme = z.read("xl/theme/theme1.xml").decode()
            styles = z.read("xl/styles.xml").decode()
        assert '<a:minorFont>\n        <a:latin typeface="Arial"/>' in theme
        assert '<a:majorFont>\n        <a:latin typeface="Aptos Narrow"/>' in theme
        first = re.search(r"<fonts[^>]*>\s*<font>(.*?)</font>", styles, re.S).group(1)
        assert f'<name val="{name}"' in first
        assert f'<sz val="{size}"' in first
        fonts = re.search(r"<fonts[^>]*>(.*?)</fonts>", styles, re.S).group(1)
        assert len(re.findall(r"<font[ >]", fonts)) == suite.THEME_FIRST + suite.THEME_PROBES


@pytest.mark.excel
def test_excel_for_mac_draws_a_cut_cell_in_calibri_12_whatever_the_theme_and_normal_font():
    # Captured, Excel 16.115 for Mac (standard font Aptos Narrow 12): under a theme of Arial body
    # and Aptos Narrow heading fonts and a Times New Roman 10 or 16 Normal font, a cut probe's
    # glyphs start where Calibri 12's advances put them (7 and 5.5 points apart, as under Calibri
    # 11 and a Calibri theme), not Arial's (9, 6) or Aptos Narrow's (8, 5); the row fits the
    # larger of Calibri 12's 16 pt and the Normal font's row.
    if not suite.OUTPUT.exists():
        pytest.skip("not captured yet")
    lines = suite.OUTPUT.read_text(encoding="utf-8").split("\n")
    assert "# standard_font: Aptos Narrow 12.0" in lines
    rows = [line.split("\t") for line in lines if line and not line.startswith("#")]
    records = [dict(zip(rows[0], row, strict=True)) for row in rows[1:]]
    cut = [r for r in records if int(r["entry"]) >= 1000]
    assert {r["glyphs"] for r in cut} == {"3 10 15.5"}
    themed = [r for r in records if r["group"] in suite.NORMALS]
    assert len(themed) == 2 * suite.THEME_PROBES
    for r in themed:
        if int(r["entry"]) >= 1000:
            assert float(r["row_height"]) == {"theme_10": 16.0, "theme_16": 20.0}[r["group"]]
            assert (int(r["red"]), int(r["dark"])) == (0, 235)
        else:
            assert float(r["row_height"]) == 30.0


def test_the_faces_groups_put_references_and_probes_at_their_entries(built):
    for group, normal in suite.FACES.items():
        styles, sheet = _parts(built[group])
        fonts = re.findall(r"<font>(.*?)</font>", styles, re.S)
        assert len(fonts) == suite.FACES_FIRST + suite.FACES_PROBES
        first = "Calibri" if normal is None else normal[0]
        if group == "faces_aptos":
            with zipfile.ZipFile(built[group]) as z:
                theme = z.read("xl/theme/theme1.xml").decode()
            assert '<a:minorFont>\n        <a:latin typeface="Aptos Narrow"/>' in theme
            assert '<scheme val="minor"' in fonts[0]
        assert f'<name val="{first}"' in fonts[0]
        xfs = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", styles, re.S).group(1)
        font_of = [int(f) for f in re.findall(r'<xf [^>]*fontId="(\d+)"', xfs)]
        for cell, entry in suite.probes(group):
            style = re.search(rf'<c r="{cell}"[^>]* s="(\d+)"', sheet)
            assert font_of[int(style.group(1))] == entry, (group, cell)
        for k, (face, _) in enumerate(suite.REFERENCES, 1):
            assert f'<name val="{face}"' in fonts[k], (group, face)
