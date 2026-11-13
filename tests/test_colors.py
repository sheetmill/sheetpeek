import importlib.util
import io
import re
import zipfile

import openpyxl
import pytest
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import COLOR_INDEX, Color

import sheetpeek
from sheetpeek._colors import apply_tint, resolve, theme_colors
from tests.workbooks import rewrite, save


@pytest.fixture(scope="module")
def default_theme():
    stream = io.BytesIO()
    openpyxl.Workbook().save(stream)
    with zipfile.ZipFile(stream) as archive:
        return theme_colors(archive.read("xl/theme/theme1.xml"))


def test_theme_order(default_theme):
    # openpyxl's default theme is Office 2007's. Excel's index 0 is lt1 and 1 is dk1, the
    # reverse of the XML order (captured, suite theme_colors; ARCHITECTURE.md §3.3).
    assert len(default_theme) == 12
    assert default_theme[:4] == ("FFFFFF", "000000", "EEECE1", "1F497D")
    assert default_theme[4] == "4F81BD"


def test_unreadable_theme():
    assert theme_colors(None) == ()
    assert theme_colors(b"not xml") == ()
    assert theme_colors(b"<theme/>") == ()


def test_tint():
    assert apply_tint("4F81BD", 0) == "4F81BD"
    assert apply_tint("4F81BD", 1) == "FFFFFF"
    assert apply_tint("4F81BD", -1) == "000000"
    lighter, darker = apply_tint("808080", 0.5), apply_tint("808080", -0.5)
    assert int(lighter[:2], 16) > 0x80 > int(darker[:2], 16)


def test_resolve(default_theme):
    def check(kind, value, tint=0.0):
        return resolve(kind, value, tint, default_theme, COLOR_INDEX)

    assert check(None, None) is None
    assert check("auto", None) is None
    assert check("rgb", "FF123456") == "123456"
    assert check("rgb", "abcdef") == "ABCDEF"
    assert check("rgb", "Values must be of type <class 'str'>") is None
    assert check("indexed", 10) == "FF0000"
    assert check("indexed", 64) is None  # system foreground: automatic
    assert check("indexed", 99) is None
    assert check("theme", 4) == "4F81BD"
    assert check("theme", 4, 0.4) == apply_tint("4F81BD", 0.4)
    assert check("theme", 40) is None


@pytest.mark.parametrize(
    "value",
    [
        "&quot;&gt;&lt;script&gt;alert(1)&lt;/script&gt;",
        "red;background:url(x)",
        "&#155;ABCDE",
        "12345",
    ],
)
@pytest.mark.parametrize("attribute", ["srgbClr val", "sysClr lastClr"])
def test_a_theme_color_that_is_not_one_reaches_no_output(tmp_path, value, attribute):
    # It would go into a style attribute as it stands, and crash rich.
    element, name = attribute.split()
    wb = openpyxl.Workbook()
    wb.active["A1"] = "text"
    wb.active["A1"].fill = PatternFill("solid", fgColor=Color(theme=4))
    wb.active["A1"].font = Font(color=Color(theme=4))
    path = save(wb, tmp_path / "theme.xlsx")

    def edit(text):
        new = f'<a:accent1><a:{element} {name}="{value}"/></a:accent1>'
        return re.sub(r"<a:accent1>.*?</a:accent1>", new, text, flags=re.S)

    rewrite(path, "xl/theme/theme1.xml", edit)
    sheet = sheetpeek.read(path)
    assert sheet.cell(1, 1).style.fill is None
    assert sheet.cell(1, 1).style.font.color is None
    page = sheetpeek.render(sheet, "html")
    assert "<script" not in page.lower()
    assert "url(x)" not in page.lower()
    sheetpeek.render(sheet, "text")
    if importlib.util.find_spec("rich"):
        sheetpeek.render(sheet, "rich")


def test_resolve_gives_only_colors():
    assert resolve("rgb", "<SCRIPT", 0.0, (), ()) is None
    assert resolve("theme", 0, 0.0, ('"><X>Y',), ()) is None
    assert resolve("indexed", 0, 0.0, (), ('FF"><X>Y',)) is None
    assert resolve("rgb", "FFABCDEF", 0.0, (), ()) == "ABCDEF"


# Capture on Excel 16.114, unverified on Microsoft 365 (14 cells' own fills, read exactly): files
# store a tint as a 16-bit fraction, and Excel draws the stored form as the round
# tint it stands for.
STORED_TINTS = [
    (0.59999389629810485, 0.6),
    (0.3999755851924192, 0.4),
    (0.79998168889431442, 0.8),
    (0.249977111117893, 0.25),
    (0.499984740745262, 0.5),
    (0.14999847407452621, 0.15),
    (-0.249977111117893, -0.25),
]


@pytest.mark.parametrize(("stored", "round_"), STORED_TINTS)
@pytest.mark.parametrize("rgb", ["9BBB59", "000000", "1F497D", "EEECE1", "C0504D"])
def test_a_stored_tint_draws_as_its_round_value(rgb, stored, round_):
    assert apply_tint(rgb, stored) == apply_tint(rgb, round_)


def test_a_third_of_the_way_to_white_from_black():
    # Excel 16.114 capture: tint 0.333333 on black is 555555 (lightness 80 of 240, not 79).
    assert apply_tint("000000", 0.333333) == "555555"


def test_without_the_tolerance_stored_tints_fall_short(monkeypatch):
    # The mutation check: truncating without the tolerance breaks the captured cases.
    monkeypatch.setattr(sheetpeek._colors, "_TINT_TOLERANCE", 0.0)
    wrong = [
        (rgb, stored)
        for stored, round_ in STORED_TINTS
        for rgb in ("9BBB59", "000000", "1F497D", "EEECE1", "C0504D")
        if apply_tint(rgb, stored) != apply_tint(rgb, round_)
    ]
    assert wrong
    assert apply_tint("000000", 0.333333) != "555555"
