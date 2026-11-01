import io
import zipfile

import openpyxl
import pytest
from openpyxl.styles.colors import COLOR_INDEX

from sheetpeek._colors import apply_tint, resolve, theme_colors


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
