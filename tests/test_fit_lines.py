"""Row fitting's second wrap (``_excel._fit_lines``, ARCHITECTURE.md §4.7), without Excel.

``_fit_lines`` asks xlformat's ``autofit_width`` only for lines near the budget; the bounds it
answers the others from must never change a line. Its count is compared with a wrap that asks
xlformat for every line, over texts, faces, sizes, widths and both platforms.
"""

import random

import pytest
import xlformat

from sheetpeek import _platform
from sheetpeek._excel import _fit_lines
from sheetpeek._runs import wrap

WORDS = ("a", "aee", "wwww", "Hxg", "spilled", "column", "1234567", "MMMMMMMM", "café", "ñandú")
FONTS = (
    xlformat.Font("Calibri", 11),
    xlformat.Font("Calibri", 11, italic=True),
    xlformat.Font("Arial", 10),
    xlformat.Font("Times New Roman", 24, bold=True),
    xlformat.Font("Calibri", 6),
)


def _reference(text, budget, font, normal, platform):
    unit = 0.75 if platform.name == "windows" else 1.0
    step = 4 if platform.name == "windows" else 5
    period = round(xlformat.digit_width(normal, platform=platform) / unit) * step
    lean = xlformat.overhang(font, platform=platform)

    def width(a, b):
        line = text[a:b]
        found = xlformat.autofit_width(line, font, normal_font=normal, platform=platform)
        counted = round((xlformat.text_width(line, font, platform=platform) + lean) / unit)
        if counted >= 3 * period and counted % period == 0:
            found -= unit
        return found

    return len(wrap(text, budget, width))


@pytest.mark.parametrize("name", ["mac", "windows"])
def test_the_bounds_change_no_line(name):
    platform = _platform.resolve(name)
    normal = xlformat.Font("Calibri", 11)
    rng = random.Random(7)
    with _platform.measuring(platform):
        for _ in range(1500):
            font = rng.choice(FONTS)
            text = " ".join(rng.choice(WORDS) for _ in range(rng.randint(1, 8)))
            budget = rng.choice((30.0, 48.0, 64.0, 66.0, 90.0, 97.0, 128.0)) + rng.randint(-8, 8)
            got = _fit_lines(text, budget, font, normal)
            assert got == _reference(text, budget, font, normal, platform), (text, font, budget)


def test_a_line_in_the_band_is_fitted_as_two():
    # A 64 px column on Windows draws "bbbb cccc" (59 px) whole and fits it as two lines:
    # 59 + 7 > 64 (captured, windows/text_style group align, wrap_band).
    platform = _platform.resolve("windows")
    with _platform.measuring(platform):
        font = xlformat.Font("Calibri", 11)
        assert _fit_lines("bbbb cccc", 48.0, font, font) == 2
        assert _fit_lines("bbbb ccc", 48.0, font, font) == 1
