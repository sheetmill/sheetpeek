"""The capture suite ``theme_colors_more`` (stored 16-bit tints, tints of black), without Excel.

Its workbook is built and read back with sheetpeek as built, every tint written as the ``repr``
of its float, and what sheetpeek draws for each cell is the suite's ``EXPECTED``. Its measuring
(Excel's AppleScript answer turned into records) runs on a synthetic answer. The comparison with
``tests/captures/theme_colors_more.tsv`` is skipped until that capture exists.
"""

import re
import zipfile

import openpyxl
import pytest
from capture import theme_colors_more as suite

import sheetpeek
from sheetpeek._colors import apply_tint
from tests.capture_report import REPORT

CAPTURE = suite.OUTPUT


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    return suite.build(tmp_path_factory.mktemp("theme_colors_more") / "theme_colors_more.xlsx")


@pytest.fixture(scope="module")
def sheet(built):
    return sheetpeek.read(built)


def _cell(sheet, address):
    position = sheetpeek.CellRange.parse(address)
    return sheet.cell(position.min_row, position.min_col)


# --- the workbook


def test_the_cases_cover_the_rules():
    groups = [c[0] for c in suite.CASES]
    assert groups.count("stored") == 70
    assert groups.count("third") == 6
    assert len({c[1] for c in suite.CASES}) == len(suite.CASES) == len(suite.EXPECTED)
    assert {c[2] for c in suite.CASES if c[0] == "stored"} == set(range(10))
    # every stored tint is a 16-bit fraction near its round tint, but not equal to it
    for stored, round_ in suite.STORED_TINTS:
        assert stored != round_
        assert abs(stored - round_) < 2**-15


def test_the_workbook_stores_each_tint_as_its_repr(built):
    with zipfile.ZipFile(built) as z:
        styles = z.read("xl/styles.xml").decode("utf-8")
    written = set(re.findall(r'tint="([^"]*)"', styles))
    assert written == {repr(c[3]) for c in suite.CASES}


def test_exact_tints_rewrites_every_tint():
    text = '<color theme="1" tint="0.59999389629810485"/><color theme="2" tint="-2.5E-1"/>'
    assert suite.exact_tints(text) == (
        '<color theme="1" tint="0.5999938962981048"/><color theme="2" tint="-0.25"/>'
    )


def test_the_workbook_reads_back_as_built(built, sheet):
    ws = openpyxl.load_workbook(built).active
    for _group, address, theme, tint, _round in suite.CASES:
        assert ws[address].value == "x"
        assert ws[address].fill.fgColor.theme == theme
        assert ws[address].fill.fgColor.tint == tint
        cell = _cell(sheet, address)
        assert cell.font.theme_color == theme, address


def test_expected_is_what_sheetpeek_draws(sheet):
    drawn = {
        address: (_cell(sheet, address).fill, _cell(sheet, address).font.color)
        for _group, address, *_ in suite.CASES
    }
    assert drawn == suite.EXPECTED


def test_no_case_draws_magenta():
    for fill, font in suite.EXPECTED.values():
        for color in (fill, font):
            r, g, b = (int(color[i : i + 2], 16) for i in (0, 2, 4))
            assert not (r > 200 and b > 200 and g < 80), color


# --- the measuring, on a synthetic answer


def _rgb(hexed):
    return ",".join(str(int(hexed[i : i + 2], 16)) for i in (0, 2, 4))


def _answer(colors=None):
    colors = colors or suite.EXPECTED
    return "\n".join(
        f"{address}\t{_rgb(colors[address][0])}\t{_rgb(colors[address][1])}"
        for _group, address, *_ in suite.CASES
    )


def test_measure_returns_the_colors_answered():
    records = suite.measure(_answer())
    assert len(records) == len(suite.CASES)
    first = records[0]
    assert first == {
        "group": "stored",
        "cell": "A1",
        "theme": "0",
        "tint": "0.5999938962981048",
        "round": "0.6",
        "fill": suite.EXPECTED["A1"][0],
        "font": suite.EXPECTED["A1"][1],
    }
    assert {r["cell"]: (r["fill"], r["font"]) for r in records} == suite.EXPECTED


def test_measure_keeps_fill_and_font_apart():
    colors = dict(suite.EXPECTED)
    colors["B12"] = ("545454", "565656")
    record = next(r for r in suite.measure(_answer(colors)) if r["cell"] == "B12")
    assert (record["fill"], record["font"]) == ("545454", "565656")


def test_measure_refuses_an_answer_that_does_not_fit():
    lines = _answer().split("\n")
    with pytest.raises(ValueError, match="expected 76 cells"):
        suite.measure("\n".join(lines[:-1]))
    with pytest.raises(ValueError, match="answered for B1"):
        suite.measure("\n".join([lines[1], lines[0], *lines[2:]]))
    broken = lines[0].replace("\t255,255,255", "\tmissing value", 1)
    with pytest.raises(ValueError, match="not an r,g,b"):
        suite.measure("\n".join([broken, *lines[1:]]))


def test_write_measures_a_saved_answer_again(tmp_path):
    saved = tmp_path / "excel.tsv"
    saved.write_text("# excel: 16.115\n# date: 2026-01-01\n" + _answer() + "\n", encoding="utf-8")
    output = tmp_path / "out.tsv"
    suite.write(saved, output)
    lines = output.read_text(encoding="utf-8").splitlines()
    assert lines[:3] == ["# excel: 16.115", "# date: 2026-01-01", "# suite: theme_colors_more"]
    rows = [line.split("\t") for line in lines if not line.startswith("# ")]
    assert rows[0] == list(suite.COLUMNS)
    assert len(rows) == 1 + len(suite.CASES)
    assert rows[-1] == ["third", "C13", "3", "0.3333333333", "0.3333333333", *suite.EXPECTED["C13"]]


# --- the comparison with Excel
#
# Excel answers whole 0-255 channels, and the integer-HLS tint (§3.3) reproduced every one of the
# 96 cells of the ``theme_colors`` capture exactly, so the tolerance is none: a channel off by one
# is the rule failing (a lightness truncated one step short), not noise.


@pytest.fixture(scope="module")
def captured():
    if not CAPTURE.exists():
        pytest.skip("not captured yet: run python -m capture.theme_colors_more run")
    lines = [ln for ln in CAPTURE.read_text(encoding="utf-8").splitlines() if ln[:2] != "# "]
    names = lines[0].split("\t")
    return [dict(zip(names, line.split("\t"), strict=True)) for line in lines[1:]]


@pytest.mark.excel
def test_the_capture_holds_every_case(captured):
    assert [(r["group"], r["cell"]) for r in captured] == [(c[0], c[1]) for c in suite.CASES]


@pytest.mark.excel
def test_stored_tints_draw_as_excel_draws_them(captured, sheet):
    # Rule tint_stored_fraction (§3.3, ``_colors.apply_tint``): a stored 16-bit tint draws as
    # the round tint it stands for. Compared twice: with what sheetpeek draws for the stored
    # tint, and with sheetpeek's integer-HLS tint at the round value (captured for 0.2 to 0.8
    # and -0.25 by ``theme_colors``).
    rows = [r for r in captured if r["group"] == "stored"]
    drawn, as_round = {}, {}
    for r in rows:
        cell = _cell(sheet, r["cell"])
        excel = (r["fill"], r["font"])
        if (cell.fill, cell.font.color) != excel:
            drawn[r["cell"]] = ((cell.fill, cell.font.color), excel)
        base = sheet.theme_colors[int(r["theme"])]
        round_ = apply_tint(base, float(r["round"]))
        if (round_, round_) != excel:
            as_round[r["cell"]] = (r["tint"], round_, excel)
    REPORT["theme colors stored tints"] = (
        f"theme colors, stored tints: {len(rows)} cells compared with Excel, {len(drawn)} "
        f"deviations, {len(as_round)} not drawn as the round tint"
    )
    assert drawn == {}
    assert as_round == {}


@pytest.mark.excel
def test_a_third_of_the_way_to_white_matches_excel(captured, sheet):
    # Rule tint_third_black (§3.3): tint 0.333333 on black is 555555 (lightness 80 of 240, not
    # 79), and its neighbors 0.33333 and 0.3333333333 with it; theme 3 (1F497D) beside it.
    rows = [r for r in captured if r["group"] == "third"]
    drawn = {}
    for r in rows:
        cell = _cell(sheet, r["cell"])
        if (cell.fill, cell.font.color) != (r["fill"], r["font"]):
            drawn[r["cell"]] = ((cell.fill, cell.font.color), (r["fill"], r["font"]))
    REPORT["theme colors thirds"] = (
        f"theme colors, tints of a third: {len(rows)} cells compared with Excel, "
        f"{len(drawn)} deviations"
    )
    black = {r["cell"]: (r["fill"], r["font"]) for r in rows if r["theme"] == "1"}
    assert drawn == {}
    assert black == dict.fromkeys(("A12", "B12", "C12"), ("555555", "555555"))
