"""The capture suite ``row_height_fitted`` (rows fitted to turned and to styled text,
ARCHITECTURE.md §4.7, §4.16, §4.17), without Excel.

Its workbooks are built and read back with sheetpeek as built, and its reading of Excel's answer
runs on a synthetic answer. The comparisons with ``tests/captures/row_height_fitted.tsv`` are
skipped until that capture exists; each compares sheetpeek's fitted row with Excel's, and the
cases that differ are named in ``PINNED`` with their reason, so a new difference fails and so
does a fixed one.
"""

import pytest
from capture import row_height_fitted as suite

import sheetpeek
from tests.capture_report import REPORT

CAPTURE = suite.OUTPUT


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    folder = tmp_path_factory.mktemp("row_height_fitted")
    return folder, suite.build(folder)


@pytest.fixture(scope="module")
def laid_out(built):
    """sheetpeek's height of each case's row, by case name."""
    folder, _ = built
    found = {}
    for group in suite.GROUPS:
        cases = suite.rows(group)
        sheet = sheetpeek.read(folder / suite.workbook_name(group), ref=f"A1:B{len(cases)}")
        heights = {row.index: row.height_pt for row in sheet.layout("excel").rows}
        for row, case in cases:
            found[case.name] = heights[row]
    return found


def test_every_case_is_built_as_described(built):
    folder, paths = built
    assert [p.name for p in paths] == [suite.workbook_name(g) for g in suite.GROUPS]
    assert len({c.name for c in suite.CASES}) == len(suite.CASES)
    for group in suite.GROUPS:
        sheet = sheetpeek.read(folder / suite.workbook_name(group))
        assert (sheet.normal_font.name, sheet.normal_font.size) == ("Calibri", 11.0)
        assert all(row.height is None for row in sheet.rows), group
        for row, case in suite.rows(group):
            cell = sheet.cell(row, 1)
            assert cell.value == case.text, case
            font = cell.style.font
            assert (font.name, font.size) == (case.face, case.size), case
            assert font.vert_align == case.vert_align, case
            assert font.underline == case.underline, case
            assert cell.style.rotation == case.rotation, case
            if case.beside:
                assert sheet.cell(row, 2).style.font.vert_align == case.beside, case


def test_the_cases_cover_the_rules():
    turned = [c for c in suite.CASES if c.group == "turned"]
    assert {c.rotation for c in turned} == {90, 88, 89, -90}
    assert {c.size for c in turned if c.rotation == 90} == set(range(9, 37))
    styled = [c for c in suite.CASES if c.group == "styled"]
    assert {c.size for c in styled if c.vert_align == "subscript" and not c.beside} == set(
        range(6, 41)
    )
    assert 12.5 in {c.size for c in styled if c.name.startswith("plain")}
    assert suite.stored_rotation(-90) == 180
    assert suite.stored_rotation(90) == 90


def _answer(heights=None):
    """A synthetic answer from Excel: every case's row, applied font, a height."""
    lines = ["# excel: Excel 16.0 (1, Test edition, Test channel) for Mac, en_US", "# captured: x"]
    for group in suite.GROUPS:
        for row, case in suite.rows(group):
            height = (heights or {}).get(case.name, 15.0)
            lines.append(
                "\t".join(
                    [suite.workbook_name(group), str(row), case.face, f"{case.size}", str(height)]
                )
            )
    return "\n".join(lines) + "\n"


def _parse(path):
    lines = [ln for ln in path.read_text(encoding="utf-8").splitlines() if ln[:2] != "# "]
    names = lines[0].split("\t")
    return [dict(zip(names, line.split("\t"), strict=True)) for line in lines[1:]]


def test_reading_an_answer_returns_its_values(tmp_path):
    answer = tmp_path / "answer.tsv"
    answer.write_text(_answer({"90 long 20": 294.0, "plain 12.5": 17.25}), encoding="utf-8")
    output = tmp_path / "row_height_fitted.tsv"
    suite.write(answer, output)
    records = {r["name"]: r for r in _parse(output)}
    assert len(records) == len(suite.CASES)
    assert records["90 long 20"]["height"] == "294"  # 294.0 as 294
    assert records["90 long 20"]["rotation"] == "90"
    assert records["plain 12.5"]["height"] == "17.25"
    assert records["plain 12.5"]["applied"] == "Calibri 12.5"
    assert records["-90 short 20"]["rotation"] == "-90"


def test_an_answer_for_another_row_is_refused(tmp_path):
    text = _answer().replace(f"{suite.workbook_name('turned')}\t1\t", "x.xlsx\t1\t", 1)
    with pytest.raises(ValueError, match=r"answered for x\.xlsx"):
        suite.records(text)


# --- the comparisons with the capture


def _names(text):
    return frozenset(text.split("|"))


PINNED: dict[str, frozenset[str]] = {
    # At sizes xlformat has not measured (13, 15, 17, 19, 21 to 23, 25 to 36) the width W comes
    # from xlformat's table scaled from the nearest measured size, which is off by up to 20 pt in
    # the long text; every row at a measured size, 90 and -90 degrees, is exact.
    "turned": _names(
        "90 medium 13|90 long 13|90 medium 15|90 long 15|90 short 17|90 medium 17|90 long 17|"
        "90 short 19|90 medium 19|90 long 19|90 short 21|90 medium 21|90 long 21|"
        "90 medium 22|90 long 22|90 medium 23|90 long 23|90 short 25|90 medium 25|90 long 25|"
        "90 medium 26|90 long 26|90 short 27|90 medium 27|90 long 27|90 medium 28|90 long 28|"
        "90 short 29|90 medium 29|90 long 29|90 medium 30|90 short 31|90 medium 31|"
        "90 medium 32|90 short 33|90 medium 33|90 medium 34|90 short 35|90 medium 35|"
        "90 short 36|90 ox1 17|90 ox2 17|90 ox4 17|90 ix3 19|90 ox1 21|90 ox2 21|90 ox4 21|"
        "90 ix3 22|90 ix3 23|90 ox1 25|90 ox2 25|90 ox4 25|90 ox1 27|90 ox2 27|90 ox4 27|"
        "90 ix3 27|90 ix3 28|90 ox1 29|90 ox2 29|90 ox4 29|90 ix3 30|90 ox1 31|90 ox2 31|"
        "90 ox4 31|90 ix3 31|90 ix3 32|90 ox1 33|90 ox2 33|90 ox4 33|90 ox1 35|90 ox2 35|"
        "90 ox4 35|90 ix3 35|90 ox1 36|90 ox2 36|90 ox4 36|90 ix3 36"
    ),
    # The row a font fits at a size the row model has not measured: half sizes, Calibri 19, 21, 23
    # and past 24 (sheetpeek takes 24's 31 pt), and Arial, Times New Roman, Courier New and Aptos
    # Narrow at sizes between and past the measured ones.
    "plain": _names(
        "plain 11.5|plain 12.5|plain 13.5|plain 14.5|plain 15.5|plain 16.5|plain 17.5|"
        "plain 18.5|plain 19|plain 20.5|plain 21|plain 22.5|plain 23|plain 24.5|plain 25|"
        "plain 25.5|plain 26|plain 26.5|plain 27|plain 27.5|plain 28|plain 28.5|plain 29|"
        "plain 29.5|plain 30|plain 30.5|plain 31|plain 31.5|plain 32|plain 32.5|plain 33|"
        "plain 33.5|plain 34|plain 34.5|plain 35|plain 35.5|plain 36|plain 36.5|plain 37|"
        "plain 37.5|plain 38|plain 38.5|plain 39|plain 39.5|plain 40|plain 41|plain 42|"
        "plain 43|plain 44|plain 45|plain 46|plain 47|plain 48|plain 49|plain 50|plain 51|"
        "plain 52|plain 53|plain 54|plain 55|plain 56|plain 57|plain 58|plain 59|plain 60|"
        "plain 61|plain 62|plain 63|plain 64|plain 65|plain 66|plain 67|plain 68|plain 69|"
        "plain 70|plain 71|plain 72|Courier New 12|Arial 13|Courier New 13|"
        "Times New Roman 14|Arial 15|Times New Roman 15|Courier New 15|Times New Roman 16|"
        "Courier New 16|Arial 17|Times New Roman 17|Times New Roman 18|Arial 19|"
        "Courier New 19|Aptos Narrow 19|Times New Roman 20|Courier New 20|Arial 21|"
        "Times New Roman 21|Courier New 21|Aptos Narrow 21|Arial 22|Times New Roman 22|"
        "Courier New 22|Arial 23|Courier New 23|Aptos Narrow 23|Times New Roman 24|"
        "Courier New 24|Arial 25|Courier New 25|Aptos Narrow 25|Arial 26|Times New Roman 26|"
        "Courier New 26|Aptos Narrow 26|Arial 27|Times New Roman 27|Courier New 27|"
        "Aptos Narrow 27|Arial 28|Times New Roman 28|Courier New 28|Aptos Narrow 28|Arial 29|"
        "Times New Roman 29|Courier New 29|Aptos Narrow 29|Arial 30|Times New Roman 30|"
        "Courier New 30|Aptos Narrow 30|Arial 31|Times New Roman 31|Courier New 31|"
        "Aptos Narrow 31|Arial 32|Times New Roman 32|Courier New 32|Aptos Narrow 32|Arial 33|"
        "Times New Roman 33|Courier New 33|Aptos Narrow 33|Arial 34|Times New Roman 34|"
        "Courier New 34|Aptos Narrow 34|Arial 35|Times New Roman 35|Courier New 35|"
        "Aptos Narrow 35|Arial 36|Times New Roman 36|Courier New 36|Aptos Narrow 36|Arial 37|"
        "Times New Roman 37|Courier New 37|Aptos Narrow 37|Arial 38|Times New Roman 38|"
        "Courier New 38|Aptos Narrow 38|Arial 39|Times New Roman 39|Courier New 39|"
        "Aptos Narrow 39|Arial 40|Times New Roman 40|Courier New 40|Aptos Narrow 40"
    ),
    # Superscript and subscript text, alone and side by side, from 12 pt, and an accounting
    # underline at 6 to 10 and from 19 pt: sheetpeek's linear growth is not Excel's, which is not
    # yet modeled.
    "styled": _names(
        "singleAccounting 6|singleAccounting 7|singleAccounting 8|singleAccounting 9|"
        "singleAccounting 10|superscript 12|both 12|subscript 13|both 13|subscript 14|"
        "both 14|subscript 15|superscript 16|both 16|superscript 17|subscript 17|both 17|"
        "superscript 18|subscript 18|both 18|subscript 19|both 19|singleAccounting 19|"
        "superscript 20|subscript 20|both 20|subscript 21|singleAccounting 21|superscript 22|"
        "subscript 22|both 22|superscript 23|subscript 23|both 23|singleAccounting 23|"
        "superscript 24|subscript 24|both 24|subscript 25|singleAccounting 25|subscript 26|"
        "both 26|singleAccounting 26|superscript 27|subscript 27|singleAccounting 27|"
        "superscript 28|subscript 28|both 28|singleAccounting 28|superscript 29|subscript 29|"
        "both 29|singleAccounting 29|superscript 30|subscript 30|both 30|singleAccounting 30|"
        "superscript 31|subscript 31|both 31|singleAccounting 31|superscript 32|subscript 32|"
        "both 32|singleAccounting 32|superscript 33|subscript 33|both 33|singleAccounting 33|"
        "superscript 34|subscript 34|both 34|singleAccounting 34|superscript 35|subscript 35|"
        "both 35|singleAccounting 35|superscript 36|subscript 36|both 36|singleAccounting 36|"
        "superscript 37|subscript 37|both 37|singleAccounting 37|superscript 38|subscript 38|"
        "both 38|singleAccounting 38|superscript 39|subscript 39|both 39|singleAccounting 39|"
        "superscript 40|subscript 40|both 40|singleAccounting 40"
    ),
}
"""Cases where sheetpeek's row differs from Excel's, by comparison, with their reasons."""


@pytest.fixture(scope="module")
def captured():
    if not CAPTURE.exists():
        pytest.skip("not captured yet: run python -m capture.row_height_fitted run")
    return {r["name"]: r for r in _parse(CAPTURE)}


def _compare(key, captured, laid_out, cases, bound=0.0):
    differ = sorted(
        c.name for c in cases if abs(float(captured[c.name]["height"]) - laid_out[c.name]) > bound
    )
    REPORT[f"row_height_fitted {key}"] = (
        f"row_height_fitted {key}: {len(cases) - len(differ)} of {len(cases)} rows as Excel fits "
        f"them, {len(differ)} known deviations"
    )
    assert set(differ) == PINNED.get(key, frozenset()), (
        key,
        sorted(set(differ) ^ PINNED.get(key, frozenset())),
    )


@pytest.mark.excel
def test_rows_at_90_and_minus_90_degrees_as_captured(captured, laid_out):
    # §4.17, fitted rows: the 90-degree rule, exact.
    cases = [c for c in suite.CASES if c.group == "turned" and c.rotation in (90, -90)]
    _compare("turned", captured, laid_out, cases)


@pytest.mark.excel
def test_rows_at_88_and_89_degrees_as_captured(captured, laid_out):
    # §4.17, fitted rows: the general rule, within rotation's FITTED_BOUND.
    cases = [c for c in suite.CASES if c.rotation in (88, 89)]
    _compare("angled", captured, laid_out, cases, bound=2.0)


@pytest.mark.excel
def test_rows_of_plain_text_at_every_size_as_captured(captured, laid_out):
    # §4.7: a cell font's row, in each face, at whole and half sizes.
    cases = [c for c in suite.CASES if c.group == "styled" and not c.vert_align and not c.underline]
    _compare("plain", captured, laid_out, cases)


@pytest.mark.excel
def test_rows_of_superscript_subscript_and_accounting_text_as_captured(captured, laid_out):
    # §4.16, §4.13: the rows superscript, subscript and accounting-underlined text fit.
    cases = [c for c in suite.CASES if c.vert_align or c.underline]
    _compare("styled", captured, laid_out, cases)


@pytest.mark.excel
def test_excel_applied_each_font_as_written(captured):
    for case in suite.CASES:
        # Excel reports a half size cut to the whole one below (12.5 as 12), as for 409.5
        # (package_more).
        assert captured[case.name]["applied"] == f"{case.face} {int(case.size)}", case.name
