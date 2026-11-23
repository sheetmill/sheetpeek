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
    # Superscript and subscript text, alone and side by side, from 12 pt: sheetpeek's linear
    # growth is not Excel's, which is not yet modeled (§4.13).
    "styled": _names(
        "both 12|both 13|both 14|both 16|both 17|both 18|both 19|both 20|both 21|both 22|both 23|"
        "both 24|both 25|both 26|both 27|both 28|both 29|both 30|both 31|both 32|both 33|both 34|"
        "both 35|both 36|both 37|both 38|both 39|both 40|subscript 13|subscript 14|subscript 15|"
        "subscript 17|subscript 18|subscript 19|subscript 20|subscript 21|subscript 22|"
        "subscript 23|subscript 24|subscript 26|subscript 27|subscript 28|subscript 29|"
        "subscript 30|subscript 32|subscript 33|subscript 34|subscript 35|subscript 36|"
        "subscript 37|subscript 38|subscript 39|subscript 40|superscript 12|superscript 16|"
        "superscript 17|superscript 18|superscript 19|superscript 20|superscript 21|superscript 22|"
        "superscript 23|superscript 24|superscript 25|superscript 26|superscript 27|superscript 28|"
        "superscript 29|superscript 30|superscript 31|superscript 32|superscript 33|superscript 34|"
        "superscript 35|superscript 36|superscript 37|superscript 38|superscript 39|superscript 40"
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
