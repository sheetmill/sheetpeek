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

PINNED: dict[str, frozenset[str]] = {}
"""Cases where sheetpeek's row differs from Excel's, by comparison, each with its reason in
``REASONS``. Filled from the capture."""
REASONS: dict[str, str] = {}


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
        assert captured[case.name]["applied"] == f"{case.face} {case.size:g}", case.name
