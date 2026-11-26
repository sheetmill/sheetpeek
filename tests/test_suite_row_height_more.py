"""The capture suite ``row_height_more`` (row heights under other Normal fonts and wrapped rows in
other fonts, ARCHITECTURE.md §4.7), without Excel.

Its workbooks are built and read back with sheetpeek as built, and sheetpeek's layout of each row
is the suite's ``EXPECTED``. Its reading of Excel's answer runs on a synthetic answer. The
comparison with ``tests/captures/row_height_more.tsv`` is skipped until that capture exists.
"""

import openpyxl
import pytest
from capture import row_height_more as suite

import sheetpeek
from tests.capture_report import REPORT


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    folder = tmp_path_factory.mktemp("row_height_more")
    return folder, suite.build(folder)


def _layout(folder, number):
    rows = max(c.row for c in suite.CASES if c.workbook == number)
    sheet = sheetpeek.read(folder / suite.workbook_name(number), ref=f"A1:A{rows}", platform="mac")
    return sheet, {row.index: row for row in sheet.layout("excel").rows}


def test_every_case_is_built_as_described(built):
    folder, paths = built
    assert [p.name for p in paths] == [suite.workbook_name(k) for k in range(1, 18)]
    assert len(suite.CASES) == 3 * len(suite.NORMALS) + len(suite.CELL_ROWS)
    for k in range(1, 18):
        cases = [c for c in suite.CASES if c.workbook == k]
        sheet, _ = _layout(folder, k)
        assert (sheet.normal_font.name, sheet.normal_font.size) == cases[0].normal
        wb = openpyxl.load_workbook(folder / suite.workbook_name(k))
        ws = wb.active
        for case in cases:
            if case.label == "no row element":
                assert case.row not in ws.row_dimensions or ws.cell(case.row, 1).value is None
                continue
            cell = sheet.cell(case.row, 1)
            assert (cell.font.name, cell.font.size) == case.font, case
            assert cell.style.wrap == case.wrapped, case
            assert cell.value == (suite.WRAPPED if case.wrapped else 1), case


def test_no_row_element_in_row_2(built):
    folder, _ = built
    for k in range(1, len(suite.NORMALS) + 1):
        sheet, _ = _layout(folder, k)
        assert [r.listed for r in sheet.rows] == [True, False, True]


def test_expected_is_what_sheetpeek_lays_out(built):
    folder, _ = built
    found = {}
    for k in range(1, 18):
        _, rows = _layout(folder, k)
        for case in suite.CASES:
            if case.workbook == k:
                found[(k, case.row)] = rows[case.row].height_pt
    assert found == suite.EXPECTED


def _answer(heights=None, applied=None):
    """A synthetic answer from Excel: every case's row, with the given or EXPECTED heights."""
    lines = ["# excel: Excel 16.0 (1, Test edition, Test channel) for Mac, en_US", "# captured: x"]
    for case in suite.CASES:
        key = (case.workbook, case.row)
        height = (heights or {}).get(key, suite.EXPECTED[key])
        name, size = (applied or {}).get(key, case.font)
        lines.append(
            "\t".join(
                [suite.workbook_name(case.workbook), str(case.row), name, f"{size}.0", str(height)]
            )
        )
    return "\n".join(lines) + "\n"


def test_reading_an_answer_returns_its_values(tmp_path):
    answer = tmp_path / "answer.tsv"
    answer.write_text(
        _answer(heights={(1, 1): 15.75, (17, 18): 62}, applied={(5, 1): ("Courier", 11)}),
        encoding="utf-8",
    )
    output = tmp_path / "row_height_more.tsv"
    suite.write(answer, output)
    rows = _parse(output)
    assert len(rows) == len(suite.CASES)
    first = rows[0]
    assert (first["normal"], first["row"], first["label"], first["height"]) == (
        "Arial 11",
        "1",
        "number",
        "15.75",
    )
    assert rows[1]["height"] == "14"  # 14.0 as 14
    by_key = {(r["workbook"], r["row"]): r for r in rows}
    assert by_key[("5", "1")]["applied"] == "Courier 11"
    assert by_key[("17", "18")]["height"] == "62"
    assert by_key[("17", "18")]["font"] == "Calibri 14"
    assert by_key[("17", "18")]["wrapped"] == "yes"
    assert output.read_text(encoding="utf-8").startswith(
        "# suite: row_height_more (capture.row_height_more); Excel 16.0 (1, Test edition"
    )


def test_an_answer_out_of_order_is_refused():
    lines = _answer().split("\n")
    lines[2], lines[3] = lines[3], lines[2]
    with pytest.raises(ValueError, match="expected"):
        suite.records("\n".join(lines))
    with pytest.raises(ValueError, match="lines from Excel"):
        suite.records("\n".join(lines[:-3]))


# The comparison with Excel ------------------------------------------------------------------
#
# Excel's "row height" is its own layout of the row in points, read through AppleScript, not a
# measurement of pixels: the suite row_height reads whole points at every font it captured, and
# sheetpeek's heights are whole points from the same kind of read. The tolerance is therefore
# zero. A row whose applied font differs from the font asked for is Excel substituting a face;
# it is reported, not compared.


def _parse(path):
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    return [dict(zip(header, row, strict=True)) for row in body]


def _captured():
    if not suite.OUTPUT.exists():
        pytest.skip("not captured yet: run python -m capture.row_height_more run")
    return {(int(r["workbook"]), int(r["row"])): r for r in _parse(suite.OUTPUT)}


KNOWN: frozenset[tuple[int, int]] = frozenset()
"""Rows where sheetpeek's height differs from Excel's, each with its reason: none. (Times New
Roman and Georgia, once measured as Calibri and a point taller, take xlformat's own rows.)"""


def _compare(name, cases):
    captured = _captured()
    substituted = []
    wrong = []
    for case in cases:
        record = captured[(case.workbook, case.row)]
        if record["applied"] != f"{case.font[0]} {case.font[1]}":
            substituted.append((case.label, case.normal, record["applied"]))
            continue
        expected = suite.EXPECTED[(case.workbook, case.row)]
        key = (case.workbook, case.row)
        if (float(record["height"]) != expected) != (key in KNOWN):
            wrong.append((case.normal, case.label, record["height"], expected))
    known = sum(1 for case in cases if (case.workbook, case.row) in KNOWN)
    REPORT[f"row_height_more {name}"] = (
        f"row_height_more {name}: {len(cases) - known} of {len(cases)} rows as Excel lays out, "
        f"{known} known deviations"
    )
    assert not substituted, f"Excel substituted a face: {substituted}"
    assert not wrong, (
        f"(Normal font, row, Excel, sheetpeek; a new difference or a known one gone): {wrong}"
    )


@pytest.mark.excel
def test_other_normal_faces_take_their_own_heights_as_captured():
    # §4.7: Arial, Times New Roman, Courier New and Georgia at 11 and 14 as a Normal font give
    # their rows xlformat's height for the face, listed (row 1) or not (row 2).
    faces = set(suite.OTHER_FACES)
    _compare(
        "other Normal faces",
        [c for c in suite.CASES if c.normal[0] in faces and not c.wrapped],
    )


@pytest.mark.excel
def test_normal_sizes_between_measured_ones_take_their_own_heights_as_captured():
    # §4.7: Calibri and Aptos Narrow at 13, 15, 17 and 22 take xlformat's height at the size.
    _compare(
        "Normal sizes between",
        [
            c
            for c in suite.CASES
            if c.workbook <= len(suite.NORMALS) and c.normal[1] in suite.BETWEEN and not c.wrapped
        ],
    )


@pytest.mark.excel
def test_rows_fit_cell_fonts_the_table_has_not_measured_as_captured():
    # §4.7: under a Calibri 11 Normal font, a number in each of those fonts fits its row to the
    # font's height from xlformat's row model.
    _compare(
        "cell fonts fitted",
        [c for c in suite.CASES if c.workbook == len(suite.NORMALS) + 1 and not c.wrapped],
    )


@pytest.mark.excel
def test_three_wrapped_lines_in_other_fonts_as_captured():
    # §4.7: k wrapped lines take k(F + 1) plus the inset difference; Calibri 11 (48 pt) is the
    # captured control. Wrapped in the Normal font of every workbook, and in Calibri 14 and 20,
    # Aptos Narrow 11 and the four other faces under Calibri 11.
    _compare("wrapped lines", [c for c in suite.CASES if c.wrapped])
