"""The capture suite ``wrap_band`` (where a fitted wrapped row takes a second line, ARCHITECTURE.md
§4.7), without Excel.

The workbooks are built and laid out with sheetpeek, whose column widths and row heights are
compared with Excel's in ``tests/captures/wrap_band.tsv`` (Excel for Mac) and
``tests/captures/windows/wrap_band.tsv`` (Excel for Windows). The reading of Excel's answer runs
on a synthetic answer.
"""

import pytest
from capture import wrap_band as suite

import sheetpeek
from tests.capture_report import REPORT

pytest.importorskip("openpyxl")

CAPTURES = {
    "mac": suite.OUTPUT,
    "windows": suite.ROOT / "tests" / "captures" / "windows" / "wrap_band.tsv",
}
TALLER = {"mac": 48, "windows": 64}
"""Rows Excel fits taller than sheetpeek, each by one line: a line of the text lies between the
width Excel draws whole (the column less its insets) and the width its row fitting keeps whole
(the column less autofit's padding and long-text bonus); sheetpeek fits a row to the lines it
draws (§4.7)."""


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    folder = tmp_path_factory.mktemp("wrap_band")
    return {column: suite.build_one(column, folder) for column in suite.COLUMN_KEYS}


def test_each_case_is_built_as_described(built):
    names = []
    for column, path in built.items():
        sheet = sheetpeek.read(path)
        for row, case in enumerate(suite.cases(column), 1):
            cell = sheet.cell(row, 1)
            assert (cell.font.name, cell.font.size) == case.face, case.name
            assert cell.value == case.text, case.name
            assert cell.style.wrap, case.name
            assert cell.style.indent == (1 if case.variant.startswith("indent") else 0)
            names.append(case.name)
    assert len(names) == len(set(names)) == 628


def test_each_column_is_its_width_on_its_platform(built):
    # A w column is its pixels on Windows, an m column its points on the Mac, in sheetpeek's
    # layout as in Excel's (both captures read every column's width back).
    for column, path in built.items():
        case = suite.cases(column)[0]
        rendered = sheetpeek.read(path, platform=case.platform).layout("excel")
        scale = 4 / 3 if case.platform == "windows" else 1
        assert rendered.columns[0].width_pt * scale == suite.column_size(column), column


def test_an_answer_reads_into_records(tmp_path):
    answer = tmp_path / "answer.tsv"
    lines = ["# excel: 16.115 (test), en_US", "# captured: 2026-11-25"]
    for column in suite.COLUMN_KEYS:
        name = suite.workbook_name(column)
        lines += [
            f"{name}\t{k}\t{c.face[0]}\t{c.face[1]}\t15.0\t8.43\t48.0"
            for k, c in enumerate(suite.cases(column), 1)
        ]
    answer.write_text("\n".join(lines) + "\n", encoding="utf-8")
    out = tmp_path / "wrap_band.tsv"
    suite.write(answer, out)
    rows = [r for r in out.read_text(encoding="utf-8").splitlines() if not r.startswith("#")]
    assert rows[0].split("\t") == list(suite.COLUMNS)
    assert len(rows) == 628 + 1
    with pytest.raises(ValueError, match="expected"):
        suite.records("\n".join(lines[:-1]))


@pytest.mark.excel
@pytest.mark.parametrize("platform", ["mac", "windows"])
def test_sheetpeek_fits_the_rows_excel_fits_but_the_band(built, platform):
    path = CAPTURES[platform]
    lines = [r for r in path.read_text(encoding="utf-8").splitlines() if not r.startswith("# ")]
    header = lines[0].split("\t")
    records = [dict(zip(header, line.split("\t"), strict=True)) for line in lines[1:]]
    assert len(records) == 628
    laid = {
        column: sheetpeek.read(p, platform=platform).layout("excel") for column, p in built.items()
    }
    taller, other = [], []
    for record in records:
        rendered = laid[record["column"]]
        assert float(record["width"]) == rendered.columns[0].width_pt, record["name"]
        ours = rendered.rows[int(record["row"]) - 1].height_pt
        excel = float(record["height"])
        if excel > ours:
            taller.append(record["name"])
        elif excel != ours:
            other.append((record["name"], excel, ours))
    REPORT[f"wrap_band ({platform})"] = (
        f"wrap_band, Excel for {'Mac' if platform == 'mac' else 'Windows'}: "
        f"{len(records) - len(taller)} of {len(records)} rows as Excel fits them, "
        f"{len(taller)} known deviations (fitted a line taller than drawn)"
    )
    assert other == []
    assert len(taller) == TALLER[platform]
