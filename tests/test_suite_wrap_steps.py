"""The capture suite ``wrap_steps`` (fitted wrapped rows at autofit's later bonus steps and under
a Courier 10 Normal font, ARCHITECTURE.md §4.7), without Excel.

The workbooks are built and laid out with sheetpeek, whose row heights are compared with Excel's
in ``tests/captures/wrap_steps.tsv`` (Excel for Mac) and ``tests/captures/windows/wrap_steps.tsv``
(Excel for Windows) once they exist. The reading of Excel's answer runs on a synthetic answer.
"""

import pytest
from capture import wrap_steps as suite

import sheetpeek
from tests.capture_report import REPORT

pytest.importorskip("openpyxl")

CAPTURES = {
    "mac": suite.OUTPUT,
    "windows": suite.ROOT / "tests" / "captures" / "windows" / "wrap_steps.tsv",
}
KNOWN: dict[str, set[str]] = {"mac": set(), "windows": set()}
"""Cases whose row sheetpeek fits otherwise than Excel, by platform, with the reason."""


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    folder = tmp_path_factory.mktemp("wrap_steps")
    return {column.key: suite.build_one(column, folder) for column in suite.COLUMN_LIST}


def test_each_case_is_built_as_described(built):
    names = []
    for column in suite.COLUMN_LIST:
        sheet = sheetpeek.read(built[column.key])
        assert (sheet.normal_font.name, sheet.normal_font.size) == column.normal
        for row, case in enumerate(suite.cases(column), 1):
            cell = sheet.cell(row, 1)
            assert (cell.font.name, cell.font.size) == column.cell, case.name
            assert cell.value == case.text, case.name
            assert cell.style.wrap, case.name
            names.append(case.name)
    assert len(names) == len(set(names)) == 388


def test_each_text_is_its_target_wide_on_its_platform():
    import xlformat

    for column in suite.COLUMN_LIST:
        scale = 4 / 3 if column.platform == "windows" else 1
        for case in suite.cases(column):
            width = xlformat.text_width(
                case.text, xlformat.Font(*column.cell), platform=column.platform
            )
            assert round(width * scale, 9) == case.target, case.name


def test_an_answer_reads_into_records(tmp_path):
    answer = tmp_path / "answer.tsv"
    lines = ["# excel: 16.115 (test), en_US", "# captured: 2026-11-25"]
    for column in suite.COLUMN_LIST:
        lines += [
            f"{column.workbook}\t{k}\tArial\t10\t15.0\t8.43\t48.0"
            for k in range(1, len(suite.cases(column)) + 1)
        ]
    answer.write_text("\n".join(lines) + "\n", encoding="utf-8")
    out = tmp_path / "wrap_steps.tsv"
    suite.write(answer, out)
    rows = [r for r in out.read_text(encoding="utf-8").splitlines() if not r.startswith("#")]
    assert rows[0].split("\t") == list(suite.COLUMNS)
    assert len(rows) == 388 + 1
    with pytest.raises(ValueError, match="expected"):
        suite.records("\n".join(lines[:-1]))


@pytest.mark.excel
@pytest.mark.parametrize("platform", ["mac", "windows"])
def test_sheetpeek_fits_the_rows_excel_fits(built, platform):
    path = CAPTURES[platform]
    if not path.exists():
        pytest.skip(f"not captured yet: {path.name} on {platform}")
    lines = [r for r in path.read_text(encoding="utf-8").splitlines() if not r.startswith("# ")]
    header = lines[0].split("\t")
    records = [dict(zip(header, line.split("\t"), strict=True)) for line in lines[1:]]
    assert len(records) == 388
    laid = {key: sheetpeek.read(p, platform=platform).layout("excel") for key, p in built.items()}
    differ = {
        r["name"]
        for r in records
        if laid[r["column"]].rows[int(r["row"]) - 1].height_pt != float(r["height"])
    }
    REPORT[f"wrap_steps ({platform})"] = (
        f"wrap_steps, Excel for {'Mac' if platform == 'mac' else 'Windows'}: "
        f"{len(records) - len(differ)} of {len(records)} rows as Excel fits them, "
        f"{len(differ)} known deviations"
    )
    assert differ == KNOWN[platform]
