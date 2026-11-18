"""The characters layout against Excel's own layout, captured from its PDF export.

``tests/captures/layout.tsv`` records, per row of the capture workbook (``capture/layout.py``),
the column boundaries Excel drew no gridline across, which is how far each text spilled. The
same workbook is laid out here in ``characters`` mode and every spilling cell's reach is
compared. Characters mode works in whole characters where Excel works in points, so a few rows
differ by design; they are listed in ``KNOWN_DEVIATIONS`` with the reason, and the count is
reported, never skipped silently. The other captured facts are checked row by row: text, color
and alignment of the overflow rows, hidden and width-0 columns, fills, clipping cases.
"""

import io
import re
from pathlib import Path

import openpyxl
import pytest
import xlformat
from capture import package_probe
from capture.layout import build

import sheetpeek
from tests.capture_report import REPORT
from tests.watchdog import in_child
from tests.workbooks import (
    NS,
    SHEET1,
    link_shared_strings,
    replace_once,
    replace_part,
    replace_sheet_data,
    rewrite,
)

CAPTURE = Path(__file__).with_name("captures") / "layout.tsv"

# Case prefix → the column of the cell whose spill is compared. Only rows of zeros: a digit is
# one character and 6 pt, so reach in characters and in points can be compared directly.
ORIGINS = {"spill left zeros": "A", "spill right zeros": "H", "spill center zeros": "E"}
KNOWN_DEVIATIONS = {
    # A 53 pt column holds 8.83 digits, a characters column 8: four columns hold 35 digits in
    # characters mode, 34 in Excel.
    "spill left zeros=35": "a column holds 8.83 digits, not 8",
    "spill right zeros=35": "a column holds 8.83 digits, not 8",
    # Right-aligned text that would start exactly at a gridline's far edge spills in Excel; the
    # same text left-aligned, ending at a gridline, does not. Characters have no such edge.
    "spill right zeros=26": "Excel's one-point asymmetry for right-aligned text",
}


def records():
    lines = [line.rstrip("\n") for line in CAPTURE.read_text(encoding="utf-8").splitlines()]
    return [line.split("\t") for line in lines if line and not line.startswith("#")][1:]


def captured_rows():
    for number, case, missing, *_ in records():
        yield int(number), case, {int(k) for k in missing.split(",") if k}


def captured_baselines(row):
    """Per captured run, its baseline in points below the top edge of the row's top gridline."""
    record = next(r for r in records() if int(r[0]) == row)
    return [float(b) for b in record[5].split()] if len(record) > 5 else []


def captured_runs(row):
    """The text runs of a captured row: (text, x0, x1, color) in PDF points."""
    record = next(r for r in records() if int(r[0]) == row)
    runs = record[3] if len(record) > 3 else ""
    return [
        (text, float(x0), float(x1), color)
        for text, x0, x1, color in re.findall(r"'(.*?)'@(-?[\d.]+)-(-?[\d.]+)#([0-9A-F]{6})", runs)
    ]


def captured(case):
    return next((number, missing) for number, name, missing in captured_rows() if name == case)


@pytest.fixture(scope="module")
def sheet(tmp_path_factory):
    path = build(tmp_path_factory.mktemp("capture") / "layout.xlsx")
    # openpyxl saves N136's =1/0 without a cached value; write the one Excel saves (§4.8).
    rewrite(
        path,
        SHEET1,
        replace_once(
            '<c r="N136" s="6"><f>1/0</f><v /></c>',
            '<c r="N136" s="6" t="e"><f>1/0</f><v>#DIV/0!</v></c>',
        ),
    )
    return sheetpeek.read(path)


@pytest.fixture(scope="module")
def rendered(sheet):
    return sheet.layout("characters")


@pytest.fixture(scope="module")
def excel(sheet):
    return sheet.layout("excel")


def column(letter):
    return sheetpeek.CellRange.parse(f"{letter}1").min_col


@pytest.mark.excel
def test_spill_reach_matches_excel(rendered):
    visible = [c.letter for c in rendered.columns if not c.added]
    deviations = {}
    compared = 0
    for number, case, missing in captured_rows():
        prefix = case.partition("=")[0]
        if prefix not in ORIGINS:
            continue
        letter = ORIGINS[prefix]
        p = visible.index(letter)
        left = right = 0
        while p - left in missing:
            left += 1
        while p + right + 1 in missing:
            right += 1
        compared += 1
        ours = rendered.cell(number, column(letter)).spill
        if ours != (left, right):
            deviations[case] = (ours, (left, right))
    assert compared == 105
    REPORT["layout reach"] = (
        f"layout: {compared} spills of digits compared with Excel, "
        f"{len(deviations)} known deviations: "
        + "; ".join(f"{case} ({KNOWN_DEVIATIONS.get(case, 'UNEXPECTED')})" for case in deviations)
    )
    assert set(deviations) == set(KNOWN_DEVIATIONS), deviations


@pytest.mark.excel
def test_a_hidden_value_is_crossed(rendered):
    empty, with_value = captured("hidden K empty"), captured("hidden K holds a value")
    assert empty[1] == with_value[1]  # Excel: the same gridlines are left out
    assert (
        rendered.cell(empty[0], column("I")).spill
        == rendered.cell(with_value[0], column("I")).spill
    )
    number, missing = captured("hidden K holds a value, right")
    assert 10 in missing  # J|L, across the hidden K
    assert rendered.cell(number, column("M")).spill[0] >= 2


@pytest.mark.excel
def test_competing_spills(rendered):
    # Excel: both texts spill into the shared empty cell B and are drawn over each other, so the
    # gridlines on both sides of B are left out. Characters cannot overlap: the first spill
    # placed takes B, and the other is clipped (ARCHITECTURE.md §9).
    number, missing = captured("compete A left C right")
    assert {1, 2} <= missing
    assert rendered.cell(number, column("A")).spill == (0, 1)
    assert rendered.cell(number, column("C")).spill == (0, 0)
    competing = [name for _, name, _ in captured_rows() if name.startswith("compete ")]
    REPORT["layout competing"] = (
        f"layout: {len(competing)} rows of competing spills, known deviation by design "
        "(Excel overlaps both texts; characters give the cell to the first spill)"
    )


@pytest.mark.excel
def test_centered_text_blocked_on_one_side(rendered):
    # Excel: the blocked side is clipped, the other takes only what its own text needs.
    number, missing = captured("center zeros=12 blocked left")
    assert missing == {4}  # D|E only
    assert rendered.cell(number, column("D")).spill == (0, 1)
    for case, letter in [
        ("center blocked left: C value, D centered", "D"),
        ("center blocked right: D centered, E value", "D"),
    ]:
        number, missing = captured(case)
        left, right = rendered.cell(number, column(letter)).spill
        if "left:" in case:
            assert missing == {4, 5}
            assert left == 0
            assert right > 0
        else:
            assert missing == {2, 3}
            assert left > 0
            assert right == 0


@pytest.mark.excel
@pytest.mark.parametrize(
    "case",
    [
        "fill long, B empty",
        "justify long, B empty",
        "distributed long, B empty",
        "wrapped long, B empty",
    ],
)
def test_cells_that_do_not_spill(rendered, case):
    number, missing = captured(case)
    assert missing == set()
    assert rendered.cell(number, column("A")).spill == (0, 0)


@pytest.mark.excel
def test_a_line_break_without_wrap_is_not_drawn(rendered):
    number, missing = captured("line break without wrap")
    assert missing == {1}  # the joined text spills
    cell = rendered.cell(number, column("A"))
    assert cell.spill[1] > 0
    row = rendered.rows[number - 1]
    drawn = "".join(
        cell.lines[0] + (gap.lines[0] or "|")
        for cell, gap in zip(row.cells[:3], row.gaps[:3], strict=True)
    )
    assert drawn.startswith("first linesecond line")


@pytest.mark.excel
def test_a_zero_width_column_is_hidden(rendered):
    assert "O" not in [c.letter for c in rendered.columns]
    empty, with_value = (
        captured("width-0 O empty, M spills"),
        captured("width-0 O holds a value, M spills"),
    )
    assert empty[1] == with_value[1]
    assert (
        rendered.cell(empty[0], column("M")).spill
        == rendered.cell(with_value[0], column("M")).spill
    )


@pytest.mark.excel
def test_fills_under_merges_and_spills(rendered):
    # Excel paints the anchor's fill over the whole merge and ignores a covered cell's own fill;
    # spilled text is drawn over a neighbor's fill, and the spilling cell's fill stays home.
    yellow = "FFFF00"
    number, _ = captured("merge fill A:C: anchor yellow, covered B blue")
    assert {c.fill for c in rendered.rows[number - 1].cells[:3]} == {yellow}
    number, _ = captured("merge fill A:C: anchor none, covered B blue")
    assert {c.fill for c in rendered.rows[number - 1].cells[:3]} == {None}
    number, _ = captured("spill over fill: A text, B yellow")
    b = rendered.cell(number, column("B"))
    assert (b.fill, b.source) == (yellow, (number, 1))
    number, _ = captured("spill from a filled cell: A yellow text, B empty")
    b = rendered.cell(number, column("B"))
    assert (b.fill, b.source) == (None, (number, 1))


# Left gridlines in the capture: A at 18 pt, then 53 pt per default column, K hidden, N 22 pt.
COLUMN_X = {"A": (18, 53), "M": (601, 53), "N": (654, 22)}


@pytest.mark.excel
@pytest.mark.parametrize("row", range(130, 139))
def test_overflow_color_and_alignment(rendered, row):
    # Rows 130 to 138: the # run in the section color or the font color, General's reduction
    # under "@", a number under "@" left-aligned and under General right-aligned.
    ((text, x0, x1, color),) = captured_runs(row)
    letter = next(k for k in COLUMN_X if any(
        c.column == column(k) and c.cell.kind != "empty" for c in rendered.rows[row - 1].cells
    ))  # fmt: skip
    cell = rendered.cell(row, column(letter))
    assert cell.text == text
    assert (cell.color or "000000") == color
    if "#" not in text:
        left, width = COLUMN_X[letter]
        if x0 == left + 3:  # starts at the left inset
            assert cell.align == "left"
        else:
            assert abs(x1 - (left + width - 2)) < 1  # ends at the right inset
            assert cell.align == "right"


# excel mode ----------------------------------------------------------------------------------

PDF_X = 18.0  # the PDF x of column A's left gridline
PRINT_PITCH = 14.0  # rows 15 pt in the file print 14 pt apart
CAPTURE_COLUMNS = 14  # visible columns in the capture's print area: A to J, L, M, N, P
KNOWN_BOUNDARIES = {
    # Rows 180 and 184 draw the same 49.79 pt text; only the one holding a line feed spills, so a
    # line feed adds to the width that decides spill. Rows 195 and 197 show it adds under 2 pt.
    "escaped CR LF without wrap": "a line feed's small width, which layout drops with it",
}


def excel_runs(excel, number):
    """The runs drawn in a sheet row: its own, and a vertical merge's from the rows above."""
    row = excel.rows[number - 1]
    runs = list(row.runs)
    for above in excel.rows[: number - 1]:
        runs += [r for r in above.runs if r.clip[3] > row.y_pt and r.y > row.y_pt]
    return runs


@pytest.mark.excel
def test_excel_boundaries_match_excel(excel):
    # Every row: the gridlines Excel left out are the boundaries text or a merge crosses, plus
    # those a fill covers (a fill covers its own gridlines, CAPTURING.md).
    deviations = {}
    for number, case, missing in captured_rows():
        row = excel.rows[number - 1]
        fills = {
            k + 1 for k in range(len(row.cells) - 1) if row.cells[k].fill or row.cells[k + 1].fill
        }
        crossed = {k + 1 for k, c in enumerate(row.crossed) if c}
        ours = (crossed | fills) & set(range(1, CAPTURE_COLUMNS))
        if ours != missing:
            deviations[case] = (sorted(ours), sorted(missing))
    REPORT["excel boundaries"] = (
        f"excel: {len(excel.rows)} rows' boundaries compared with Excel, "
        f"{len(deviations)} known deviations: "
        + "; ".join(f"{case} ({KNOWN_BOUNDARIES.get(case, 'UNEXPECTED')})" for case in deviations)
    )
    assert set(deviations) == set(KNOWN_BOUNDARIES), deviations


KNOWN_RUN_DEVIATIONS: dict[str, str] = {}
"""Runs whose start or end differs from Excel's, by case, with the reason (none now)."""


@pytest.mark.excel
def test_excel_runs_match_excel(excel):
    # Each captured run whose text is one of ours: the start to the point, the end within a
    # point, the color. pdfminer joins and splits runs where texts overlap or are clipped; a
    # captured run that is a prefix of ours is compared by its start only, others not at all.
    # Competing spills interleave two texts' glyphs; their rows are checked by the boundaries
    # and the clips instead.
    compared = starts = 0
    deviations = {}
    for number, case, _ in captured_rows():
        if case.startswith("compete "):
            continue
        ours = excel_runs(excel, number)
        for text, x0, x1, color in captured_runs(number):
            same = [r for r in ours if r.text == text.rstrip(" ")]
            prefix = [
                r for r in ours if len(text) > 1 and r.text.startswith(text) and r.text != text
            ]
            for run in same[:1] or prefix[:1]:
                starts += 1
                if abs(run.x + PDF_X - x0) > 0.01:
                    deviations[f"{case}: start of {text!r}"] = (run.x + PDF_X, x0)
            for run in same[:1]:
                compared += 1
                end = run.x + run.width + PDF_X
                if text == run.text and abs(end - x1) > 1:
                    deviations[f"{case}: end of {text!r}"] = (end, x1)
                if (run.color or "000000") != color:
                    deviations[f"{case}: color of {text!r}"] = (run.color, color)
    known = {k: v for k, v in deviations.items() if k.split(":")[0] in KNOWN_RUN_DEVIATIONS}
    REPORT["excel runs"] = (
        f"excel: {starts} run starts and {compared} runs compared with Excel, "
        f"{len(deviations) - len(known)} deviations, {len(known)} known: "
        + "; ".join(f"{case} ({why})" for case, why in KNOWN_RUN_DEVIATIONS.items())
    )
    assert starts >= 150
    assert {k: v for k, v in deviations.items() if k not in known} == {}


@pytest.mark.excel
def test_excel_baselines(excel):
    # Print puts a bottom-aligned line 3 pt above the next row's gridline and a top-aligned one
    # (or one too tall for its row) 11 pt below the top. The layout follows the screen, which
    # puts the top-aligned line 1 pt lower (suite text_style, ARCHITECTURE.md §3.6).
    checked = 0
    for number, _, _ in captured_rows():
        row = excel.rows[number - 1]
        for (text, x0, _, _), baseline in zip(
            captured_runs(number), captured_baselines(number), strict=True
        ):
            for run in row.runs:
                if (
                    run.source[0] == number
                    and run.text == text.rstrip(" ")
                    and abs(run.x + PDF_X - x0) < 0.01
                ):
                    checked += 1
                    from_bottom = row.y_pt + row.height_pt - run.y
                    assert from_bottom == PRINT_PITCH - baseline or run.y - row.y_pt == baseline + 1
    assert checked > 150
    # A merge over three rows (layout rows 152 to 154): top, center and bottom alignment. The
    # capture's merge is 42 pt high in print and the layout's 45 pt.
    runs = {r.text: r.y - excel.rows[151].y_pt for r in excel.rows[151].runs}
    assert captured_baselines(152) == [11.0]
    assert runs["top"] == 11 + 1
    assert captured_baselines(153) == [10.0]  # 14 + 10 = 24 = floor((42 - 15) / 2) + 11
    assert runs["center"] == (45 - 3 - 12) // 2 + 12  # midway between 12 and 42
    assert captured_baselines(154) == [11.0, 11.0]  # 28 + 11 = 39 = 42 - 3
    assert runs["bottom"] == runs["default"] == 45 - 3


@pytest.mark.excel
def test_excel_clipping_and_competing_spills(excel):
    def clip(case, letter):
        number, _ = captured(case)
        (run,) = [r for r in excel.rows[number - 1].runs if r.source[1] == column(letter)]
        return run.clip[0], run.clip[2]

    x = {c.letter: c.x_pt for c in excel.columns}
    # Blocked text is cut at its cell's edge (rendered page, rows 120 to 122).
    assert clip("blocked left", "A") == (x["A"] + 1, x["B"])
    assert clip("blocked right", "D") == (x["D"] + 1, x["E"])
    assert clip("blocked center", "D") == (x["D"] + 1, x["E"])
    # Both competing spills take B, and are drawn over each other (rows 123 to 126).
    number, _ = captured("compete A left C right")
    a, c = sorted(excel.rows[number - 1].runs, key=lambda r: r.source[1])
    assert a.clip[2] > x["B"]
    assert c.clip[0] < x["C"]
    assert excel.cell(number, column("A")).spill == (0, 1)
    assert excel.cell(number, column("C")).spill == (1, 0)
    # A # run is centered in its cell, in the section color (rows 130 to 136).
    for number in (130, 131, 132, 133, 135, 136):  # 134 is 1.23E+09
        (run,) = excel.rows[number - 1].runs
        assert run.text == "###"


# theme colors --------------------------------------------------------------------------------

THEME_CAPTURE = Path(__file__).with_name("captures") / "theme_colors.tsv"


@pytest.mark.excel
def test_theme_colors_match_excel(tmp_path):
    # Every theme index 0 to 11 with eight tints, as a fill and as a font color. The font color
    # of theme 1 without tint ("Text 1") is automatic by design (ARCHITECTURE.md §3.3).
    from capture.theme_colors import build

    sheet = sheetpeek.read(build(tmp_path / "theme_colors.xlsx"))
    lines = THEME_CAPTURE.read_text(encoding="utf-8").splitlines()
    records = [line.split("\t") for line in lines if line and not line.startswith("#")][1:]
    deviations = {}
    for address, theme, tint, fill, font in records:
        position = sheetpeek.CellRange.parse(address)
        cell = sheet.cell(position.min_row, position.min_col)
        automatic = (theme, float(tint)) == ("1", 0.0)
        if cell.fill != fill or cell.font.color != (None if automatic else font):
            deviations[address] = (cell.fill, fill, cell.font.color, font)
    assert len(records) == 96
    REPORT["theme colors"] = (
        f"theme colors: {len(records)} cells compared with Excel, {len(deviations)} deviations"
    )
    assert deviations == {}


# merged widths and error values ---------------------------------------------------------------


def suite(name):
    lines = (Path(__file__).with_name("captures") / name).read_text(encoding="utf-8").splitlines()
    return [line.split("\t") for line in lines if line and not line.startswith("# ")][1:]


@pytest.mark.excel
def test_merge_width_matches_excel(tmp_path):
    # A merged anchor is formatted at the merged width (the visible columns' points, converted
    # back to dialog units), hidden columns left out: what Excel shows, read with string value.
    from capture.cell_text import build

    book = sheetpeek.open(build(tmp_path / "cell_text.xlsx"))
    rows: dict[tuple[str, bool], list[str]] = {}
    for width, merge, _value, _fmt, shown in suite("merge_width.tsv"):
        rows.setdefault((width, merge == "A:C, B hidden"), []).append(shown)
    compared, deviations = 0, {}
    for (width, hidden), texts in rows.items():
        rendered = book.sheet(("h" if hidden else "w") + width).layout("excel")
        for row, shown in enumerate(texts, 1):
            compared += 1
            if rendered.cell(row, 1).text != shown:
                deviations[(width, hidden, row)] = (rendered.cell(row, 1).text, shown)
    REPORT["merge width"] = (
        f"merge width: {compared} cells compared with Excel, {len(deviations)} deviations"
    )
    assert compared == 150
    assert deviations == {}


@pytest.mark.excel
def test_error_width_matches_excel(tmp_path):
    from capture.cell_text import ERROR_WIDTHS, ERRORS, build

    rendered = sheetpeek.open(build(tmp_path / "cell_text.xlsx")).sheet("errors").layout("excel")
    records = suite("error_width.tsv")
    deviations = {}
    for error, width, shown in records:
        cell = rendered.cell(ERRORS.index(error) + 1, ERROR_WIDTHS.index(float(width)) + 2)
        if cell.text != shown:
            deviations[(error, width)] = (cell.text, shown)
    REPORT["error width"] = (
        f"error width: {len(records)} cells compared with Excel, {len(deviations)} deviations"
    )
    assert len(records) == 56
    assert deviations == {}


# text -----------------------------------------------------------------------------------------


def shared_string(tmp_path, xml):
    """A workbook whose A1 is the shared string given as the content of an ``<si>``."""
    wb = openpyxl.Workbook()
    wb.active["A1"] = "placeholder"
    path = tmp_path / "strings.xlsx"
    wb.save(path)
    link_shared_strings(path)
    replace_part(path, "xl/sharedStrings.xml", f"<sst {NS}><si>{xml}</si></sst>".encode())
    rewrite(
        path,
        SHEET1,
        replace_sheet_data('<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>'),
    )
    return sheetpeek.read(path).cell(1, 1).value


@pytest.mark.excel
def test_shared_strings_read_as_excel_shows_them(tmp_path):
    # Excel shows the literal _x0041_ (layout row 191, through an inline string) and drops the
    # spaces of a <t> without xml:space (xlrecalc's capture `whitespace`).
    got = {
        "shared string _x005F_x0041_": shared_string(tmp_path, "<t>_x005F_x0041_</t>"),
        "shared string x005F_keep": shared_string(tmp_path, "<t>x005F_keep</t>"),
        "whitespace without xml:space": shared_string(tmp_path, "<t>  pad  </t>"),
    }
    assert got == {
        "shared string _x005F_x0041_": "_x0041_",
        "shared string x005F_keep": "x005F_keep",
        "whitespace without xml:space": "pad",
    }
    REPORT["text"] = f"text: {len(got)} shared strings compared with Excel, 0 deviations"


@pytest.mark.excel
def test_a_lone_carriage_return_does_not_break_a_line(excel):
    # Rows 182, 186, 188: Excel breaks at a line feed, with or without a carriage return before
    # it. A carriage return alone is drawn as nothing, and the word then breaks by width.
    runs = {number: [r.text for r in excel_runs(excel, number)] for number in (182, 186, 188)}
    assert runs[182][0] == "firstsecon"
    assert runs[186] == ["first", "second"]
    assert runs[188] == ["first", "second"]


@pytest.mark.excel
def test_escapes_decode_as_excel_shows_them(excel):
    expected = {190: "A", 191: "_x0041_", 192: "J", 193: "_x041_"}
    for number, text in expected.items():
        assert [r.text for r in excel_runs(excel, number)] == [text]


@pytest.mark.excel
def test_control_characters_are_not_drawn(excel):
    for number in (198, 199, 200):
        assert [r.text for r in excel_runs(excel, number)] == ["ab"]


# packages ------------------------------------------------------------------------------------

PACKAGE_CAPTURE = Path(__file__).with_name("captures") / "package.tsv"
PACKAGE_DEVIATIONS = {
    # Excel gives a row of font size 1 a height of 4 pt; sizes under 9 are outside the row-height
    # table (ARCHITECTURE.md §4.7).
    "font_size_negative": "a font size outside the row-height table",
}


def package_rows():
    for line in PACKAGE_CAPTURE.read_text(encoding="utf-8").splitlines():
        if not line.startswith(("#", "case\t")):
            name, outcome, *sizes = line.split("\t")
            yield name, outcome, sizes


def _points(width, font):
    return xlformat.column_metrics(float(width), font).points if float(width) else 0.0


def compare_packages():
    """Every case against the capture, in a child: some of them would be bombs without a guard."""
    deviations = {}
    for name, outcome, sizes in package_rows():
        data = io.BytesIO(dict(package_probe.CASES)[name](package_probe.base()))
        try:
            sheet = sheetpeek.read(data, ref="A1:B2")
        except sheetpeek.PackageError as error:
            if outcome != "refused":
                deviations[name] = f"refused, where Excel opens it: {error}"
            continue
        if outcome == "refused":
            deviations[name] = "read, where Excel refuses it"
            continue
        rendered = sheet.layout("excel", show_hidden=True)
        sheetpeek.render(rendered, "html")
        sheet.layout("characters")
        sheet.layout("none")
        _, _, width_a, height_1, width_b, height_2, font_size = sizes
        font = xlformat.Font(sheet.normal_font.name, sheet.normal_font.size)
        ours = (
            sheet.cell(1, 1).value,
            sheet.cell(1, 1).style.font.size,
            rendered.columns[0].width_pt,
            rendered.columns[1].width_pt,
        )
        expected = (111, float(font_size), _points(width_a, font), _points(width_b, font))
        if ours != expected:
            deviations[name] = (ours, expected)
        heights = (rendered.rows[0].height_pt, rendered.rows[1].height_pt)
        if heights != (float(height_1), float(height_2)) and name not in PACKAGE_DEVIATIONS:
            deviations[name] = (heights, (height_1, height_2))
    return deviations


@pytest.mark.excel
def test_every_package_case_is_captured():
    assert [name for name, _, _ in package_rows()] == [name for name, _ in package_probe.CASES]


@pytest.mark.excel
def test_a_package_is_taken_as_excel_takes_it():
    deviations = in_child(compare_packages)
    REPORT["package"] = (
        f"package: {len(package_probe.CASES)} cases compared with Excel, "
        f"{len(PACKAGE_DEVIATIONS)} known deviations: "
        + "; ".join(f"{case} ({why})" for case, why in PACKAGE_DEVIATIONS.items())
    )
    assert deviations == {}
