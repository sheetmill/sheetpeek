"""The screen suite ``extensions`` (sheet extension URIs and case, ARCHITECTURE.md §3.7), no Excel.

Its workbooks are built and read back with sheetpeek as built, and what sheetpeek takes each
extension for is the suite's ``expected``, from sheetpeek's own table. Its measuring runs on a
synthetic capture with known marks, and its reading of a saved sheet on hand-written XML. The
comparison with ``tests/captures/extensions.tsv`` is skipped until that capture exists.
"""

import zipfile

import openpyxl
import pytest
from capture import extensions
from capture.screen.window import MARKER
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

import sheetpeek
from sheetpeek import _detect
from tests.screen_canvas import SCALE, Canvas


def _rgb(text):
    return tuple(int(text[k : k + 2], 16) for k in (0, 2, 4))


def _distance(a, b):
    return sum((x - y) ** 2 for x, y in zip(_rgb(a), _rgb(b), strict=True)) ** 0.5


# --- the workbooks


def test_every_group_builds_and_reads_back_as_built(tmp_path):
    placed = extensions.build(tmp_path)
    assert len(placed) == len(extensions.GROUPS) * len(extensions.CASES)
    for group, chosen in extensions.GROUPS.items():
        path = tmp_path / f"extensions_{group}.xlsx"
        sheet = sheetpeek.read(path)
        assert sheet.name == extensions.SHEET
        with zipfile.ZipFile(path) as package:
            xml = package.read("xl/worksheets/sheet1.xml").decode()
        assert xml.count("<ext ") == len(chosen), group
        for uri in chosen.values():
            assert f'<ext uri="{uri}"' in xml, group
        assert extensions.IGNORED in xml
        codes = {n.code for n in sheetpeek.notices(path, "html")}
        # what sheetpeek takes each extension for: the sparkline reported, the rest silent; a
        # re-cased URI draws nothing and is not reported, as Excel ignores it
        took = extensions.expected(group)
        assert ("sparkline" in codes) == took.get("sparkline", False), group
        assert "sheet.unknown_extension" not in codes, group


def test_the_cells_hold_what_the_cases_say(tmp_path):
    extensions.build(tmp_path)
    sheet = sheetpeek.read(tmp_path / "extensions_exact.xlsx")

    def value(address):
        letters, row = coordinate_from_string(address)
        return sheet.cell(row, column_index_from_string(letters)).value

    assert [value(f"{c}4") for c in "CDEFGH"] == [1, 2, 3, 4, 5, 6]
    assert {value(c) for c in extensions.NUMBER_TEXT} == {"123"}
    # nothing but the sparkline in J4: the used range stops at column H
    assert column_index_from_string(extensions.SPARKLINE[0]) > sheet.ref.max_col


def test_sheetpeek_matches_the_captured_uris_exactly():
    assert set(extensions.URIS.values()) <= set(_detect._URIS_EXACT)
    assert extensions.expected("exact") == dict.fromkeys(extensions.URIS, True)
    for kind in extensions.URIS:
        assert extensions.expected(f"{kind}_lower") == {kind: False}


def test_the_sheet_must_end_as_openpyxl_writes_it():
    with pytest.raises(ValueError, match="pageMargins"):
        extensions.sheet("<worksheet><drawing r:id='x'/></worksheet>", "exact")
    with pytest.raises(ValueError, match="does not end"):
        extensions.sheet("<worksheet/>", "exact")


def test_no_case_draws_near_the_marker_color():
    for color in (extensions.SERIES, extensions.OTHER, "008000", "000000"):
        assert _distance(color, MARKER) > 200


def test_every_cell_is_its_own_and_the_markers_lie_beyond():
    cells = [coordinate_from_string(c) for c in extensions.CASES]
    cells += [coordinate_from_string(c) for c in (extensions.VALIDATION, extensions.PROTECTED)]
    assert len(set(cells)) == len(cells)
    letters, row = coordinate_from_string(extensions.MARKERS[1])
    assert row > max(r for _, r in cells)
    assert column_index_from_string(letters) > max(column_index_from_string(c) for c, _ in cells)


# --- measuring, on a synthetic capture


def _canvas():
    ws = openpyxl.Workbook().active
    extensions.fill(ws)
    return Canvas(ws, list(extensions.CASES))


def _records(cap):
    return {r["cell"]: r for r in extensions.measure(cap)}


def test_a_blank_capture_reads_no_sparkline_and_no_triangle():
    found = _records(_canvas().capture())
    assert found["J4"]["ink_pixels"] == "0"
    assert found["J4"]["ink_color"] == "none"
    for cell in (*extensions.NUMBER_TEXT, *extensions.ERRORS):
        assert (found[cell]["triangle_pixels"], found[cell]["triangle_color"]) == ("0", "none")


def test_column_bars_read_as_the_sparkline_ink_and_its_box():
    canvas = _canvas()
    x0, y0, _, y1 = canvas.box("J4")
    series = _rgb(extensions.SERIES)
    for k in range(6):  # six bars 10 px wide, 10 px apart, rising
        left = x0 + 10 + 20 * k
        canvas.paint(left, y1 - 10 - 6 * (k + 1), left + 10, y1 - 10, series)
    canvas.paint(x0 - 1, y0, x0 + 1, y1, (212, 212, 212))  # a gridline is not ink
    found = _records(canvas.capture())["J4"]
    assert found["ink_pixels"] == str(sum(10 * 6 * (k + 1) for k in range(6)))
    assert found["ink_color"] == extensions.SERIES
    assert found["ink_w"] == f"{110 / SCALE:g}"
    assert found["ink_h"] == f"{36 / SCALE:g}"


def test_a_green_corner_reads_as_the_triangle_only_in_its_cell():
    canvas = _canvas()
    x0, y0, _, _ = canvas.box("C7")
    for k in range(8):  # a right triangle 8 px on a side in the top left corner
        canvas.paint(x0 + 1, y0 + 1 + k, x0 + 1 + 8 - k, y0 + 2 + k, (0, 128, 0))
    canvas.paint(x0 + 4, y0 + 12, x0 + 20, y0 + 24, (0, 0, 0))  # text is not green
    found = _records(canvas.capture())
    assert found["C7"]["triangle_pixels"] == str(sum(8 - k for k in range(8)))
    assert found["C7"]["triangle_color"] == "008000"
    assert found["E7"]["triangle_pixels"] == "0"


# --- reading a saved sheet


_SHEET = (
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
    'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
    "<sheetData/>{main}<extLst>{exts}</extLst></worksheet>"
)


def _saved(main="", exts=""):
    return _SHEET.format(main=main, exts=exts).encode()


def test_a_saved_sheet_tells_where_each_kind_went():
    uris = extensions.URIS
    exts = (
        f'<ext uri="{uris["sparkline"]}"><x14:sparklineGroups><x14:sparklineGroup/>'
        "</x14:sparklineGroups></ext>"
        f'<ext uri="{uris["protected"].lower()}"><x14:protectedRanges>'
        '<x14:protectedRange name="ExtRange"/></x14:protectedRanges></ext>'
    )
    main = (
        '<dataValidations count="1"><dataValidation type="list" sqref="C11">'
        "<formula1>Lists!$A$1:$A$3</formula1></dataValidation></dataValidations>"
        '<ignoredErrors><ignoredError sqref="E7" numberStoredAsText="1"/></ignoredErrors>'
    )
    text = _saved(main, exts)
    listed = f"{uris['sparkline']} {uris['protected'].lower()}"
    assert extensions.saved(text, "sparkline") == ("exact", listed)
    assert extensions.saved(text, "validation") == ("main", listed)
    assert extensions.saved(text, "protected") == (f"other:{uris['protected'].lower()}", listed)
    assert extensions.saved(text, "ignored") == ("", listed)  # only E7, the sheet's own
    moved = _saved('<ignoredErrors><ignoredError sqref="G9 E9" evalError="1"/></ignoredErrors>')
    assert extensions.saved(moved, "ignored") == ("main", "")
    kept = _saved(
        exts=f'<ext uri="{uris["ignored"]}"><x14:ignoredErrors><x14:ignoredError>'
        "<xm:sqref>G7</xm:sqref></x14:ignoredError></x14:ignoredErrors></ext>"
    )
    assert extensions.saved(kept, "ignored") == ("exact", uris["ignored"])


def test_a_package_gives_its_first_sheet(tmp_path):
    extensions.build(tmp_path)
    path = tmp_path / "extensions_exact.xlsx"
    with zipfile.ZipFile(path) as package:
        assert extensions.first_sheet(path) == package.read("xl/worksheets/sheet1.xml")


def test_saved_captures_become_rows(tmp_path, monkeypatch):
    monkeypatch.setattr(extensions, "BUILD_DIR", tmp_path / "build")
    monkeypatch.setattr(extensions, "OUTPUT", tmp_path / "extensions.tsv")
    exact = tmp_path / "build" / "exact"
    exact.mkdir(parents=True)
    (exact / "saved_sheet.xml").write_bytes(_saved())
    refused = tmp_path / "build" / "ignored_lower"
    refused.mkdir()
    extensions.record_refusal("ignored_lower", "osascript: execution error: refused (-50)")
    extensions.write()
    text = (tmp_path / "extensions.tsv").read_text(encoding="utf-8").splitlines()
    assert text[:2] == ["# suite: extensions", "\t".join(extensions.COLUMNS)]
    rows = [dict(zip(extensions.COLUMNS, line.split("\t"), strict=True)) for line in text[2:]]
    assert [(r["kind"], r["group"], r["extension"]) for r in rows] == [
        ("save", "exact", kind) for kind in extensions.GROUPS["exact"]
    ] + [("refused", "ignored_lower", "")]
    assert rows[-1]["case"] == "refused: refused (-50)"


# --- the capture (suite extensions)


def _captured():
    path = extensions.OUTPUT
    if not path.exists():
        pytest.skip("not captured yet: run python -m capture.extensions run")
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header = rows[0]
    return [dict(zip(header, row, strict=False)) for row in rows[1:]]


def _screen(records, group):
    return {
        r["cell"]: r
        for r in records
        if r["kind"] == "screen" and r["group"] == group and r["zoom"] == str(extensions.ZOOM)
    }


# Tolerances. The measured cells hold nothing but what is measured: an empty J4 reads 0 ink
# pixels and an unmarked corner 0 green ones, as the synthetic tests show, so the thresholds
# only have to clear anti-aliasing. Six column bars in a 109 by 30 pt cell cover thousands of
# device pixels; 100 is far below them and far above stray edge pixels. Excel's error indicator
# is a triangle about 4 pt on a side, about 30 device pixels at 2 a point; 6 is a fifth of it.
# The bars' color is compared to the series color within 60: anti-aliased edges pull the mean
# toward white, and the profile conversion moves a color by about 10.
SPARKLINE_DRAWN = 100
TRIANGLE_DRAWN = 6
SERIES_COLOR = 60


def _drawn(record):
    return int(record.get("ink_pixels") or 0) >= SPARKLINE_DRAWN


def _marked(record):
    return int(record.get("triangle_pixels") or 0) >= TRIANGLE_DRAWN


@pytest.mark.excel
@pytest.mark.xfail(
    strict=True,
    reason="16.115: the exact group's sparkline drew nothing and the extension's exemption left "
    "G7 and G9 marked, though Excel saved all four extensions as their kinds; the screen step "
    "settles nothing and the save step decides (ARCHITECTURE §3.7)",
)
def test_the_controls_show_what_the_exact_uris_draw():
    exact = _screen(_captured(), "exact")
    if not exact:
        pytest.skip("the exact group's window was not captured")
    assert _drawn(exact["J4"]), exact["J4"]
    assert _distance(exact["J4"]["ink_color"], extensions.SERIES) <= SERIES_COLOR, exact["J4"]
    for cell in ("C7", "C9"):
        assert _marked(exact[cell]), ("error checking marks no error", cell, exact[cell])
    for cell in ("E7", "E9"):
        assert not _marked(exact[cell]), ("the sheet's ignoredErrors not honored", cell)
    for cell in ("G7", "G9"):
        # if this fails, Excel does not honor these attributes in the extension, and the
        # ignored-errors cases cannot settle its URI's case: say so in ARCHITECTURE §3.7
        assert not _marked(exact[cell]), ("the extension's exemption not honored", cell)


@pytest.mark.excel
def test_every_exact_uri_is_read_and_written_back():
    found = {
        r["extension"]: r for r in _captured() if r["kind"] == "save" and r["group"] == "exact"
    }
    if not found:
        pytest.skip("the exact group was not saved")
    for kind, record in found.items():
        assert record["saved_where"] in ("main", "exact"), (kind, record["saved_where"])


def _excel_reads(records, group, kind):
    """Whether Excel read ``group``'s extension of ``kind``: from the screen, else the save."""
    screen = _screen(records, group)
    if kind == "sparkline" and screen:
        return _drawn(screen["J4"])
    if kind == "ignored" and screen:
        return not _marked(screen["G7"]) and not _marked(screen["G9"])
    saved = [r for r in records if r["kind"] == "save" and r["group"] == group]
    if not saved:
        return None
    return saved[0]["saved_where"] in ("main", "exact")


@pytest.mark.excel
def test_sheetpeek_matches_each_uri_as_excel_does():
    records = _captured()
    refused = {r["group"] for r in records if r["kind"] == "refused"}
    settled = 0
    for kind in extensions.URIS:
        group = f"{kind}_lower"
        if group in refused:
            continue  # Excel refused the file: the record is the answer, nothing to compare
        reads = _excel_reads(records, group, kind)
        if reads is None:
            continue
        settled += 1
        # sheetpeek takes a re-cased URI as its kind exactly when Excel reads it; where Excel
        # ignores it, compare that URI exactly, as conditional formatting's
        assert extensions.expected(group)[kind] == reads, (
            kind,
            "Excel reads it" if reads else "Excel ignores the lower-case URI",
        )
    if not settled:
        pytest.skip("no lower-case group captured")


@pytest.mark.excel
def test_the_screen_and_the_save_agree():
    records = _captured()
    for group in extensions.SCREEN[1:]:
        kind = group.removesuffix("_lower")
        screen = _screen(records, group)
        saved = [r for r in records if r["kind"] == "save" and r["group"] == group]
        if not screen or not saved:
            continue
        read = saved[0]["saved_where"] in ("main", "exact")
        assert _excel_reads(records, group, kind) == read, (group, saved[0]["saved_where"])


def _agreeing():
    lines = ["# suite: extensions", "\t".join(extensions.COLUMNS)]

    def row(**fields):
        lines.append("\t".join(fields.get(c, "") for c in extensions.COLUMNS))

    # Excel reads the exact URIs and ignores the lower-case ones (captured, 16.115).
    for group in extensions.SCREEN:
        drawn = group != "sparkline_lower"
        row(kind="screen", group=group, cell="J4", zoom="100",
            ink_pixels="2400" if drawn else "0",
            ink_color=extensions.SERIES if drawn else "none")  # fmt: skip
        for cell in (*extensions.NUMBER_TEXT, *extensions.ERRORS):
            marked = cell[0] == "C" or (group == "ignored_lower" and cell[0] == "G")
            row(kind="screen", group=group, cell=cell, zoom="100",
                triangle_pixels="30" if marked else "0",
                triangle_color="008000" if marked else "none")  # fmt: skip
    for group, chosen in extensions.GROUPS.items():
        for kind, uri in chosen.items():
            where = "exact" if group == "exact" else f"other:{uri}"
            row(kind="save", group=group, extension=kind, uri=uri, saved_where=where)
    return lines


def test_the_comparison_accepts_a_capture_that_agrees_with_sheetpeek(tmp_path, monkeypatch):
    path = tmp_path / "extensions.tsv"
    monkeypatch.setattr(extensions, "OUTPUT", path)
    lines = _agreeing()
    checks = (
        test_the_controls_show_what_the_exact_uris_draw,
        test_every_exact_uri_is_read_and_written_back,
        test_sheetpeek_matches_each_uri_as_excel_does,
        test_the_screen_and_the_save_agree,
    )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    for check in checks:
        check()
    # and refuses one that does not: a lower-case sparkline drawn, then saved as read, a
    # lower-case validation saved as read, a triangle left on the exact exemption
    mutations = [
        (
            lambda line: line.startswith("screen\tsparkline_lower\tJ4"),
            ("\t0\tnone", "\t2400\t" + extensions.SERIES),
            test_sheetpeek_matches_each_uri_as_excel_does,
        ),
        (
            lambda line: line.startswith("save\tsparkline_lower\t"),
            ("\tother:", "\texact\tx"),
            test_the_screen_and_the_save_agree,
        ),
        (
            lambda line: line.startswith("save\tvalidation_lower\t"),
            ("\tother:", "\texact\tx"),
            test_sheetpeek_matches_each_uri_as_excel_does,
        ),
        (
            lambda line: line.startswith("screen\texact\tG9"),
            ("\t0\tnone", "\t30\t008000"),
            test_the_controls_show_what_the_exact_uris_draw,
        ),
    ]
    for chosen, (old, new), check in mutations:
        changed = [line.replace(old, new) if chosen(line) else line for line in lines]
        assert changed != lines, (old, new)
        path.write_text("\n".join(changed) + "\n", encoding="utf-8")
        with pytest.raises(AssertionError):
            check()
