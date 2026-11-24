"""Captures from Excel for Windows (``tests/captures/windows/``), against sheetpeek's layout with
``platform="windows"``.

Each file is the Mac suite's schema, read from a capture job's result by ``capture/job.py``; its
header names the job and holds the result's header. A comparison is skipped until its file exists.
"""

import tempfile
from pathlib import Path

import pytest

import sheetpeek
from tests.capture_report import REPORT

WINDOWS = Path(__file__).with_name("captures") / "windows"


def _rows(name):
    path = WINDOWS / name
    if not path.exists():
        pytest.skip(f"not captured yet: {path.name}")
    lines = path.read_text(encoding="utf-8").splitlines()
    head = [line for line in lines if line.startswith("# ")]
    assert "# platform: windows" in head
    return [line.split("\t") for line in lines if line and not line.startswith("# ")][1:]


@pytest.fixture(scope="module")
def cells(tmp_path_factory):
    pytest.importorskip("openpyxl")
    from capture.cell_text import build

    path = build(tmp_path_factory.mktemp("cell_text") / "cell_text.xlsx")
    return sheetpeek.open(path, platform="windows")


@pytest.mark.excel
def test_merge_width_matches_excel_for_windows(cells):
    rows: dict[tuple[str, bool], list[str]] = {}
    for width, merge, _value, _fmt, shown in _rows("merge_width.tsv"):
        rows.setdefault((width, merge == "A:C, B hidden"), []).append(shown)
    compared, deviations = 0, {}
    for (width, hidden), texts in rows.items():
        rendered = cells.sheet(("h" if hidden else "w") + width).layout("excel")
        for row, shown in enumerate(texts, 1):
            compared += 1
            if rendered.cell(row, 1).text != shown:
                deviations[(width, hidden, row)] = (rendered.cell(row, 1).text, shown)
    REPORT["merge width (Windows)"] = (
        f"merge width, Excel for Windows: {compared} cells compared, {len(deviations)} deviations"
    )
    assert compared == 150
    assert deviations == {}


@pytest.mark.excel
def test_error_width_matches_excel_for_windows(cells):
    from capture.cell_text import ERROR_WIDTHS, ERRORS

    rendered = cells.sheet("errors").layout("excel")
    records = _rows("error_width.tsv")
    deviations = {}
    for error, width, shown in records:
        cell = rendered.cell(ERRORS.index(error) + 1, ERROR_WIDTHS.index(float(width)) + 2)
        if cell.text != shown:
            deviations[(error, width)] = (cell.text, shown)
    REPORT["error width (Windows)"] = (
        f"error width, Excel for Windows: {len(records)} cells compared, {len(deviations)} "
        "deviations"
    )
    assert len(records) == 56
    assert deviations == {}


# --- borders and fills: the Mac's lines and tiles in device pixels


def _table(name):
    path = WINDOWS / name
    if not path.exists():
        pytest.skip(f"not captured yet: {path.name}")
    lines = path.read_text(encoding="utf-8").splitlines()
    head = [line for line in lines if line.startswith("# ")]
    rows = [line.split("\t") for line in lines if line and not line.startswith("# ")]
    return head, [dict(zip(rows[0], r, strict=True)) for r in rows[1:]]


def _px(length, zoom):
    """A length in sheet points at ``zoom`` as Windows device pixels (96 dpi)."""
    px = float(length) * 96 / 72 * int(zoom) / 100
    assert abs(px - round(px)) < 1e-6, (length, zoom)
    return round(px)


@pytest.mark.excel
def test_the_gridline_gray_is_the_one_recorded_for_windows():
    from capture.screen.window import GRIDLINE

    for name in ("borders.tsv", "fills.tsv"):
        head, _ = _table(name)
        (line,) = [x for x in head if x.startswith("# gridline: ")]
        gray = line.split()[2]
        assert gray == "".join(f"{c:02X}" for c in GRIDLINE["windows"]) == "E0E0E0"


@pytest.mark.excel
def test_windows_dashes_have_sheetpeeks_lengths_in_device_pixels():
    # As the Mac's test (tests/test_borders.py): every dashed side, dash and gap lengths exact in
    # the platform's unit (a device pixel here, where the Mac's is a screen point), the starting
    # point free; each device row against its own row of the tile (slantDashDot's two differ:
    # the row before the gridline is the tile's first, the row on it its third).
    from tests.test_borders import _fits_cycle, _tile_runs

    _, body = _table("borders.tsv")
    checked, wrong = 0, []
    for row in body:
        style, pattern = row["style"], row["pattern"]
        if style == "none" or pattern in ("solid", "empty", ""):
            continue
        tokens = pattern.removeprefix("irregular: ").split()
        measured = [
            (kind == "on", _px(length, row["zoom"]))
            for kind, length in zip(tokens[::2], tokens[1::2], strict=True)
        ]
        tile_row = 0 if float(row["offset"]) < 0 else 2
        checked += 1
        if not _fits_cycle(measured, _tile_runs(style, "windows", tile_row)):
            wrong.append((row["zoom"], style, row["edge"], row["offset"], pattern))
    REPORT["borders (Windows)"] = (
        f"borders, Excel for Windows: {checked} dashed sides compared, {len(wrong)} deviations"
    )
    assert checked == 104  # one device row a unit, where the Mac has two; slantDashDot's 16
    assert wrong == []


@pytest.mark.excel
def test_windows_lines_sit_where_the_macs_do_in_device_pixels():
    # Each style's device rows against its gridline, in Windows pixels, are the Mac's in screen
    # points: the Mac's two device rows a screen point read as one row each.
    mac = (WINDOWS.parent / "borders.tsv").read_text(encoding="utf-8").splitlines()
    rows = [x.split("\t") for x in mac if x and not x.startswith("#")]
    mac_body = [dict(zip(rows[0], r, strict=True)) for r in rows[1:]]
    _, body = _table("borders.tsv")

    def offsets(records, unit):
        found: dict[tuple[str, str, str], set[int]] = {}
        for r in records:
            n = float(r["offset"]) * int(r["zoom"]) / 100 / unit
            found.setdefault((r["zoom"], r["style"], r["edge"]), set()).add(int(n // 1))
        return found

    assert offsets(body, 0.75) == offsets(mac_body, 1.0)


@pytest.mark.excel
def test_windows_fill_tiles_are_sheetpeeks_in_device_pixels():
    from sheetpeek._fill_border_model import TILES

    def turns(tile):
        rows = tile.split("/")
        h, w = len(rows), len(rows[0])
        return {
            "/".join("".join(rows[(y + dy) % h][(x + dx) % w] for x in range(w)) for y in range(h))
            for dy in range(h)
            for dx in range(w)
        }

    _, body = _table("fills.tsv")
    compared, wrong = 0, []
    for row in body:
        if row["pattern"] not in TILES:
            continue
        compared += 1
        tile = TILES[row["pattern"]]
        # the measured tile repeats the pattern to a whole number of units on each axis
        h, w = len(row["tile_sheet"].split("/")), len(row["tile_sheet"].split("/")[0])
        ours = "/".join(
            "".join(tile[y % len(tile)][x % len(tile[0])] for x in range(w)) for y in range(h)
        )
        if row["tile_unit"] != "device_px" or row["tile_sheet"] not in turns(ours):
            wrong.append((row["zoom"], row["pattern"], row["cell"], row["tile_sheet"]))
    REPORT["fills (Windows)"] = (
        f"fills, Excel for Windows: {compared} pattern tiles compared, {len(wrong)} deviations"
    )
    assert compared == 74  # 17 patterns in two color sets and the 3 anchor cells, 2 zooms
    assert wrong == []


# --- text_steps: indents and shrink to fit in device pixels


@pytest.fixture(scope="module")
def steps_built(tmp_path_factory):
    pytest.importorskip("openpyxl")
    from capture import text_steps

    return text_steps.build(tmp_path_factory.mktemp("text_steps"))


def _steps_drawn(path, group):
    """sheetpeek's runs for each case with platform="windows": the first run's left and the last
    run's right from the cell's edges, in points, and the run's scale."""
    from capture import text_steps
    from openpyxl.utils import column_index_from_string
    from openpyxl.utils.cell import coordinate_from_string

    rendered = sheetpeek.read(path, platform="windows").layout("excel")
    columns = {c.index: c for c in rendered.columns}
    found = {}
    for case in text_steps.cases(group):
        letters, r = coordinate_from_string(case.cell)
        c = column_index_from_string(letters)
        row = next(x for x in rendered.rows if x.index == r)
        runs = [t for t in row.runs if t.source == (r, c)]
        left = columns[c].x_pt
        found[case.cell] = (
            min(t.x for t in runs) - left,
            max(t.x + t.width for t in runs) - left - columns[c].width_pt,
            runs[0].scale,
        )
    return found


@pytest.mark.excel
def test_windows_indents_move_text_as_sheetpeek_moves_it(steps_built):
    # As the Mac's test: the ink's edge at each level less its edge at level 0, against
    # sheetpeek's runs, within one device pixel (0.75 pt).
    from capture import text_steps

    _, body = _table("text_steps.tsv")
    captured = {(r["group"], r["cell"]): r for r in body}
    compared, wrong = 0, {}
    for group in text_steps.GROUPS:
        if not group.startswith("indent"):
            continue
        drawn = _steps_drawn(steps_built[group], group)
        found = text_steps.cases(group)
        for case in found:
            base = found[found.index(case) - int(case.name.split()[-1]) * 2]
            side, key = (0, "ink_left") if " left " in case.name else (1, "ink_right")
            moved = float(captured[(group, case.cell)][key]) - float(
                captured[(group, base.cell)][key]
            )
            expected = drawn[case.cell][side] - drawn[base.cell][side]
            compared += 1
            if abs(moved - expected) > 0.75:
                wrong[(group, case.name)] = (moved, expected)
    REPORT["indent (Windows)"] = (
        f"indent, Excel for Windows: {compared} cases compared, {len(wrong)} deviations"
    )
    assert compared == 108
    assert wrong == {}


@pytest.mark.excel
def test_windows_shrinks_each_length_to_the_pixel_level_sheetpeek_gives_it(steps_built):
    # The "H" advance Excel draws each length at, in device pixels: its ink width plus the
    # unshrunk text's side bearings (from the group's references), over its length; in Calibri
    # 11's group, the ink width over each length's unshrunk copy's times the unshrunk advance
    # (xlformat's, 9 px). A text clipped at the cell's right edge (ink to its last pixel) has no
    # measurable advance: sheetpeek must draw it at the smallest level, 2 px, as nothing fits.
    import xlformat
    from capture import text_steps

    _, body = _table("text_steps.tsv")
    captured = {(r["group"], r["cell"]): r for r in body}

    def ink(record):
        return float(record["cell_width"]) - float(record["ink_left"]) + float(record["ink_right"])

    compared, clipped, wrong = 0, 0, {}
    for group in ("shrink", "shrink_calibri16", "shrink_aptos_narrow11"):
        drawn = _steps_drawn(steps_built[group], group)
        if group == "shrink":
            font = xlformat.Font("Calibri", 11)
            unshrunk = round(xlformat.text_width("H", font=font, platform="windows") * 4 / 3)
        else:
            short, long = (ink(captured[(group, c)]) for c in text_steps.REFERENCES)
            unshrunk = round((long - short) / 10 * 4 / 3)
            bearings = (long - short) - short
        for case in text_steps.cases(group):
            if not case.name.startswith("shrunk"):
                continue
            record = captured[(group, case.cell)]
            ours = drawn[case.cell][2] * unshrunk
            compared += 1
            if float(record["ink_right"]) == 0:
                clipped += 1
                if round(ours, 6) != 2:
                    wrong[(group, case.name)] = ("clipped", ours)
                continue
            if group == "shrink":
                level = round(float(record["width_ratio"]) * unshrunk)
            else:
                n = int(case.name.split()[1])
                level = round((ink(record) + bearings) * 4 / 3 / n)
            if level != round(ours, 6):
                wrong[(group, case.name)] = (level, ours)
    REPORT["shrink (Windows)"] = (
        f"shrink, Excel for Windows: {compared} lengths compared ({clipped} clipped), "
        f"{len(wrong)} deviations"
    )
    assert compared == 27 + 48 + 48
    assert wrong == {}


# --- text_style: cell widths, fitted rows and line counts at zoom 100

KNOWN_ROWS: dict[str, tuple[int, ...]] = {
    # Known deviations, rows by group: rows fitted to their text, which sheetpeek fits with Excel
    # for Mac's rules; they wait for xlformat's Windows vertical metrics.
    "align": (14, 16, 18),
    "shrink": (4, 12),
    "fonts": (10, 12, 14, 16, 18, 20, 24),
    "wrap": tuple(range(4, 31, 2)),
    "lone": (4, 8),
}
KNOWN_LINES: dict[str, tuple[str, ...]] = {
    # Known deviations. Vertical justify and distributed text on Windows keeps on one line what
    # sheetpeek breaks (Excel fits "bbbb cccc", 59 px, on one line of a 64 px cell): their line
    # budget on Windows is not modeled.
    "align": ("G8", "G10", "G14", "G16", "G18"),
    "lone": ("C4", "E4", "G4", "I4", "I8"),
}
SHOWN_LINES: dict[str, tuple[str, ...]] = {
    # Not compared: a 15 pt row shows one line of several, and the capture counts the lines shown.
    "lone": ("C6", "E6", "G6", "I6"),
}


@pytest.mark.excel
def test_windows_text_style_widths_rows_and_lines_are_sheetpeeks():
    pytest.importorskip("openpyxl")
    import tempfile

    import openpyxl
    from capture import text_style
    from openpyxl.utils import column_index_from_string
    from openpyxl.utils.cell import coordinate_from_string

    _, body = _table("text_style.tsv")
    widths = rows = lines = 0
    wrong_widths, wrong_rows, wrong_lines = [], set(), set()
    with tempfile.TemporaryDirectory() as tmp:
        for group in text_style.GROUPS:
            wb = openpyxl.Workbook()
            text_style.fill(wb.active, group)
            path = Path(tmp) / f"{group}.xlsx"
            wb.save(path)
            rendered = sheetpeek.read(path, platform="windows").layout("excel")
            columns = {c.index: c for c in rendered.columns}
            by_row = {r.index: r for r in rendered.rows}
            for record in body:
                if record["group"] != group or record["zoom"] != "100":
                    continue
                letters, r = coordinate_from_string(record["cell"])
                c = column_index_from_string(letters)
                widths += 1
                if float(record["cell_width"]) != columns[c].width_pt:
                    wrong_widths.append((group, record["cell"]))
                rows += 1
                if float(record["row_height"]) != by_row[r].height_pt:
                    wrong_rows.add((group, r))
                if record["lines"] and record["cell"] not in SHOWN_LINES.get(group, ()):
                    lines += 1
                    ours = {round(t.y, 3) for t in by_row[r].runs if t.source == (r, c)}
                    if int(record["lines"]) != len(ours):
                        wrong_lines.add((group, record["cell"]))
    known_rows = {(g, r) for g, found in KNOWN_ROWS.items() for r in found}
    known_lines = {(g, c) for g, found in KNOWN_LINES.items() for c in found}
    REPORT["text_style (Windows)"] = (
        f"text_style, Excel for Windows: {widths} cell widths, {len(wrong_widths)} deviations; "
        f"{len(known_rows)} of the rows a known deviation (vertical metrics); {lines} line counts, "
        f"{len(known_lines)} a known deviation (justify and distributed line budget)"
    )
    assert (widths, rows, lines) == (141, 141, 78)
    assert wrong_widths == []
    assert wrong_rows == known_rows
    assert wrong_lines == known_lines


# --- shrink_fonts: the Windows shrink levels on fonts text_steps' other groups do not hold

KNOWN_FIT = {
    # Known deviations: Excel draws one level lower than sheetpeek's fit test gives, (n + 0.375)
    # times the level no more than the cell less a pixel; which length takes which level is not
    # modeled exactly on either platform (§4.13).
    "Calibri 20 shrunk 4": (13, 14),
    "Arial 11 shrunk 7": (7, 8),
}


@pytest.mark.excel
def test_windows_shrink_levels_hold_on_fonts_of_their_own():
    # Excel's level per case in device pixels: the ink width over the length less the side
    # bearings in advances (the font's two unshrunk references give its advance and bearings;
    # bearings scale with the drawn size). Every level is a whole pixel within 0.25, never 6.
    from capture import text_steps

    _, body = _table("shrink_fonts.tsv")
    records = {r["name"]: r for r in body}

    def ink(record):
        width = float(record["cell_width"]) - float(record["ink_left"])
        return (width + float(record["ink_right"])) * 4 / 3

    with tempfile.TemporaryDirectory() as tmp:
        path = text_steps.build(Path(tmp))["shrink_fonts"]
        drawn = _steps_drawn(path, "shrink_fonts")
    levels, wrong = {}, {}
    for face, size in text_steps.FONTS:
        font = f"{face} {size}"
        short = ink(records[f"{font} reference 10"])
        long = ink(records[f"{font} reference 20"])
        advance = (long - short) / 10
        bearings = 10 * advance - short
        for case in text_steps.cases("shrink_fonts"):
            if not case.name.startswith(f"{font} shrunk "):
                continue
            n = int(case.name.split()[-1])
            found = ink(records[case.name]) / (n - bearings / advance)
            assert abs(found - round(found)) <= 0.25, (case.name, found)
            level = round(found)
            levels[case.name] = level
            ours = round(drawn[case.cell][2] * advance, 6)
            if level != ours:
                wrong[case.name] = (level, ours)
    REPORT["shrink fonts (Windows)"] = (
        f"shrink, fresh fonts, Excel for Windows: {len(levels)} lengths, {len(wrong)} known "
        "deviations (the fit test)"
    )
    assert len(levels) == 55
    assert 6 not in levels.values()
    assert wrong == KNOWN_FIT


# --- line_budget: two words a few pixels short of the cell


@pytest.mark.excel
def test_windows_line_budget_against_sheetpeek():
    # Excel's lines from the capture's line bands, an "i"'s dots left out: on Windows they stand
    # 2 px above their stems, a band of 1 or 2 px of their own. Known deviations, one reason:
    # Excel for Windows' cell insets are 3 px left and 2 px right under cell fonts to 12 points
    # (2.25 and 1.5 pt, where sheetpeek takes xlformat's 3 and 2 pt on both platforms), so the
    # texts the cell less 6 and 5 px fit one line; larger cell fonts take larger insets (Calibri
    # 14 breaks them too), as yet unmodeled.
    pytest.importorskip("openpyxl")
    import openpyxl
    from capture import text_style
    from openpyxl.utils import column_index_from_string
    from openpyxl.utils.cell import coordinate_from_string

    _, body = _table("line_budget.tsv")
    records = {r["name"]: r for r in body}
    wb = openpyxl.Workbook()
    found = text_style.fill(wb.active, "line_budget")
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "line_budget.xlsx"
        wb.save(path)
        rendered = sheetpeek.read(path, platform="windows").layout("excel")
    by_row = {r.index: r for r in rendered.rows}
    wrong = set()
    for case in found:
        bands = records[case.name]["line_tops"].split()
        excel = sum(1 for band in bands if float(band.split("+")[1]) > 2.25)
        letters, r = coordinate_from_string(case.cell)
        c = column_index_from_string(letters)
        ours = len({round(t.y, 3) for t in by_row[r].runs if t.source == (r, c)})
        if excel != ours:
            wrong.add(case.name)
    known = {
        case.name
        for case in found
        if "Calibri 14" not in case.name and case.name.endswith(("cell-6px", "cell-5px"))
    }
    REPORT["line budget (Windows)"] = (
        f"line budget, Excel for Windows: {len(found)} cells, {len(wrong)} known deviations "
        "(insets 3 + 2 px under fonts to 12 points)"
    )
    assert len(found) == 136
    assert wrong == known


# --- font_cut: Excel for Windows' font table cut


@pytest.mark.excel
def test_windows_cuts_the_font_table_where_captured():
    # Excel for Windows draws a probe past its font table's cut in an 11-point face, black for a
    # list font and in its own color for a rich-text run, its row the larger of 15 pt and the
    # Normal font's. Unexplained, listed: the cut is at entry 1,000 under a Calibri 11 Normal font
    # (groups list and rich) and at 999 under a Times New Roman Normal font (the theme groups);
    # and the face is neither Calibri 11 nor Aptos Narrow 11 (its glyphs start at 2.25, 9.75 and
    # 13.5 pt where Calibri 11's start at 3, 9 and 13.5, and its "H" on the inset, where both
    # faces' start a pixel inside it).
    _, body = _table("font_cut.tsv")
    first_cut = {"list": 1000, "rich": 1000, "theme_10": 999, "theme_16": 999}
    assert len(body) == 50
    for r in body:
        cut = int(r["entry"]) >= first_cut[r["group"]]
        normal_row = 20.25 if r["group"] == "theme_16" else 15.0
        assert float(r["row_height"]) == (normal_row if cut else 30.0), r
        if cut:
            glyphs = "2.25 9 13.5" if r["group"] == "rich" else "2.25 9.75 13.5"
            assert r["glyphs"] == glyphs, r
            if r["group"] == "rich":
                assert int(r["dark"]) == 0, r  # its own red
            else:
                assert int(r["dark"]) > int(r["red"]), r  # black, ClearType fringes aside


@pytest.mark.excel
def test_windows_draws_a_cut_cell_in_aptos_narrow_11():
    # font_faces: a cut probe's glyphs and ink equal the Aptos Narrow 11 reference's exactly (its
    # standard font), and no other reference's; the cut is at entry 1,000 under a Calibri 11
    # Normal font and at 999 under Times New Roman 10, both under openpyxl's theme. Open: whether
    # the extra entry comes with a Normal font other than the theme's body font, or other than
    # Calibri (they part under an Aptos Narrow theme with an Aptos Narrow Normal font).
    _, body = _table("font_faces.tsv")
    assert len(body) == 22
    for group, first_cut in (("faces_calibri", 1000), ("faces_times", 999)):
        rows = [r for r in body if r["group"] == group]
        refs = {r["font_name"]: r for r in rows if int(r["entry"]) <= 5}
        for r in rows:
            if int(r["entry"]) <= 5:
                continue
            if int(r["entry"]) >= first_cut:
                matches = {
                    name
                    for name, ref in refs.items()
                    if (ref["glyphs"], ref["red"], ref["dark"])
                    == (r["glyphs"], r["red"], r["dark"])
                }
                assert matches == {"Aptos Narrow"}, r
            else:
                assert float(r["row_height"]) == 30.0, r


@pytest.mark.excel
def test_windows_cuts_at_1000_under_its_theme_body_font():
    # font_aptos: under a theme of Aptos Narrow body text and an Aptos Narrow 11 Normal font, the
    # cut is at entry 1,000 (not 999), the cut probes in Aptos Narrow 11: the extra entry comes
    # with a Normal font other than the theme's body font, not with one other than Calibri.
    _, body = _table("font_aptos.tsv")
    assert len(body) == 11
    refs = {r["font_name"]: r for r in body if int(r["entry"]) <= 5}
    for r in body:
        entry = int(r["entry"])
        if entry <= 5:
            continue
        if entry >= 1000:
            aptos = refs["Aptos Narrow"]
            assert (r["glyphs"], r["red"], r["dark"]) == (
                aptos["glyphs"],
                aptos["red"],
                aptos["dark"],
            )
            assert float(r["row_height"]) == 15.0, r
        else:
            assert float(r["row_height"]) == 30.0, r
