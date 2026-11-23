"""Capture which cells set a row's shared bottom baseline, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.row_bottoms run
    uv run --group capture python -m capture.row_bottoms run cells   # some groups only
    uv run --group capture python -m capture.row_bottoms read     # the saved captures again
    uv run --group capture python -m capture.row_bottoms build DIR    # the workbooks only

Bottom-aligned text in a row sits on one baseline, the largest bottom offset ``B`` over the
Normal font and the row's fonts (ARCHITECTURE.md §3.6; captured for cells with values in
``text_style_more``). This suite asks which other cells count, and where text that spans rows
or spreads down sits. Every probe is "HHH" (no descender: its ink ends on the baseline) in
Calibri 11 under a Calibri 11 Normal font (``B`` 3), bottom aligned in column C; the neighbor
is Calibri 24 (``B`` 6) in column E unless stated. Rows are fixed at 32 pt, so a counted
neighbor puts the probe's baseline 26 pt below the row's top and one not counted 29. Each group
is its own workbook and window at zoom 100, in the dark appearance as the other suites.

- ``cells``. Row 4: the probe alone (the control). Row 5: a neighbor with a value (the captured
  rule). Row 6: an empty cell styled Calibri 24. Row 7: a value in Calibri 24 in the hidden
  column G. Rows 8, 9 and 10: the neighbor turned 45 degrees, turned 90 degrees, and stacked.
  Rows 11 and 12: the neighbor merged down E11:E12, anchored in row 11, a probe in each row.
  Rows 13 and 14: the neighbor in superscript and in subscript (drawn at about 16 pt, ``B`` 4:
  a baseline of 28 if the drawn size counts). Row 15: the row's own style Calibri 24
  (``customFormat``), no neighbor cell. Row 16: a Calibri 11 neighbor with a 24 pt rich-text
  run. Rows 17 and 18: the probe merged down C17:C18 (64 pt), the neighbor in row 17 (the
  merge's first row); rows 19 and 20 the same with the neighbor in row 20 (its last row): which
  row's bottom a merge across rows takes (61 from the merge's top when no neighbor counts, 58
  when one does). Rows 21 and 22, fixed at 70 pt: three lines "HHH" vertically distributed and
  justified beside the neighbor, rows 23 and 24 the same alone (the last line 64 or 67 from the
  top).
- ``scripts_1`` to ``scripts_11``, rows fixed at 64 pt. The probe beside "HHH" in superscript or
  subscript in E, at every whole and half size from 6 to 40 (13 rows a group): the probe's
  baseline (the row's shared bottom) and the neighbor's (its raise or lowering). The candidates:
  the size taken whole (its ceiling) or as it is; raise ``floor(S/2)`` and lowering
  ``floor(S/5)``; a superscript's bottom its own ``B``, a subscript's ``l + B(d)``, ``d`` the
  drawn size ``round(2S/3)``; and sheetpeek's layout before the run.
- ``column_style``. Column E styled Calibri 24 (``<col style>``), no cell in it. Row 4, fixed at
  32 pt, and row 5, fitted (no height): the probe. Whether a column's style counts in a row's
  bottom, and in a fitted row's height (31 pt if it counts, 15 if not).

Geometry is read for one cell a column (row 1) and one a row (column A, which gives the row's
height). ``read`` writes ``tests/captures/row_bottoms.tsv``: one row per case, the ink of each
line (``line_tops``, from the cell's top, a merge's from the merge's top) and the row's height,
measured as ``text_style`` measures (its code, by import), lengths in points at zoom 100.
``PREDICTIONS`` holds, for each case, the last baseline each candidate rule gives, written
before the run.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Font
from openpyxl.utils.cell import coordinate_from_string

from capture import text_style
from capture.text_style import Case

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "row_bottoms"
OUTPUT = ROOT / "tests" / "captures" / "row_bottoms.tsv"
SCRIPT_SIZES = tuple(sorted([float(n) for n in range(6, 41)] + [n + 0.5 for n in range(6, 40)]))
"""``scripts_<k>``: every whole and half size from 6 to 40."""
SCRIPT_ROWS = tuple((kind, size) for size in SCRIPT_SIZES for kind in ("superscript", "subscript"))
SCRIPT_GROUP_ROWS = 13
"""Rows of 64 pt a window holds at zoom 100."""
SCRIPT_ROW = 64
SCRIPT_GROUPS = tuple(f"scripts_{k + 1}" for k in range(-(-len(SCRIPT_ROWS) // SCRIPT_GROUP_ROWS)))
GROUPS = ("cells", "column_style", *SCRIPT_GROUPS)
MARKERS = {
    "cells": ("B2", "I3"),
    "column_style": ("B2", "G3"),
    **dict.fromkeys(SCRIPT_GROUPS, ("B2", "G3")),
}
"""The second marker sits just right of every case, in row 3."""
PROBE = "HHH"
ROW = 32
"""Every fixed row's height, but the spread rows'."""
SPREAD_ROW = 70
NEIGHBOR = Font(name="Calibri", size=24)
COLUMNS = (
    "group", "name", "zoom", "cell", "row_height", "lines", "line_tops", "baselines",
)  # fmt: skip

# The last baseline, from the cell's (or merge's) top, under each candidate rule. A fixed row of
# 32 pt: 29 with the probe's own bottom (3), 26 with the neighbor's (6); 28 with a 16 pt drawn
# size's (4). A 64 pt merge: 61 or 58. A 70 pt row: 67 or 64, or 61 when the wrapped lines also
# sit the difference of the line insets higher (Calibri 24's 4 against the probe's 1), as
# bottom-aligned wrapped text does (§3.6).
PREDICTIONS: dict[str, dict[str, float]] = {
    "probe alone": {"own": 29},
    "neighbor with a value": {"counted": 26, "not counted": 29},
    "empty styled neighbor": {"counted": 26, "not counted": 29},
    "neighbor in a hidden column": {"counted": 26, "not counted": 29},
    "neighbor at 45 degrees": {"counted": 26, "not counted": 29},
    "neighbor at 90 degrees": {"counted": 26, "not counted": 29},
    "stacked neighbor": {"counted": 26, "not counted": 29},
    "merge anchored above, first row": {"counted": 26, "not counted": 29},
    "merge anchored above, second row": {"counted": 26, "not counted": 29},
    "superscript neighbor": {"counted": 26, "drawn size": 28, "not counted": 29},
    # "lowered": l + B(d), the rule the scripts groups settled (added after the cells capture,
    # where no candidate fitted: 4 + B(16) = 8 above the bottom)
    "subscript neighbor": {"counted": 26, "drawn size": 28, "not counted": 29, "lowered": 24},
    "row style": {"counted": 26, "not counted": 29},
    "neighbor with a large run": {"counted": 26, "not counted": 29},
    "merged probe, neighbor in its first row": {
        "last row's": 61, "first row's": 58, "largest": 58,
    },
    "merged probe, neighbor in its last row": {
        "last row's": 58, "first row's": 61, "largest": 58,
    },
    "distributed beside the neighbor": {"counted": 64, "counted, inset": 61, "not counted": 67},
    "justify beside the neighbor": {"counted": 64, "counted, inset": 61, "not counted": 67},
    "distributed alone": {"own": 67},
    "justify alone": {"own": 67},
    "column style, fixed row": {"counted": 26, "not counted": 29},
    "column style, fitted row": {"counted": 25, "not counted": 12},
    "superscript 6 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 6 glyph": {"whole": 58, "own": 58, "before": 58},
    "subscript 6 probe": {"whole": 60, "own": 60, "before": 61},
    "subscript 6 glyph": {"whole": 61, "own": 61, "before": 62},
    "superscript 6.5 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 6.5 glyph": {"whole": 58, "own": 58, "before": 58},
    "subscript 6.5 probe": {"whole": 60, "own": 60, "before": 61},
    "subscript 6.5 glyph": {"whole": 61, "own": 61, "before": 62},
    "superscript 7 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 7 glyph": {"whole": 58, "own": 58, "before": 58},
    "subscript 7 probe": {"whole": 60, "own": 60, "before": 61},
    "subscript 7 glyph": {"whole": 61, "own": 61, "before": 62},
    "superscript 7.5 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 7.5 glyph": {"whole": 57, "own": 58, "before": 57},
    "subscript 7.5 probe": {"whole": 60, "own": 60, "before": 61},
    "subscript 7.5 glyph": {"whole": 61, "own": 61, "before": 62},
    "superscript 8 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 8 glyph": {"whole": 57, "own": 57, "before": 57},
    "subscript 8 probe": {"whole": 60, "own": 60, "before": 61},
    "subscript 8 glyph": {"whole": 61, "own": 61, "before": 62},
    "superscript 8.5 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 8.5 glyph": {"whole": 57, "own": 57, "before": 57},
    "subscript 8.5 probe": {"whole": 61, "own": 61, "before": 61},
    "subscript 8.5 glyph": {"whole": 62, "own": 62, "before": 63},
    "superscript 9 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 9 glyph": {"whole": 57, "own": 57, "before": 57},
    "subscript 9 probe": {"whole": 61, "own": 61, "before": 61},
    "subscript 9 glyph": {"whole": 62, "own": 62, "before": 63},
    "superscript 9.5 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 9.5 glyph": {"whole": 56, "own": 57, "before": 56},
    "subscript 9.5 probe": {"whole": 60, "own": 61, "before": 61},
    "subscript 9.5 glyph": {"whole": 62, "own": 62, "before": 63},
    "superscript 10 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 10 glyph": {"whole": 56, "own": 56, "before": 56},
    "subscript 10 probe": {"whole": 60, "own": 60, "before": 61},
    "subscript 10 glyph": {"whole": 62, "own": 62, "before": 63},
    "superscript 10.5 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 10.5 glyph": {"whole": 56, "own": 56, "before": 56},
    "subscript 10.5 probe": {"whole": 60, "own": 60, "before": 61},
    "subscript 10.5 glyph": {"whole": 62, "own": 62, "before": 63},
    "superscript 11 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 11 glyph": {"whole": 56, "own": 56, "before": 56},
    "subscript 11 probe": {"whole": 60, "own": 60, "before": 61},
    "subscript 11 glyph": {"whole": 62, "own": 62, "before": 63},
    "superscript 11.5 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 11.5 glyph": {"whole": 55, "own": 56, "before": 55},
    "subscript 11.5 probe": {"whole": 60, "own": 60, "before": 61},
    "subscript 11.5 glyph": {"whole": 62, "own": 62, "before": 63},
    "superscript 12 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 12 glyph": {"whole": 55, "own": 55, "before": 55},
    "subscript 12 probe": {"whole": 60, "own": 60, "before": 61},
    "subscript 12 glyph": {"whole": 62, "own": 62, "before": 63},
    "superscript 12.5 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 12.5 glyph": {"whole": 55, "own": 55, "before": 55},
    "subscript 12.5 probe": {"whole": 60, "own": 60, "before": 61},
    "subscript 12.5 glyph": {"whole": 62, "own": 62, "before": 63},
    "superscript 13 probe": {"whole": 61, "own": 61, "before": 61},
    "superscript 13 glyph": {"whole": 55, "own": 55, "before": 55},
    "subscript 13 probe": {"whole": 60, "own": 60, "before": 61},
    "subscript 13 glyph": {"whole": 62, "own": 62, "before": 63},
    "superscript 13.5 probe": {"whole": 60, "own": 60, "before": 60},
    "superscript 13.5 glyph": {"whole": 53, "own": 54, "before": 54},
    "subscript 13.5 probe": {"whole": 60, "own": 60, "before": 60},
    "subscript 13.5 glyph": {"whole": 62, "own": 62, "before": 62},
    "superscript 14 probe": {"whole": 60, "own": 60, "before": 60},
    "superscript 14 glyph": {"whole": 53, "own": 53, "before": 53},
    "subscript 14 probe": {"whole": 60, "own": 60, "before": 60},
    "subscript 14 glyph": {"whole": 62, "own": 62, "before": 63},
    "superscript 14.5 probe": {"whole": 60, "own": 60, "before": 60},
    "superscript 14.5 glyph": {"whole": 53, "own": 53, "before": 53},
    "subscript 14.5 probe": {"whole": 58, "own": 59, "before": 60},
    "subscript 14.5 glyph": {"whole": 61, "own": 61, "before": 63},
    "superscript 15 probe": {"whole": 60, "own": 60, "before": 60},
    "superscript 15 glyph": {"whole": 53, "own": 53, "before": 53},
    "subscript 15 probe": {"whole": 58, "own": 58, "before": 60},
    "subscript 15 glyph": {"whole": 61, "own": 61, "before": 63},
    "superscript 15.5 probe": {"whole": 60, "own": 60, "before": 60},
    "superscript 15.5 glyph": {"whole": 52, "own": 53, "before": 53},
    "subscript 15.5 probe": {"whole": 58, "own": 58, "before": 60},
    "subscript 15.5 glyph": {"whole": 61, "own": 61, "before": 63},
    "superscript 16 probe": {"whole": 60, "own": 60, "before": 60},
    "superscript 16 glyph": {"whole": 52, "own": 52, "before": 52},
    "subscript 16 probe": {"whole": 58, "own": 58, "before": 60},
    "subscript 16 glyph": {"whole": 61, "own": 61, "before": 63},
    "superscript 16.5 probe": {"whole": 59, "own": 59, "before": 59},
    "superscript 16.5 glyph": {"whole": 51, "own": 51, "before": 51},
    "subscript 16.5 probe": {"whole": 58, "own": 58, "before": 59},
    "subscript 16.5 glyph": {"whole": 61, "own": 61, "before": 62},
    "superscript 17 probe": {"whole": 59, "own": 59, "before": 59},
    "superscript 17 glyph": {"whole": 51, "own": 51, "before": 51},
    "subscript 17 probe": {"whole": 58, "own": 58, "before": 59},
    "subscript 17 glyph": {"whole": 61, "own": 61, "before": 62},
    "superscript 17.5 probe": {"whole": 59, "own": 59, "before": 59},
    "superscript 17.5 glyph": {"whole": 50, "own": 51, "before": 51},
    "subscript 17.5 probe": {"whole": 58, "own": 58, "before": 59},
    "subscript 17.5 glyph": {"whole": 61, "own": 61, "before": 62},
    "superscript 18 probe": {"whole": 59, "own": 59, "before": 59},
    "superscript 18 glyph": {"whole": 50, "own": 50, "before": 50},
    "subscript 18 probe": {"whole": 58, "own": 58, "before": 59},
    "subscript 18 glyph": {"whole": 61, "own": 61, "before": 62},
    "superscript 18.5 probe": {"whole": 59, "own": 59, "before": 59},
    "superscript 18.5 glyph": {"whole": 50, "own": 50, "before": 50},
    "subscript 18.5 probe": {"whole": 58, "own": 58, "before": 59},
    "subscript 18.5 glyph": {"whole": 61, "own": 61, "before": 62},
    "superscript 19 probe": {"whole": 59, "own": 59, "before": 59},
    "superscript 19 glyph": {"whole": 50, "own": 50, "before": 50},
    "subscript 19 probe": {"whole": 58, "own": 58, "before": 59},
    "subscript 19 glyph": {"whole": 61, "own": 61, "before": 62},
    "superscript 19.5 probe": {"whole": 59, "own": 59, "before": 59},
    "superscript 19.5 glyph": {"whole": 49, "own": 50, "before": 50},
    "subscript 19.5 probe": {"whole": 57, "own": 58, "before": 59},
    "subscript 19.5 glyph": {"whole": 61, "own": 61, "before": 63},
    "superscript 20 probe": {"whole": 59, "own": 59, "before": 59},
    "superscript 20 glyph": {"whole": 49, "own": 49, "before": 49},
    "subscript 20 probe": {"whole": 57, "own": 57, "before": 59},
    "subscript 20 glyph": {"whole": 61, "own": 61, "before": 63},
    "superscript 20.5 probe": {"whole": 58, "own": 58, "before": 58},
    "superscript 20.5 glyph": {"whole": 48, "own": 48, "before": 48},
    "subscript 20.5 probe": {"whole": 56, "own": 56, "before": 58},
    "subscript 20.5 glyph": {"whole": 60, "own": 60, "before": 62},
    "superscript 21 probe": {"whole": 58, "own": 58, "before": 58},
    "superscript 21 glyph": {"whole": 48, "own": 48, "before": 48},
    "subscript 21 probe": {"whole": 56, "own": 56, "before": 58},
    "subscript 21 glyph": {"whole": 60, "own": 60, "before": 62},
    "superscript 21.5 probe": {"whole": 58, "own": 58, "before": 58},
    "superscript 21.5 glyph": {"whole": 47, "own": 48, "before": 48},
    "subscript 21.5 probe": {"whole": 56, "own": 56, "before": 58},
    "subscript 21.5 glyph": {"whole": 60, "own": 60, "before": 62},
    "superscript 22 probe": {"whole": 58, "own": 58, "before": 58},
    "superscript 22 glyph": {"whole": 47, "own": 47, "before": 47},
    "subscript 22 probe": {"whole": 56, "own": 56, "before": 58},
    "subscript 22 glyph": {"whole": 60, "own": 60, "before": 62},
    "superscript 22.5 probe": {"whole": 58, "own": 58, "before": 58},
    "superscript 22.5 glyph": {"whole": 47, "own": 47, "before": 47},
    "subscript 22.5 probe": {"whole": 56, "own": 56, "before": 58},
    "subscript 22.5 glyph": {"whole": 60, "own": 60, "before": 62},
    "superscript 23 probe": {"whole": 58, "own": 58, "before": 58},
    "superscript 23 glyph": {"whole": 47, "own": 47, "before": 47},
    "subscript 23 probe": {"whole": 56, "own": 56, "before": 58},
    "subscript 23 glyph": {"whole": 60, "own": 60, "before": 62},
    "superscript 23.5 probe": {"whole": 58, "own": 58, "before": 58},
    "superscript 23.5 glyph": {"whole": 46, "own": 47, "before": 47},
    "subscript 23.5 probe": {"whole": 56, "own": 56, "before": 58},
    "subscript 23.5 glyph": {"whole": 60, "own": 60, "before": 62},
    "superscript 24 probe": {"whole": 58, "own": 58, "before": 58},
    "superscript 24 glyph": {"whole": 46, "own": 46, "before": 46},
    "subscript 24 probe": {"whole": 56, "own": 56, "before": 58},
    "subscript 24 glyph": {"whole": 60, "own": 60, "before": 62},
    "superscript 24.5 probe": {"whole": 57, "own": 57, "before": 57},
    "superscript 24.5 glyph": {"whole": 45, "own": 45, "before": 45},
    "subscript 24.5 probe": {"whole": 54, "own": 56, "before": 57},
    "subscript 24.5 glyph": {"whole": 59, "own": 60, "before": 61},
    "superscript 25 probe": {"whole": 57, "own": 57, "before": 57},
    "superscript 25 glyph": {"whole": 45, "own": 45, "before": 45},
    "subscript 25 probe": {"whole": 54, "own": 54, "before": 57},
    "subscript 25 glyph": {"whole": 59, "own": 59, "before": 61},
    "superscript 25.5 probe": {"whole": 57, "own": 57, "before": 57},
    "superscript 25.5 glyph": {"whole": 44, "own": 45, "before": 45},
    "subscript 25.5 probe": {"whole": 54, "own": 54, "before": 57},
    "subscript 25.5 glyph": {"whole": 59, "own": 59, "before": 62},
    "superscript 26 probe": {"whole": 57, "own": 57, "before": 57},
    "superscript 26 glyph": {"whole": 44, "own": 44, "before": 45},
    "subscript 26 probe": {"whole": 54, "own": 54, "before": 57},
    "subscript 26 glyph": {"whole": 59, "own": 59, "before": 62},
    "superscript 26.5 probe": {"whole": 57, "own": 57, "before": 57},
    "superscript 26.5 glyph": {"whole": 44, "own": 44, "before": 44},
    "subscript 26.5 probe": {"whole": 54, "own": 54, "before": 57},
    "subscript 26.5 glyph": {"whole": 59, "own": 59, "before": 62},
    "superscript 27 probe": {"whole": 57, "own": 57, "before": 57},
    "superscript 27 glyph": {"whole": 44, "own": 44, "before": 44},
    "subscript 27 probe": {"whole": 54, "own": 54, "before": 57},
    "subscript 27 glyph": {"whole": 59, "own": 59, "before": 62},
    "superscript 27.5 probe": {"whole": 56, "own": 56, "before": 56},
    "superscript 27.5 glyph": {"whole": 42, "own": 43, "before": 43},
    "subscript 27.5 probe": {"whole": 54, "own": 54, "before": 56},
    "subscript 27.5 glyph": {"whole": 59, "own": 59, "before": 61},
    "superscript 28 probe": {"whole": 56, "own": 56, "before": 56},
    "superscript 28 glyph": {"whole": 42, "own": 42, "before": 43},
    "subscript 28 probe": {"whole": 54, "own": 54, "before": 56},
    "subscript 28 glyph": {"whole": 59, "own": 59, "before": 61},
    "superscript 28.5 probe": {"whole": 56, "own": 56, "before": 56},
    "superscript 28.5 glyph": {"whole": 42, "own": 42, "before": 42},
    "subscript 28.5 probe": {"whole": 54, "own": 54, "before": 56},
    "subscript 28.5 glyph": {"whole": 59, "own": 59, "before": 61},
    "superscript 29 probe": {"whole": 56, "own": 56, "before": 56},
    "superscript 29 glyph": {"whole": 42, "own": 42, "before": 42},
    "subscript 29 probe": {"whole": 54, "own": 54, "before": 56},
    "subscript 29 glyph": {"whole": 59, "own": 59, "before": 61},
    "superscript 29.5 probe": {"whole": 56, "own": 56, "before": 56},
    "superscript 29.5 glyph": {"whole": 41, "own": 42, "before": 42},
    "subscript 29.5 probe": {"whole": 53, "own": 54, "before": 56},
    "subscript 29.5 glyph": {"whole": 59, "own": 59, "before": 61},
    "superscript 30 probe": {"whole": 56, "own": 56, "before": 56},
    "superscript 30 glyph": {"whole": 41, "own": 41, "before": 42},
    "subscript 30 probe": {"whole": 53, "own": 53, "before": 56},
    "subscript 30 glyph": {"whole": 59, "own": 59, "before": 61},
    "superscript 30.5 probe": {"whole": 56, "own": 56, "before": 56},
    "superscript 30.5 glyph": {"whole": 41, "own": 41, "before": 41},
    "subscript 30.5 probe": {"whole": 52, "own": 53, "before": 56},
    "subscript 30.5 glyph": {"whole": 58, "own": 59, "before": 61},
    "superscript 31 probe": {"whole": 56, "own": 56, "before": 56},
    "superscript 31 glyph": {"whole": 41, "own": 41, "before": 41},
    "subscript 31 probe": {"whole": 52, "own": 52, "before": 56},
    "subscript 31 glyph": {"whole": 58, "own": 58, "before": 62},
    "superscript 31.5 probe": {"whole": 55, "own": 55, "before": 55},
    "superscript 31.5 glyph": {"whole": 39, "own": 40, "before": 40},
    "subscript 31.5 probe": {"whole": 52, "own": 52, "before": 55},
    "subscript 31.5 glyph": {"whole": 58, "own": 58, "before": 61},
    "superscript 32 probe": {"whole": 55, "own": 55, "before": 55},
    "superscript 32 glyph": {"whole": 39, "own": 39, "before": 40},
    "subscript 32 probe": {"whole": 52, "own": 52, "before": 55},
    "subscript 32 glyph": {"whole": 58, "own": 58, "before": 61},
    "superscript 32.5 probe": {"whole": 55, "own": 55, "before": 55},
    "superscript 32.5 glyph": {"whole": 39, "own": 39, "before": 39},
    "subscript 32.5 probe": {"whole": 52, "own": 52, "before": 55},
    "subscript 32.5 glyph": {"whole": 58, "own": 58, "before": 61},
    "superscript 33 probe": {"whole": 55, "own": 55, "before": 55},
    "superscript 33 glyph": {"whole": 39, "own": 39, "before": 39},
    "subscript 33 probe": {"whole": 52, "own": 52, "before": 55},
    "subscript 33 glyph": {"whole": 58, "own": 58, "before": 61},
    "superscript 33.5 probe": {"whole": 55, "own": 55, "before": 55},
    "superscript 33.5 glyph": {"whole": 38, "own": 39, "before": 39},
    "subscript 33.5 probe": {"whole": 52, "own": 52, "before": 55},
    "subscript 33.5 glyph": {"whole": 58, "own": 58, "before": 61},
    "superscript 34 probe": {"whole": 55, "own": 55, "before": 55},
    "superscript 34 glyph": {"whole": 38, "own": 38, "before": 39},
    "subscript 34 probe": {"whole": 52, "own": 52, "before": 55},
    "subscript 34 glyph": {"whole": 58, "own": 58, "before": 61},
    "superscript 34.5 probe": {"whole": 55, "own": 55, "before": 55},
    "superscript 34.5 glyph": {"whole": 38, "own": 38, "before": 38},
    "subscript 34.5 probe": {"whole": 51, "own": 52, "before": 55},
    "subscript 34.5 glyph": {"whole": 58, "own": 58, "before": 61},
    "superscript 35 probe": {"whole": 55, "own": 55, "before": 55},
    "superscript 35 glyph": {"whole": 38, "own": 38, "before": 38},
    "subscript 35 probe": {"whole": 51, "own": 51, "before": 55},
    "subscript 35 glyph": {"whole": 58, "own": 58, "before": 61},
    "superscript 35.5 probe": {"whole": 54, "own": 54, "before": 54},
    "superscript 35.5 glyph": {"whole": 36, "own": 37, "before": 37},
    "subscript 35.5 probe": {"whole": 51, "own": 51, "before": 54},
    "subscript 35.5 glyph": {"whole": 58, "own": 58, "before": 60},
    "superscript 36 probe": {"whole": 54, "own": 54, "before": 54},
    "superscript 36 glyph": {"whole": 36, "own": 36, "before": 37},
    "subscript 36 probe": {"whole": 51, "own": 51, "before": 54},
    "subscript 36 glyph": {"whole": 58, "own": 58, "before": 60},
    "superscript 36.5 probe": {"whole": 54, "own": 54, "before": 54},
    "superscript 36.5 glyph": {"whole": 36, "own": 36, "before": 36},
    "subscript 36.5 probe": {"whole": 50, "own": 51, "before": 54},
    "subscript 36.5 glyph": {"whole": 57, "own": 58, "before": 61},
    "superscript 37 probe": {"whole": 54, "own": 54, "before": 54},
    "superscript 37 glyph": {"whole": 36, "own": 36, "before": 36},
    "subscript 37 probe": {"whole": 50, "own": 50, "before": 54},
    "subscript 37 glyph": {"whole": 57, "own": 57, "before": 61},
    "superscript 37.5 probe": {"whole": 54, "own": 54, "before": 54},
    "superscript 37.5 glyph": {"whole": 35, "own": 36, "before": 36},
    "subscript 37.5 probe": {"whole": 50, "own": 50, "before": 54},
    "subscript 37.5 glyph": {"whole": 57, "own": 57, "before": 61},
    "superscript 38 probe": {"whole": 54, "own": 54, "before": 54},
    "superscript 38 glyph": {"whole": 35, "own": 35, "before": 36},
    "subscript 38 probe": {"whole": 50, "own": 50, "before": 54},
    "subscript 38 glyph": {"whole": 57, "own": 57, "before": 61},
    "superscript 38.5 probe": {"whole": 54, "own": 54, "before": 54},
    "superscript 38.5 glyph": {"whole": 35, "own": 35, "before": 36},
    "subscript 38.5 probe": {"whole": 50, "own": 50, "before": 54},
    "subscript 38.5 glyph": {"whole": 57, "own": 57, "before": 61},
    "superscript 39 probe": {"whole": 54, "own": 54, "before": 54},
    "superscript 39 glyph": {"whole": 35, "own": 35, "before": 35},
    "subscript 39 probe": {"whole": 50, "own": 50, "before": 54},
    "subscript 39 glyph": {"whole": 57, "own": 57, "before": 61},
    "superscript 39.5 probe": {"whole": 53, "own": 53, "before": 53},
    "superscript 39.5 glyph": {"whole": 33, "own": 34, "before": 34},
    "subscript 39.5 probe": {"whole": 49, "own": 50, "before": 53},
    "subscript 39.5 glyph": {"whole": 57, "own": 57, "before": 60},
    "superscript 40 probe": {"whole": 53, "own": 53, "before": 53},
    "superscript 40 glyph": {"whole": 33, "own": 33, "before": 34},
    "subscript 40 probe": {"whole": 49, "own": 49, "before": 53},
    "subscript 40 glyph": {"whole": 57, "own": 57, "before": 60},
}  # fmt: skip
"""Each case's last baseline under each candidate (for ``scripts_<k>``: ``whole``, the size
taken whole; ``own``, as it is; ``before``, sheetpeek's layout before the run). For the fitted
row of ``column_style``, a counted column style makes the row Calibri 24's 31 pt and puts the
probe on 31 - 6."""
FITTED_HEIGHTS = {"column style, fitted row": {"counted": 31, "not counted": 15}}
"""The fitted row's height under each candidate."""
SHEETPEEK: dict[str, str] = {
    "probe alone": "own",
    "neighbor with a value": "counted",
    "empty styled neighbor": "counted",
    "neighbor in a hidden column": "counted",
    "neighbor at 45 degrees": "not counted",
    "neighbor at 90 degrees": "not counted",
    "stacked neighbor": "not counted",
    "merge anchored above, first row": "not counted",
    "merge anchored above, second row": "not counted",
    "superscript neighbor": "counted",
    "subscript neighbor": "lowered",
    "row style": "counted",
    "neighbor with a large run": "counted",
    "merged probe, neighbor in its first row": "first row's",
    "merged probe, neighbor in its last row": "first row's",
    "distributed beside the neighbor": "not counted",
    "justify beside the neighbor": "not counted",
    "distributed alone": "own",
    "justify alone": "own",
    "column style, fixed row": "counted",
    "column style, fitted row": "counted",
    "superscript 6 probe": "whole",
    "superscript 6 glyph": "whole",
    "subscript 6 probe": "whole",
    "subscript 6 glyph": "whole",
    "superscript 6.5 probe": "whole",
    "superscript 6.5 glyph": "whole",
    "subscript 6.5 probe": "whole",
    "subscript 6.5 glyph": "whole",
    "superscript 7 probe": "whole",
    "superscript 7 glyph": "whole",
    "subscript 7 probe": "whole",
    "subscript 7 glyph": "whole",
    "superscript 7.5 probe": "whole",
    "superscript 7.5 glyph": "whole",
    "subscript 7.5 probe": "whole",
    "subscript 7.5 glyph": "whole",
    "superscript 8 probe": "whole",
    "superscript 8 glyph": "whole",
    "subscript 8 probe": "whole",
    "subscript 8 glyph": "whole",
    "superscript 8.5 probe": "whole",
    "superscript 8.5 glyph": "whole",
    "subscript 8.5 probe": "whole",
    "subscript 8.5 glyph": "whole",
    "superscript 9 probe": "whole",
    "superscript 9 glyph": "whole",
    "subscript 9 probe": "whole",
    "subscript 9 glyph": "whole",
    "superscript 9.5 probe": "whole",
    "superscript 9.5 glyph": "whole",
    "subscript 9.5 probe": "whole",
    "subscript 9.5 glyph": "whole",
    "superscript 10 probe": "whole",
    "superscript 10 glyph": "whole",
    "subscript 10 probe": "whole",
    "subscript 10 glyph": "whole",
    "superscript 10.5 probe": "whole",
    "superscript 10.5 glyph": "whole",
    "subscript 10.5 probe": "whole",
    "subscript 10.5 glyph": "whole",
    "superscript 11 probe": "whole",
    "superscript 11 glyph": "whole",
    "subscript 11 probe": "whole",
    "subscript 11 glyph": "whole",
    "superscript 11.5 probe": "whole",
    "superscript 11.5 glyph": "whole",
    "subscript 11.5 probe": "whole",
    "subscript 11.5 glyph": "whole",
    "superscript 12 probe": "whole",
    "superscript 12 glyph": "whole",
    "subscript 12 probe": "whole",
    "subscript 12 glyph": "whole",
    "superscript 12.5 probe": "whole",
    "superscript 12.5 glyph": "whole",
    "subscript 12.5 probe": "whole",
    "subscript 12.5 glyph": "whole",
    "superscript 13 probe": "whole",
    "superscript 13 glyph": "whole",
    "subscript 13 probe": "whole",
    "subscript 13 glyph": "whole",
    "superscript 13.5 probe": "whole",
    "superscript 13.5 glyph": "whole",
    "subscript 13.5 probe": "whole",
    "subscript 13.5 glyph": "whole",
    "superscript 14 probe": "whole",
    "superscript 14 glyph": "whole",
    "subscript 14 probe": "whole",
    "subscript 14 glyph": "whole",
    "superscript 14.5 probe": "whole",
    "superscript 14.5 glyph": "whole",
    "subscript 14.5 probe": "whole",
    "subscript 14.5 glyph": "whole",
    "superscript 15 probe": "whole",
    "superscript 15 glyph": "whole",
    "subscript 15 probe": "whole",
    "subscript 15 glyph": "whole",
    "superscript 15.5 probe": "whole",
    "superscript 15.5 glyph": "whole",
    "subscript 15.5 probe": "whole",
    "subscript 15.5 glyph": "whole",
    "superscript 16 probe": "whole",
    "superscript 16 glyph": "whole",
    "subscript 16 probe": "whole",
    "subscript 16 glyph": "whole",
    "superscript 16.5 probe": "whole",
    "superscript 16.5 glyph": "whole",
    "subscript 16.5 probe": "whole",
    "subscript 16.5 glyph": "whole",
    "superscript 17 probe": "whole",
    "superscript 17 glyph": "whole",
    "subscript 17 probe": "whole",
    "subscript 17 glyph": "whole",
    "superscript 17.5 probe": "whole",
    "superscript 17.5 glyph": "whole",
    "subscript 17.5 probe": "whole",
    "subscript 17.5 glyph": "whole",
    "superscript 18 probe": "whole",
    "superscript 18 glyph": "whole",
    "subscript 18 probe": "whole",
    "subscript 18 glyph": "whole",
    "superscript 18.5 probe": "whole",
    "superscript 18.5 glyph": "whole",
    "subscript 18.5 probe": "whole",
    "subscript 18.5 glyph": "whole",
    "superscript 19 probe": "whole",
    "superscript 19 glyph": "whole",
    "subscript 19 probe": "whole",
    "subscript 19 glyph": "whole",
    "superscript 19.5 probe": "whole",
    "superscript 19.5 glyph": "whole",
    "subscript 19.5 probe": "whole",
    "subscript 19.5 glyph": "whole",
    "superscript 20 probe": "whole",
    "superscript 20 glyph": "whole",
    "subscript 20 probe": "whole",
    "subscript 20 glyph": "whole",
    "superscript 20.5 probe": "whole",
    "superscript 20.5 glyph": "whole",
    "subscript 20.5 probe": "whole",
    "subscript 20.5 glyph": "whole",
    "superscript 21 probe": "whole",
    "superscript 21 glyph": "whole",
    "subscript 21 probe": "whole",
    "subscript 21 glyph": "whole",
    "superscript 21.5 probe": "whole",
    "superscript 21.5 glyph": "whole",
    "subscript 21.5 probe": "whole",
    "subscript 21.5 glyph": "whole",
    "superscript 22 probe": "whole",
    "superscript 22 glyph": "whole",
    "subscript 22 probe": "whole",
    "subscript 22 glyph": "whole",
    "superscript 22.5 probe": "whole",
    "superscript 22.5 glyph": "whole",
    "subscript 22.5 probe": "whole",
    "subscript 22.5 glyph": "whole",
    "superscript 23 probe": "whole",
    "superscript 23 glyph": "whole",
    "subscript 23 probe": "whole",
    "subscript 23 glyph": "whole",
    "superscript 23.5 probe": "whole",
    "superscript 23.5 glyph": "whole",
    "subscript 23.5 probe": "whole",
    "subscript 23.5 glyph": "whole",
    "superscript 24 probe": "whole",
    "superscript 24 glyph": "whole",
    "subscript 24 probe": "whole",
    "subscript 24 glyph": "whole",
    "superscript 24.5 probe": "whole",
    "superscript 24.5 glyph": "whole",
    "subscript 24.5 probe": "whole",
    "subscript 24.5 glyph": "whole",
    "superscript 25 probe": "whole",
    "superscript 25 glyph": "whole",
    "subscript 25 probe": "whole",
    "subscript 25 glyph": "whole",
    "superscript 25.5 probe": "whole",
    "superscript 25.5 glyph": "whole",
    "subscript 25.5 probe": "whole",
    "subscript 25.5 glyph": "whole",
    "superscript 26 probe": "whole",
    "superscript 26 glyph": "whole",
    "subscript 26 probe": "whole",
    "subscript 26 glyph": "whole",
    "superscript 26.5 probe": "whole",
    "superscript 26.5 glyph": "whole",
    "subscript 26.5 probe": "whole",
    "subscript 26.5 glyph": "whole",
    "superscript 27 probe": "whole",
    "superscript 27 glyph": "whole",
    "subscript 27 probe": "whole",
    "subscript 27 glyph": "whole",
    "superscript 27.5 probe": "whole",
    "superscript 27.5 glyph": "whole",
    "subscript 27.5 probe": "whole",
    "subscript 27.5 glyph": "whole",
    "superscript 28 probe": "whole",
    "superscript 28 glyph": "whole",
    "subscript 28 probe": "whole",
    "subscript 28 glyph": "whole",
    "superscript 28.5 probe": "whole",
    "superscript 28.5 glyph": "whole",
    "subscript 28.5 probe": "whole",
    "subscript 28.5 glyph": "whole",
    "superscript 29 probe": "whole",
    "superscript 29 glyph": "whole",
    "subscript 29 probe": "whole",
    "subscript 29 glyph": "whole",
    "superscript 29.5 probe": "whole",
    "superscript 29.5 glyph": "whole",
    "subscript 29.5 probe": "whole",
    "subscript 29.5 glyph": "whole",
    "superscript 30 probe": "whole",
    "superscript 30 glyph": "whole",
    "subscript 30 probe": "whole",
    "subscript 30 glyph": "whole",
    "superscript 30.5 probe": "whole",
    "superscript 30.5 glyph": "whole",
    "subscript 30.5 probe": "whole",
    "subscript 30.5 glyph": "whole",
    "superscript 31 probe": "whole",
    "superscript 31 glyph": "whole",
    "subscript 31 probe": "whole",
    "subscript 31 glyph": "whole",
    "superscript 31.5 probe": "whole",
    "superscript 31.5 glyph": "whole",
    "subscript 31.5 probe": "whole",
    "subscript 31.5 glyph": "whole",
    "superscript 32 probe": "whole",
    "superscript 32 glyph": "whole",
    "subscript 32 probe": "whole",
    "subscript 32 glyph": "whole",
    "superscript 32.5 probe": "whole",
    "superscript 32.5 glyph": "whole",
    "subscript 32.5 probe": "whole",
    "subscript 32.5 glyph": "whole",
    "superscript 33 probe": "whole",
    "superscript 33 glyph": "whole",
    "subscript 33 probe": "whole",
    "subscript 33 glyph": "whole",
    "superscript 33.5 probe": "whole",
    "superscript 33.5 glyph": "whole",
    "subscript 33.5 probe": "whole",
    "subscript 33.5 glyph": "whole",
    "superscript 34 probe": "whole",
    "superscript 34 glyph": "whole",
    "subscript 34 probe": "whole",
    "subscript 34 glyph": "whole",
    "superscript 34.5 probe": "whole",
    "superscript 34.5 glyph": "whole",
    "subscript 34.5 probe": "whole",
    "subscript 34.5 glyph": "whole",
    "superscript 35 probe": "whole",
    "superscript 35 glyph": "whole",
    "subscript 35 probe": "whole",
    "subscript 35 glyph": "whole",
    "superscript 35.5 probe": "whole",
    "superscript 35.5 glyph": "whole",
    "subscript 35.5 probe": "whole",
    "subscript 35.5 glyph": "whole",
    "superscript 36 probe": "whole",
    "superscript 36 glyph": "whole",
    "subscript 36 probe": "whole",
    "subscript 36 glyph": "whole",
    "superscript 36.5 probe": "whole",
    "superscript 36.5 glyph": "whole",
    "subscript 36.5 probe": "whole",
    "subscript 36.5 glyph": "whole",
    "superscript 37 probe": "whole",
    "superscript 37 glyph": "whole",
    "subscript 37 probe": "whole",
    "subscript 37 glyph": "whole",
    "superscript 37.5 probe": "whole",
    "superscript 37.5 glyph": "whole",
    "subscript 37.5 probe": "whole",
    "subscript 37.5 glyph": "whole",
    "superscript 38 probe": "whole",
    "superscript 38 glyph": "whole",
    "subscript 38 probe": "whole",
    "subscript 38 glyph": "whole",
    "superscript 38.5 probe": "whole",
    "superscript 38.5 glyph": "whole",
    "subscript 38.5 probe": "whole",
    "subscript 38.5 glyph": "whole",
    "superscript 39 probe": "whole",
    "superscript 39 glyph": "whole",
    "subscript 39 probe": "whole",
    "subscript 39 glyph": "whole",
    "superscript 39.5 probe": "whole",
    "superscript 39.5 glyph": "whole",
    "subscript 39.5 probe": "whole",
    "subscript 39.5 glyph": "whole",
    "superscript 40 probe": "whole",
    "superscript 40 glyph": "whole",
    "subscript 40 probe": "whole",
    "subscript 40 glyph": "whole",
}
"""The candidate sheetpeek's layout takes for each case (before each capture, the commit of its
predictions records them)."""


def _probe(ws: Any, cell: str, *, vertical: str = "bottom", text: str = PROBE) -> None:
    ws[cell] = text
    ws[cell].font = Font(name="Calibri", size=11)
    ws[cell].alignment = Alignment(vertical=vertical, wrap_text="\n" in text)


def _neighbor(ws: Any, cell: str, **alignment: Any) -> None:
    ws[cell] = PROBE
    ws[cell].font = NEIGHBOR
    if alignment:
        ws[cell].alignment = Alignment(**alignment)


def _cells(ws: Any) -> list[Case]:
    found: list[Case] = []

    def case(name: str, cell: str, last: str | None = None) -> None:
        found.append(Case("cells", name, cell, ("lines",), last))

    for row in range(4, 21):
        ws.row_dimensions[row].height = ROW
    for row in range(21, 25):
        ws.row_dimensions[row].height = SPREAD_ROW
    ws.column_dimensions["G"].hidden = True
    _probe(ws, "C4")
    case("probe alone", "C4")
    _probe(ws, "C5")
    _neighbor(ws, "E5")
    case("neighbor with a value", "C5")
    _probe(ws, "C6")
    ws["E6"].font = NEIGHBOR
    case("empty styled neighbor", "C6")
    _probe(ws, "C7")
    _neighbor(ws, "G7")
    case("neighbor in a hidden column", "C7")
    for row, name, alignment in (
        (8, "neighbor at 45 degrees", {"text_rotation": 45}),
        (9, "neighbor at 90 degrees", {"text_rotation": 90}),
        (10, "stacked neighbor", {"text_rotation": 255}),
    ):
        _probe(ws, f"C{row}")
        _neighbor(ws, f"E{row}", **alignment)
        case(name, f"C{row}")
    _neighbor(ws, "E11")
    ws.merge_cells("E11:E12")
    _probe(ws, "C11")
    _probe(ws, "C12")
    case("merge anchored above, first row", "C11")
    case("merge anchored above, second row", "C12")
    for row, kind in ((13, "superscript"), (14, "subscript")):
        _probe(ws, f"C{row}")
        ws[f"E{row}"] = PROBE
        ws[f"E{row}"].font = Font(name="Calibri", size=24, vertAlign=kind)
        case(f"{kind} neighbor", f"C{row}")
    ws.row_dimensions[15].font = NEIGHBOR
    _probe(ws, "C15")
    case("row style", "C15")
    _probe(ws, "C16")
    ws["E16"] = CellRichText(
        TextBlock(InlineFont(rFont="Calibri", sz=11), "HH"),
        TextBlock(InlineFont(rFont="Calibri", sz=24), "HH"),
    )
    case("neighbor with a large run", "C16")
    for first, last, at, name in (
        (17, 18, 17, "merged probe, neighbor in its first row"),
        (19, 20, 20, "merged probe, neighbor in its last row"),
    ):
        _probe(ws, f"C{first}")
        ws.merge_cells(f"C{first}:C{last}")
        _neighbor(ws, f"E{at}")
        case(name, f"C{first}", f"C{last}")
    lines = "\n".join([PROBE] * 3)
    for row, vertical, neighbor in (
        (21, "distributed", True),
        (22, "justify", True),
        (23, "distributed", False),
        (24, "justify", False),
    ):
        _probe(ws, f"C{row}", vertical=vertical, text=lines)
        if neighbor:
            _neighbor(ws, f"E{row}")
        where = "beside the neighbor" if neighbor else "alone"
        case(f"{vertical} {where}", f"C{row}")
    return found


def _column_style(ws: Any) -> list[Case]:
    ws.column_dimensions["E"].font = NEIGHBOR
    ws.row_dimensions[4].height = ROW
    _probe(ws, "C4")
    _probe(ws, "C5")
    return [
        Case("column_style", "column style, fixed row", "C4", ("lines",)),
        Case("column_style", "column style, fitted row", "C5", ("lines",)),
    ]


def _scripts(ws: Any, group: str) -> list[Case]:
    k = SCRIPT_GROUPS.index(group)
    found = []
    for row, (kind, size) in enumerate(
        SCRIPT_ROWS[k * SCRIPT_GROUP_ROWS : (k + 1) * SCRIPT_GROUP_ROWS], 4
    ):
        ws.row_dimensions[row].height = SCRIPT_ROW
        _probe(ws, f"C{row}")
        ws[f"E{row}"] = PROBE
        ws[f"E{row}"].font = Font(name="Calibri", size=size, vertAlign=kind)
        found.append(Case(group, f"{kind} {size:g} probe", f"C{row}", ("lines",)))
        found.append(Case(group, f"{kind} {size:g} glyph", f"E{row}", ("lines",)))
    return found


def fill(ws: Any, group: str) -> list[Case]:
    """Lay one group out on the first worksheet of its own workbook; its cases."""
    if group in SCRIPT_GROUPS:
        return _scripts(ws, group)
    return _cells(ws) if group == "cells" else _column_style(ws)


def cases(group: str | None = None) -> list[Case]:
    """Every case of ``group``, or of every group, in order."""
    chosen = GROUPS if group is None else (group,)
    return [case for g in chosen for case in fill(openpyxl.Workbook().active, g)]


def _boxes(found: list[Case]) -> list[str]:
    cells = [c for case in found for c in (case.cell, case.last) if c]
    return list(dict.fromkeys(cells))


def references(group: str) -> list[str]:
    """The cells whose geometry is read: row 1 for each column, column A for each row."""
    columns, rows = {}, {}
    for cell in _boxes(cases(group)):
        letters, row = coordinate_from_string(cell)
        columns[f"{letters}1"] = None
        rows[f"A{row}"] = None
    return [*columns, *rows]


def complete(cap: Any, cells: list[str]) -> None:
    """Give each of ``cells`` a box from its column's row-1 cell and its row's column-A cell."""
    for cell in cells:
        if cell in cap.geometry:
            continue
        letters, row = coordinate_from_string(cell)
        column, line = cap.geometry[f"{letters}1"], cap.geometry[f"A{row}"]
        cap.geometry[cell] = {
            "left": column["left"],
            "width": column["width"],
            "top": line["top"],
            "height": line["height"],
        }


def build(directory: Path) -> dict[str, Path]:
    """Write each group to its own workbook in ``directory``; the paths by group."""
    directory.mkdir(parents=True, exist_ok=True)
    paths = {}
    for group in GROUPS:
        wb = openpyxl.Workbook()
        wb.active.title = group
        fill(wb.active, group)
        paths[group] = directory / f"{group}.xlsx"
        wb.save(paths[group])
    return paths


def run(groups: tuple[str, ...] = GROUPS) -> None:
    """Capture each group's window, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    for group in groups:
        capture(
            lambda ws, group=group: fill(ws, group),
            cells=references(group),
            out=BUILD_DIR / group,
            markers=MARKERS[group],
            zooms=(100,),
            name=f"row_bottoms_{group}",
        )
    write()


def baselines(line_tops: str) -> list[float]:
    """Each line's baseline from ``line_tops`` ("top+height" spans): the bottom of its ink."""
    found = []
    for span in line_tops.split():
        top, height = span.split("+")
        found.append(float(top) + float(height))
    return found


def measure(cap: Any, case: Case) -> dict[str, str] | None:
    """One case's measurements, or ``None`` when it is outside the window."""
    complete(cap, _boxes([case]))
    found = text_style.measure(cap, case)
    if found is not None:
        found["baselines"] = " ".join(f"{b:g}" for b in baselines(found.get("line_tops", "")))
    return found


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``row_bottoms.tsv``."""
    from capture.screen.run import header, load

    records = []
    first: dict[str, str] = {}
    for group in GROUPS:
        out = BUILD_DIR / group
        if not (out / "formats.tsv").exists():
            print(f"{group}: not captured yet, left out", file=sys.stderr)
            continue
        first = first or header(out)
        found = fill(openpyxl.Workbook().active, group)
        for cap in load(out):
            for case in found:
                measured = measure(cap, case)
                if measured is None:
                    continue
                row = coordinate_from_string(case.cell)[1]
                records.append(
                    {
                        "group": case.group,
                        "name": case.name,
                        "zoom": str(cap.zoom),
                        "cell": case.cell,
                        "row_height": cap.value(f"A{row}", "geometry.height") or "",
                        **measured,
                    }
                )
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after", "zooms"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: row_bottoms\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        paths = build(Path(argv[1]))
        print(f"{len(paths)} workbooks in {argv[1]}")
        return 0
    if argv[:1] == ["run"] and all(group in GROUPS for group in argv[1:]):
        run(tuple(argv[1:]) or GROUPS)
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
