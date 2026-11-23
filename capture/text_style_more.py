"""Capture text styles at other sizes and fonts from the screen (route C), after ``text_style``.

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.text_style_more run
    uv run --group capture python -m capture.text_style_more run rules shrunk  # some groups only
    uv run --group capture python -m capture.text_style_more read     # the saved captures again
    uv run --group capture python -m capture.text_style_more build PATH   # the workbook only

The suite ``text_style`` captured Calibri 11 (and a few 24 pt cases); this one takes the same
rules to other sizes and faces, and settles what that capture left mixed. ARCHITECTURE.md §4.13
(text styles: superscript and subscript, the rules, accounting underlines, shrink to fit, the
spreads in fitting), §4.7 (indent and wrapping) and §3.6 (baselines of fonts xlformat has not
measured, lines that do not fit their block). Six groups, each its own worksheet, workbook and
window, in the dark appearance as the other suites; rows are fitted by Excel on opening unless
a case gives a height. Texts are "H" (no descender: its ink ends on the baseline), "Hxg" or "H"
repeated and "g" (the g's descender shows the line's box below the baseline).

- ``vert``, zoom 100. Rows 4 to 15: "HHH" in C, one case a row, plain, superscript and subscript
  at 9, 14, 18 and 36 pt: the fitted row's height (Excel's read) and the baseline (§4.13,
  growth of a fitted row at other sizes; raise and lower at other sizes). Rows 17 to 19: a
  superscript in C and a subscript in E in one row at 14, 24 and 36 pt (a row with both). Rows
  21 to 26: numbers of 8 to 12 digits (11 pt) and 4 to 8 digits (22 pt) in default columns C,
  E, G, I and K, superscript, subscript and plain, format ``0``: the text Excel shows (digits or
  ``#``) and its ink (superscript measured at its drawn size for ``####``).
- ``rules``, zoom 100 and 200. Rows 4 to 10 at 9, 14, 20 and 36 pt: "HHH" with a single,
  double, single accounting and double accounting underline and a strikethrough (C, E, G, I, K,
  14 characters): each rule's offset from the baseline, its thickness and its ends (offsets
  scaled from 11 pt). Row 12: "HH" single and double accounting in the merges C12:E12 and
  G12:I12; row 14 three wrapped lines and row 16 "aa bb cc" distributed, single (C) and double
  (E) accounting (a merge spans the merge, a line per line, one line under distributed words).
  Rows 18, 20 and 22, fixed at 15, 20 and 30 pt: single accounting "HHH" (C) beside plain "HHH"
  (E), bottom aligned: the baseline (6 pt above the bottom, also in a fixed row).
- ``steps``, zoom 100. Rows 4 to 22: "H" repeated as many times as the row's number, left
  aligned with shrink to fit in default columns at Calibri 9 (C), 16 (E), 20 (G) and Aptos Narrow
  11 (I), and unshrunk in wide columns K, M, O and Q: the ink width and height over the
  reference's, so the step ``k/7`` at those sizes (§4.13, shrink step).
- ``shrunk``, zoom 100 and 200. Rows 4, 6 and 8, fixed at 15, 30 and 60 pt: top (C-F), center
  (G-J) and bottom (K-N) aligned, "HHHg" unshrunk and three texts of "H"s and a "g" that shrink
  to 6/7, 4/7 and 2/7: the baseline (the bottom of the ink's left quarter, all "H") and the
  descender's bottom. The baseline at 2/7 and in a 26 pt row did not fit "shrunk text keeps the
  unshrunk baseline" in ``text_style`` (C12 and E12 end at 21 pt, not 23; C17-C21 at 11, not
  12): row 10, fixed at 26 pt, holds the 4/7 text alone with plain 11 pt "Hg"; row 12 the same
  beside a 20 pt "Hg" (does a bottom-aligned 11 pt line share a 20 pt neighbor's baseline?);
  row 14 that row fitted (26 pt, as in ``text_style``); row 16 the 4/7 text beside a shrunk 20 pt
  one, fitted. Row 18: 60 "H"s with shrink in a default column (C) and a 2-character one (O):
  where even 1/7 does not fit. Rows 20 and 22, fixed at 15 and 20 pt: three wrapped lines
  centered (C), bottom (E) and top (G) aligned: where lines that do not fit their block start.
- ``wrap``, zoom 100. Rows 4 to 10: two words measuring 31 to 49 pt (``INDENT_NARROW``) in
  default columns C-J, left (C, E, G, I) and right (D, F, H, J) aligned with indent 0 to 3, all
  wrapped; 98 to 116 pt (``INDENT_WIDE``) the same way in columns L-S of 20 characters: the
  lines and the ink, so whether an indent narrows the width a wrapped line may take (§4.7).
  Rows 13 to 21, each alone in a fitted row: horizontal justify with wrap on (C13) and off (C15),
  vertical justify (C17) and distributed (C19) with wrap off, and plain wrap (C21), a long text:
  the fitted row's height (do justify and the vertical spreads count in fitting?).
- ``faces``, zoom 100. Rows 4, 6 and 8 (60 pt) top, center and bottom aligned "Hxg", rows 10, 12
  and 14 (90 pt) three wrapped lines of it, in Calibri 13 (C) and 17 (E), Arial (G), Times New
  Roman (I), Courier New (K), all 11 pt but Calibri's, and Calibri 11 (M) as the control: each
  line's baseline (sizes and faces xlformat has not measured take the nearest measured size,
  and Calibri's, §3.6).
- ``overflow``, zoom 100. Rows 4 to 21, fixed at 15 to 32 pt: "Haa", "Hbb", "Hcc" wrapped in
  Calibri 11, centered (C), bottom (E) and top (G), and "Haa", "Hbb" centered (I) and bottom
  (K): which lines Excel places as the block when they do not fit (§3.6). Rows 23 and 25, fixed
  at 15 pt: one line "Hg" in Calibri 14 (row 23) and 20 (row 25), bottom (C), center (E) and top
  (G): a line taller than its row. Row 27, fixed at 20 pt: "aaa", "bbb", "ccc" vertical justify
  (C) and distributed (E): the spread alignments when the lines do not fit.

Geometry is read for one cell a column (row 1) and one a row (column A, which gives the row's
height); each case's box is put together from them, so few cells are read. ``vert``'s numbers
also have their text read. A cell outside the window at zoom 200 is skipped there. ``read``
writes ``tests/captures/text_style_more.tsv``: one row per case and zoom, measured as ``text_style``
measures (its code, by import; its columns of the same names), the row's height and the shown
text (Excel's reads), and ``baselines``, each line's baseline from
the cell's top (``text_style``'s ``baseline`` is the last line's, rules left out), and
``rule_tops``, each rule's top from the cell's top. Lengths in points at the zoom.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils.cell import coordinate_from_string

from capture import text_style
from capture.text_style import Case

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "text_style_more"
OUTPUT = ROOT / "tests" / "captures" / "text_style_more.tsv"
GROUPS = ("vert", "rules", "steps", "shrunk", "wrap", "faces", "overflow")
ZOOMS = {
    "vert": (100,), "rules": (100, 200), "steps": (100,), "shrunk": (100, 200), "wrap": (100,),
    "faces": (100,), "overflow": (100,),
}  # fmt: skip
MARKERS = {
    "vert": ("B2", "N3"), "rules": ("B2", "M3"), "steps": ("B2", "S3"), "shrunk": ("B2", "Q3"),
    "wrap": ("B2", "U3"), "faces": ("B2", "O3"), "overflow": ("B2", "M3"),
}  # fmt: skip
"""The second marker sits right of every case, in row 3; at zoom 200 it may be out of the
window, and the grid is then measured from B2 and the nominal scale (``CAPTURING.md``)."""
COLUMNS = (
    "group", "name", "zoom", "cell", "row_height", "text", "lines", "line_tops", "baselines",
    "ink_left", "ink_right", "ink_top", "ink_bottom", "baseline", "glyph_top", "rule_offset",
    "rule_tops", "rule_thickness", "rule_start", "rule_end", "width_ratio", "height_ratio",
    "ink_width", "ref_width",
)  # fmt: skip

SIZES = (9, 14, 18, 36)
"""Superscript and subscript alone (``vert``)."""
BOTH_SIZES = (14, 24, 36)
RULE_SIZES = (9, 14, 20, 36)
RULE_KINDS = ("single", "double", "singleAccounting", "doubleAccounting", "strike")
STEP_FONTS = (("Calibri", 9), ("Calibri", 16), ("Calibri", 20), ("Aptos Narrow", 11))
STEP_WIDTHS = (28, 45, 53, 32)
"""The reference columns' widths in characters: 22 "H"s unshrunk with room to spare."""
FACES = (
    ("Calibri", 13), ("Calibri", 17), ("Arial", 11), ("Times New Roman", 11),
    ("Courier New", 11), ("Calibri", 11),
)  # fmt: skip
SHRUNK = {7: "HHHg", 6: "H" * 7 + "g", 4: "H" * 10 + "g", 2: "H" * 16 + "g"}
"""Texts of "H"s and a "g" and the step ``k/7`` sheetpeek shrinks each to in a default column."""
THREE = "aaa\nbbb\nccc"
HXG = "Hxg"
INDENT_NARROW = {
    31: "aaaa iii", 37: "aaaa biii", 40: "aaaa bbii", 43: "aaaa baaa", 46: "aaaa baaai",
    47: "aaaa bbaai", 49: "aaaa bbaaa",
}  # fmt: skip
"""Two words and their width with the space, in points at Calibri 11 (xlformat's tables): for a
default column, whose line may take 48 pt, less 6 pt an indent level if the indent narrows it."""
INDENT_WIDE = {
    98: "aaaaaaaaaa bbbaaaaai", 104: "aaaaaaaaaa bbbbaaaaai", 107: "aaaaaaaaaa bbbbbaaaaa",
    110: "aaaaaaaaaa bbbbbaaaaai", 113: "aaaaaaaaaa bbbbbaaaaaii",
    114: "aaaaaaaaaa bbbbbbaaaaii", 116: "aaaaaaaaaa bbbbbaaaaaiii",
}  # fmt: skip
"""The same for a column of 20 characters, whose line may take 115 pt."""
TEXT_CELLS = {"vert": tuple(f"{c}{r}" for r in range(21, 27) for c in "CEGIK")}
"""Cells whose shown text is read: the numbers, which may show ``#``."""

# fmt: off
EXPECTED: dict[str, dict[str, dict[str, Any]]] = {
    "vert": {
        "C4": {"row_height": 15, "text": "HHH", "baselines": (12,)},
        "C5": {"row_height": 15, "text": "HHH", "baselines": (8,)},
        "C6": {"row_height": 15, "text": "HHH", "baselines": (14,)},
        "C7": {"row_height": 19, "text": "HHH", "baselines": (15,)},
        "C8": {"row_height": 22, "text": "HHH", "baselines": (11,)},
        "C9": {"row_height": 22, "text": "HHH", "baselines": (21,)},
        "C10": {"row_height": 24, "text": "HHH", "baselines": (19,)},
        "C11": {"row_height": 28, "text": "HHH", "baselines": (14,)},
        "C12": {"row_height": 28, "text": "HHH", "baselines": (26,)},
        "C13": {"row_height": 47, "text": "HHH", "baselines": (37,)},
        "C14": {"row_height": 57, "text": "HHH", "baselines": (30,)},
        "C15": {"row_height": 57, "text": "HHH", "baselines": (53,)},
        "C17": {"row_height": 24, "text": "HHH", "baselines": (13,)},
        "E17": {"row_height": 24, "text": "HHH", "baselines": (23,)},
        "C18": {"row_height": 40, "text": "HHH", "baselines": (22,)},
        "E18": {"row_height": 40, "text": "HHH", "baselines": (38,)},
        "C19": {"row_height": 62, "text": "HHH", "baselines": (35,)},
        "E19": {"row_height": 62, "text": "HHH", "baselines": (58,)},
        "C21": {"row_height": 17, "text": "12345678", "baselines": (9,)},
        "E21": {"row_height": 17, "text": "123456789", "baselines": (9,)},
        "G21": {"row_height": 17, "text": "1234567890", "baselines": (9,)},
        "I21": {"row_height": 17, "text": "12345678901", "baselines": (9,)},
        "K21": {"row_height": 17, "text": "123456789012", "baselines": (9,)},
        "C22": {"row_height": 17, "text": "12345678", "baselines": (16,)},
        "E22": {"row_height": 17, "text": "123456789", "baselines": (16,)},
        "G22": {"row_height": 17, "text": "1234567890", "baselines": (16,)},
        "I22": {"row_height": 17, "text": "12345678901", "baselines": (16,)},
        "K22": {"row_height": 17, "text": "123456789012", "baselines": (16,)},
        "C23": {"row_height": 15, "text": "12345678", "baselines": (12,)},
        "E23": {"row_height": 15, "text": "#########", "baselines": (12,)},
        "G23": {"row_height": 15, "text": "#########", "baselines": (12,)},
        "I23": {"row_height": 15, "text": "#########", "baselines": (12,)},
        "K23": {"row_height": 15, "text": "#########", "baselines": (12,)},
        "C24": {"row_height": 34, "text": "1234", "baselines": (17,)},
        "E24": {"row_height": 34, "text": "12345", "baselines": (17,)},
        "G24": {"row_height": 34, "text": "123456", "baselines": (17,)},
        "I24": {"row_height": 34, "text": "######", "baselines": (17,)},
        "K24": {"row_height": 34, "text": "######", "baselines": (17,)},
        "C25": {"row_height": 34, "text": "1234", "baselines": (32,)},
        "E25": {"row_height": 34, "text": "12345", "baselines": (32,)},
        "G25": {"row_height": 34, "text": "123456", "baselines": (32,)},
        "I25": {"row_height": 34, "text": "######", "baselines": (32,)},
        "K25": {"row_height": 34, "text": "######", "baselines": (32,)},
        "C26": {"row_height": 29, "text": "1234", "baselines": (23,)},
        "E26": {"row_height": 29, "text": "####", "baselines": (23,)},
        "G26": {"row_height": 29, "text": "####", "baselines": (23,)},
        "I26": {"row_height": 29, "text": "####", "baselines": (23,)},
        "K26": {"row_height": 29, "text": "####", "baselines": (23,)},
    },
    "rules": {
        "C4": {"baselines": (12,), "rules": (("underline", 1, 1, 13, 3, -62, 1),)},
        "E4": {"baselines": (12,), "rules": (("underline", 2, -1, 11, 3, -62, 1),)},
        "G4": {"baselines": (12,), "rules": (("underline", 1, 2, 14, 3, -2, 1),)},
        "I4": {"baselines": (12,), "rules": (("underline", 2, 2, 14, 3, -2, 1),)},
        "K4": {"baselines": (12,), "rules": (("strike", 1, -2, 10, 3, -63, 1),)},
        "C6": {"baselines": (15,), "rules": (("underline", 1, 1, 16, 3, -53, 1),)},
        "E6": {"baselines": (15,), "rules": (("underline", 2, 1, 16, 3, -53, 1),)},
        "G6": {"baselines": (15,), "rules": (("underline", 1, 4, 19, 3, -2, 1),)},
        "I6": {"baselines": (15,), "rules": (("underline", 2, 4, 19, 3, -2, 1),)},
        "K6": {"baselines": (15,), "rules": (("strike", 1, -4, 11, 3, -54, 1),)},
        "C8": {"baselines": (21,), "rules": (("underline", 1, 2, 23, 3, -44, 1),)},
        "E8": {"baselines": (21,), "rules": (("underline", 2, 2, 23, 3, -44, 1),)},
        "G8": {"baselines": (21,), "rules": (("underline", 1, 5, 26, 3, -2, 1),)},
        "I8": {"baselines": (21,), "rules": (("underline", 2, 5, 26, 3, -2, 1),)},
        "K8": {"baselines": (21,), "rules": (("strike", 1, -5, 16, 3, -45, 1),)},
        "C10": {"baselines": (37,), "rules": (("underline", 1, 4, 41, 3, -14, 2),)},
        "E10": {"baselines": (37,), "rules": (("underline", 2, 4, 41, 3, -14, 2),)},
        "G10": {"baselines": (37,), "rules": (("underline", 1, 10, 47, 3, -2, 2),)},
        "I10": {"baselines": (37,), "rules": (("underline", 2, 10, 47, 3, -2, 2),)},
        "K10": {"baselines": (37,), "rules": (("strike", 1, -9, 28, 3, -15, 2),)},
        "C12": {"baselines": (12,), "rules": (("underline", 1, 3, 15, 3, -2, 1),)},
        "G12": {"baselines": (12,), "rules": (("underline", 2, 3, 15, 3, -2, 1),)},
        "C14": {"baselines": (15, 33, 51), "rules": (("underline", 1, 3, 54, 3, -2, 1),)},
        "C16": {"baselines": (13,), "rules": (("underline", 1, 3, 16, 3, -2, 1),)},
        "E14": {"baselines": (15, 33, 51), "rules": (("underline", 2, 3, 54, 3, -2, 1),)},
        "E16": {"baselines": (13,), "rules": (("underline", 2, 3, 16, 3, -2, 1),)},
        "C18": {"baselines": (9,), "rules": (("underline", 1, 3, 12, 3, -2, 1),)},
        "E18": {"baselines": (9,), "rules": ()},
        "C20": {"baselines": (14,), "rules": (("underline", 1, 3, 17, 3, -2, 1),)},
        "E20": {"baselines": (14,), "rules": ()},
        "C22": {"baselines": (24,), "rules": (("underline", 1, 3, 27, 3, -2, 1),)},
        "E22": {"baselines": (24,), "rules": ()},
    },
    "steps": {
        "C4": {"scale": 7},
        "E4": {"scale": 7},
        "G4": {"scale": 7},
        "I4": {"scale": 7},
        "C5": {"scale": 7},
        "E5": {"scale": 6},
        "G5": {"scale": 5},
        "I5": {"scale": 7},
        "C6": {"scale": 7},
        "E6": {"scale": 5},
        "G6": {"scale": 4},
        "I6": {"scale": 7},
        "C7": {"scale": 7},
        "E7": {"scale": 4},
        "G7": {"scale": 3},
        "I7": {"scale": 6},
        "C8": {"scale": 7},
        "E8": {"scale": 4},
        "G8": {"scale": 3},
        "I8": {"scale": 6},
        "C9": {"scale": 6},
        "E9": {"scale": 3},
        "G9": {"scale": 2},
        "I9": {"scale": 5},
        "C10": {"scale": 5},
        "E10": {"scale": 2},
        "G10": {"scale": 2},
        "I10": {"scale": 4},
        "C11": {"scale": 5},
        "E11": {"scale": 2},
        "G11": {"scale": 2},
        "I11": {"scale": 4},
        "C12": {"scale": 5},
        "E12": {"scale": 2},
        "G12": {"scale": 2},
        "I12": {"scale": 4},
        "C13": {"scale": 4},
        "E13": {"scale": 2},
        "G13": {"scale": 1},
        "I13": {"scale": 3},
        "C14": {"scale": 4},
        "E14": {"scale": 2},
        "G14": {"scale": 1},
        "I14": {"scale": 3},
        "C15": {"scale": 4},
        "E15": {"scale": 2},
        "G15": {"scale": 1},
        "I15": {"scale": 3},
        "C16": {"scale": 4},
        "E16": {"scale": 2},
        "G16": {"scale": 1},
        "I16": {"scale": 3},
        "C17": {"scale": 2},
        "E17": {"scale": 1},
        "G17": {"scale": 1},
        "I17": {"scale": 2},
        "C18": {"scale": 2},
        "E18": {"scale": 1},
        "G18": {"scale": 1},
        "I18": {"scale": 2},
        "C19": {"scale": 2},
        "E19": {"scale": 1},
        "G19": {"scale": 1},
        "I19": {"scale": 2},
        "C20": {"scale": 2},
        "E20": {"scale": 1},
        "G20": {"scale": 1},
        "I20": {"scale": 2},
        "C21": {"scale": 2},
        "E21": {"scale": 1},
        "G21": {"scale": 1},
        "I21": {"scale": 2},
        "C22": {"scale": 2},
        "E22": {"scale": 1},
        "G22": {"scale": 1},
        "I22": {"scale": 2},
    },
    "shrunk": {
        "C4": {"row_height": 15, "baselines": (12,), "left": 3, "scale": 7},
        "D4": {"row_height": 15, "baselines": (11,), "left": 3, "scale": 6},
        "E4": {"row_height": 15, "baselines": (8,), "left": 3, "scale": 4},
        "F4": {"row_height": 15, "baselines": (5,), "left": 3, "scale": 2},
        "G4": {"row_height": 15, "baselines": (12,), "left": 3, "scale": 7},
        "H4": {"row_height": 15, "baselines": (11,), "left": 3, "scale": 6},
        "I4": {"row_height": 15, "baselines": (10,), "left": 3, "scale": 4},
        "J4": {"row_height": 15, "baselines": (8,), "left": 3, "scale": 2},
        "K4": {"row_height": 15, "baselines": (12,), "left": 3, "scale": 7},
        "L4": {"row_height": 15, "baselines": (12,), "left": 3, "scale": 6},
        "M4": {"row_height": 15, "baselines": (12,), "left": 3, "scale": 4},
        "N4": {"row_height": 15, "baselines": (12,), "left": 3, "scale": 2},
        "C6": {"row_height": 30, "baselines": (12,), "left": 3, "scale": 7},
        "D6": {"row_height": 30, "baselines": (11,), "left": 3, "scale": 6},
        "E6": {"row_height": 30, "baselines": (8,), "left": 3, "scale": 4},
        "F6": {"row_height": 30, "baselines": (5,), "left": 3, "scale": 2},
        "G6": {"row_height": 30, "baselines": (19,), "left": 3, "scale": 7},
        "H6": {"row_height": 30, "baselines": (19,), "left": 3, "scale": 6},
        "I6": {"row_height": 30, "baselines": (18,), "left": 3, "scale": 4},
        "J6": {"row_height": 30, "baselines": (16,), "left": 3, "scale": 2},
        "K6": {"row_height": 30, "baselines": (27,), "left": 3, "scale": 7},
        "L6": {"row_height": 30, "baselines": (27,), "left": 3, "scale": 6},
        "M6": {"row_height": 30, "baselines": (27,), "left": 3, "scale": 4},
        "N6": {"row_height": 30, "baselines": (27,), "left": 3, "scale": 2},
        "C8": {"row_height": 60, "baselines": (12,), "left": 3, "scale": 7},
        "D8": {"row_height": 60, "baselines": (11,), "left": 3, "scale": 6},
        "E8": {"row_height": 60, "baselines": (8,), "left": 3, "scale": 4},
        "F8": {"row_height": 60, "baselines": (5,), "left": 3, "scale": 2},
        "G8": {"row_height": 60, "baselines": (34,), "left": 3, "scale": 7},
        "H8": {"row_height": 60, "baselines": (34,), "left": 3, "scale": 6},
        "I8": {"row_height": 60, "baselines": (33,), "left": 3, "scale": 4},
        "J8": {"row_height": 60, "baselines": (31,), "left": 3, "scale": 2},
        "K8": {"row_height": 60, "baselines": (57,), "left": 3, "scale": 7},
        "L8": {"row_height": 60, "baselines": (57,), "left": 3, "scale": 6},
        "M8": {"row_height": 60, "baselines": (57,), "left": 3, "scale": 4},
        "N8": {"row_height": 60, "baselines": (57,), "left": 3, "scale": 2},
        "C10": {"row_height": 26, "baselines": (23,), "left": 3, "scale": 4},
        "E10": {"row_height": 26, "baselines": (23,), "left": 3, "scale": 7},
        "C12": {"row_height": 26, "baselines": (21,), "left": 3, "scale": 4},
        "E12": {"row_height": 26, "baselines": (21,), "left": 3, "scale": 7},
        "G12": {"row_height": 26, "baselines": (21,), "left": 3, "scale": 7},
        "C14": {"row_height": 26, "baselines": (21,), "left": 3, "scale": 4},
        "E14": {"row_height": 26, "baselines": (21,), "left": 3, "scale": 7},
        "G14": {"row_height": 26, "baselines": (21,), "left": 3, "scale": 7},
        "C16": {"row_height": 26, "baselines": (21,), "left": 3, "scale": 4},
        "E16": {"row_height": 26, "baselines": (21,), "left": 3, "scale": 2},
        "C18": {"row_height": 15, "baselines": (11,), "left": 3, "scale": 1},
        "O18": {"row_height": 15, "baselines": (11,), "left": 3, "scale": 1},
        "C20": {"row_height": 15, "baselines": (12, 27, 42), "left": 3, "scale": 7},
        "E20": {"row_height": 15, "baselines": (12, 27, 42), "left": 3, "scale": 7},
        "G20": {"row_height": 15, "baselines": (12, 27, 42), "left": 3, "scale": 7},
        "C22": {"row_height": 20, "baselines": (7, 22, 37), "left": 3, "scale": 7},
        "E22": {"row_height": 20, "baselines": (2, 17, 32), "left": 3, "scale": 7},
        "G22": {"row_height": 20, "baselines": (12, 27, 42), "left": 3, "scale": 7},
    },
    "wrap": {
        "C4": {"row_height": 32, "lines": 1, "left": 3},
        "D4": {"row_height": 32, "lines": 1, "left": 20},
        "E4": {"row_height": 32, "lines": 1, "left": 9},
        "F4": {"row_height": 32, "lines": 1, "left": 14},
        "G4": {"row_height": 32, "lines": 1, "left": 15},
        "H4": {"row_height": 32, "lines": 1, "left": 8},
        "I4": {"row_height": 32, "lines": 2, "left": 21},
        "J4": {"row_height": 32, "lines": 2, "left": 13},
        "L4": {"row_height": 32, "lines": 1, "left": 3},
        "M4": {"row_height": 32, "lines": 1, "left": 20},
        "N4": {"row_height": 32, "lines": 1, "left": 9},
        "O4": {"row_height": 32, "lines": 1, "left": 14},
        "P4": {"row_height": 32, "lines": 1, "left": 15},
        "Q4": {"row_height": 32, "lines": 1, "left": 8},
        "R4": {"row_height": 32, "lines": 2, "left": 21},
        "S4": {"row_height": 32, "lines": 2, "left": 50},
        "C5": {"row_height": 32, "lines": 1, "left": 3},
        "D5": {"row_height": 32, "lines": 1, "left": 14},
        "E5": {"row_height": 32, "lines": 1, "left": 9},
        "F5": {"row_height": 32, "lines": 1, "left": 8},
        "G5": {"row_height": 32, "lines": 2, "left": 15},
        "H5": {"row_height": 32, "lines": 2, "left": 19},
        "I5": {"row_height": 32, "lines": 2, "left": 21},
        "J5": {"row_height": 32, "lines": 2, "left": 13},
        "L5": {"row_height": 32, "lines": 1, "left": 3},
        "M5": {"row_height": 32, "lines": 1, "left": 14},
        "N5": {"row_height": 32, "lines": 1, "left": 9},
        "O5": {"row_height": 32, "lines": 1, "left": 8},
        "P5": {"row_height": 32, "lines": 2, "left": 15},
        "Q5": {"row_height": 32, "lines": 2, "left": 54},
        "R5": {"row_height": 32, "lines": 2, "left": 21},
        "S5": {"row_height": 32, "lines": 2, "left": 48},
        "C6": {"row_height": 32, "lines": 1, "left": 3},
        "D6": {"row_height": 32, "lines": 1, "left": 11},
        "E6": {"row_height": 32, "lines": 1, "left": 9},
        "F6": {"row_height": 32, "lines": 1, "left": 5},
        "G6": {"row_height": 32, "lines": 2, "left": 15},
        "H6": {"row_height": 32, "lines": 2, "left": 19},
        "I6": {"row_height": 32, "lines": 2, "left": 21},
        "J6": {"row_height": 32, "lines": 2, "left": 13},
        "L6": {"row_height": 32, "lines": 1, "left": 3},
        "M6": {"row_height": 32, "lines": 1, "left": 11},
        "N6": {"row_height": 32, "lines": 1, "left": 9},
        "O6": {"row_height": 32, "lines": 1, "left": 5},
        "P6": {"row_height": 32, "lines": 2, "left": 15},
        "Q6": {"row_height": 32, "lines": 2, "left": 51},
        "R6": {"row_height": 32, "lines": 2, "left": 21},
        "S6": {"row_height": 32, "lines": 2, "left": 45},
        "C7": {"row_height": 32, "lines": 1, "left": 3},
        "D7": {"row_height": 32, "lines": 1, "left": 8},
        "E7": {"row_height": 32, "lines": 2, "left": 9},
        "F7": {"row_height": 32, "lines": 2, "left": 24},
        "G7": {"row_height": 32, "lines": 2, "left": 15},
        "H7": {"row_height": 32, "lines": 2, "left": 18},
        "I7": {"row_height": 32, "lines": 2, "left": 21},
        "J7": {"row_height": 32, "lines": 2, "left": 12},
        "L7": {"row_height": 32, "lines": 1, "left": 3},
        "M7": {"row_height": 32, "lines": 1, "left": 8},
        "N7": {"row_height": 32, "lines": 2, "left": 9},
        "O7": {"row_height": 32, "lines": 2, "left": 54},
        "P7": {"row_height": 32, "lines": 2, "left": 15},
        "Q7": {"row_height": 32, "lines": 2, "left": 48},
        "R7": {"row_height": 32, "lines": 2, "left": 21},
        "S7": {"row_height": 32, "lines": 2, "left": 42},
        "C8": {"row_height": 32, "lines": 1, "left": 3},
        "D8": {"row_height": 32, "lines": 1, "left": 5},
        "E8": {"row_height": 32, "lines": 2, "left": 9},
        "F8": {"row_height": 32, "lines": 2, "left": 21},
        "G8": {"row_height": 32, "lines": 2, "left": 15},
        "H8": {"row_height": 32, "lines": 2, "left": 15},
        "I8": {"row_height": 32, "lines": 2, "left": 21},
        "J8": {"row_height": 32, "lines": 2, "left": 9},
        "L8": {"row_height": 32, "lines": 1, "left": 3},
        "M8": {"row_height": 32, "lines": 1, "left": 5},
        "N8": {"row_height": 32, "lines": 2, "left": 9},
        "O8": {"row_height": 32, "lines": 2, "left": 51},
        "P8": {"row_height": 32, "lines": 2, "left": 15},
        "Q8": {"row_height": 32, "lines": 2, "left": 45},
        "R8": {"row_height": 32, "lines": 2, "left": 21},
        "S8": {"row_height": 32, "lines": 2, "left": 39},
        "C9": {"row_height": 32, "lines": 1, "left": 3},
        "D9": {"row_height": 32, "lines": 1, "left": 4},
        "E9": {"row_height": 32, "lines": 2, "left": 9},
        "F9": {"row_height": 32, "lines": 2, "left": 20},
        "G9": {"row_height": 32, "lines": 2, "left": 15},
        "H9": {"row_height": 32, "lines": 2, "left": 14},
        "I9": {"row_height": 32, "lines": 2, "left": 21},
        "J9": {"row_height": 32, "lines": 2, "left": 8},
        "L9": {"row_height": 32, "lines": 1, "left": 3},
        "M9": {"row_height": 32, "lines": 1, "left": 4},
        "N9": {"row_height": 32, "lines": 2, "left": 9},
        "O9": {"row_height": 32, "lines": 2, "left": 50},
        "P9": {"row_height": 32, "lines": 2, "left": 15},
        "Q9": {"row_height": 32, "lines": 2, "left": 44},
        "R9": {"row_height": 32, "lines": 2, "left": 21},
        "S9": {"row_height": 32, "lines": 2, "left": 38},
        "C10": {"row_height": 32, "lines": 2, "left": 3},
        "D10": {"row_height": 32, "lines": 2, "left": 24},
        "E10": {"row_height": 32, "lines": 2, "left": 9},
        "F10": {"row_height": 32, "lines": 2, "left": 18},
        "G10": {"row_height": 32, "lines": 2, "left": 15},
        "H10": {"row_height": 32, "lines": 2, "left": 12},
        "I10": {"row_height": 32, "lines": 2, "left": 21},
        "J10": {"row_height": 32, "lines": 2, "left": 6},
        "L10": {"row_height": 32, "lines": 2, "left": 3},
        "M10": {"row_height": 32, "lines": 2, "left": 54},
        "N10": {"row_height": 32, "lines": 2, "left": 9},
        "O10": {"row_height": 32, "lines": 2, "left": 48},
        "P10": {"row_height": 32, "lines": 2, "left": 15},
        "Q10": {"row_height": 32, "lines": 2, "left": 42},
        "R10": {"row_height": 32, "lines": 2, "left": 21},
        "S10": {"row_height": 32, "lines": 2, "left": 36},
        "C13": {"row_height": 96, "lines": 6, "left": 3},
        "C15": {"row_height": 96, "lines": 6, "left": 3},
        "C17": {"row_height": 96, "lines": 6, "left": 3},
        "C19": {"row_height": 96, "lines": 6, "left": 3},
        "C21": {"row_height": 96, "lines": 6, "left": 3},
    },
    "faces": {
        "C4": {"baselines": (14,)},
        "E4": {"baselines": (18,)},
        "G4": {"baselines": (12,)},
        "I4": {"baselines": (12,)},
        "K4": {"baselines": (12,)},
        "M4": {"baselines": (12,)},
        "C6": {"baselines": (35,)},
        "E6": {"baselines": (36,)},
        "G6": {"baselines": (35,)},
        "I6": {"baselines": (35,)},
        "K6": {"baselines": (34,)},
        "M6": {"baselines": (34,)},
        "C8": {"baselines": (55,)},
        "E8": {"baselines": (55,)},
        "G8": {"baselines": (55,)},
        "I8": {"baselines": (55,)},
        "K8": {"baselines": (55,)},
        "M8": {"baselines": (55,)},
        "C10": {"baselines": (14, 31, 48)},
        "E10": {"baselines": (18, 41, 64)},
        "G10": {"baselines": (12, 26, 40)},
        "I10": {"baselines": (12, 26, 40)},
        "K10": {"baselines": (12, 27, 42)},
        "M10": {"baselines": (12, 27, 42)},
        "C12": {"baselines": (33, 50, 67)},
        "E12": {"baselines": (28, 51, 74)},
        "G12": {"baselines": (36, 50, 64)},
        "I12": {"baselines": (36, 50, 64)},
        "K12": {"baselines": (34, 49, 64)},
        "M12": {"baselines": (34, 49, 64)},
        "C14": {"baselines": (49, 66, 83)},
        "E14": {"baselines": (39, 62, 85)},
        "G14": {"baselines": (54, 68, 82)},
        "I14": {"baselines": (54, 68, 82)},
        "K14": {"baselines": (53, 68, 83)},
        "M14": {"baselines": (53, 68, 83)},
    },
    "overflow": {
        "C4": {"row_height": 15, "baselines": (12, 27, 42)},
        "E4": {"row_height": 15, "baselines": (12, 27, 42)},
        "G4": {"row_height": 15, "baselines": (12, 27, 42)},
        "I4": {"row_height": 15, "baselines": (12, 27)},
        "K4": {"row_height": 15, "baselines": (12, 27)},
        "C5": {"row_height": 16, "baselines": (5, 20, 35)},
        "E5": {"row_height": 16, "baselines": (-2, 13, 28)},
        "G5": {"row_height": 16, "baselines": (12, 27, 42)},
        "I5": {"row_height": 16, "baselines": (5, 20)},
        "K5": {"row_height": 16, "baselines": (-2, 13)},
        "C6": {"row_height": 17, "baselines": (6, 21, 36)},
        "E6": {"row_height": 17, "baselines": (-1, 14, 29)},
        "G6": {"row_height": 17, "baselines": (12, 27, 42)},
        "I6": {"row_height": 17, "baselines": (6, 21)},
        "K6": {"row_height": 17, "baselines": (-1, 14)},
        "C7": {"row_height": 18, "baselines": (6, 21, 36)},
        "E7": {"row_height": 18, "baselines": (0, 15, 30)},
        "G7": {"row_height": 18, "baselines": (12, 27, 42)},
        "I7": {"row_height": 18, "baselines": (6, 21)},
        "K7": {"row_height": 18, "baselines": (0, 15)},
        "C8": {"row_height": 19, "baselines": (7, 22, 37)},
        "E8": {"row_height": 19, "baselines": (1, 16, 31)},
        "G8": {"row_height": 19, "baselines": (12, 27, 42)},
        "I8": {"row_height": 19, "baselines": (7, 22)},
        "K8": {"row_height": 19, "baselines": (1, 16)},
        "C9": {"row_height": 20, "baselines": (7, 22, 37)},
        "E9": {"row_height": 20, "baselines": (2, 17, 32)},
        "G9": {"row_height": 20, "baselines": (12, 27, 42)},
        "I9": {"row_height": 20, "baselines": (7, 22)},
        "K9": {"row_height": 20, "baselines": (2, 17)},
        "C10": {"row_height": 21, "baselines": (8, 23, 38)},
        "E10": {"row_height": 21, "baselines": (3, 18, 33)},
        "G10": {"row_height": 21, "baselines": (12, 27, 42)},
        "I10": {"row_height": 21, "baselines": (8, 23)},
        "K10": {"row_height": 21, "baselines": (3, 18)},
        "C11": {"row_height": 22, "baselines": (8, 23, 38)},
        "E11": {"row_height": 22, "baselines": (4, 19, 34)},
        "G11": {"row_height": 22, "baselines": (12, 27, 42)},
        "I11": {"row_height": 22, "baselines": (8, 23)},
        "K11": {"row_height": 22, "baselines": (4, 19)},
        "C12": {"row_height": 23, "baselines": (9, 24, 39)},
        "E12": {"row_height": 23, "baselines": (5, 20, 35)},
        "G12": {"row_height": 23, "baselines": (12, 27, 42)},
        "I12": {"row_height": 23, "baselines": (9, 24)},
        "K12": {"row_height": 23, "baselines": (5, 20)},
        "C13": {"row_height": 24, "baselines": (9, 24, 39)},
        "E13": {"row_height": 24, "baselines": (6, 21, 36)},
        "G13": {"row_height": 24, "baselines": (12, 27, 42)},
        "I13": {"row_height": 24, "baselines": (9, 24)},
        "K13": {"row_height": 24, "baselines": (6, 21)},
        "C14": {"row_height": 25, "baselines": (10, 25, 40)},
        "E14": {"row_height": 25, "baselines": (7, 22, 37)},
        "G14": {"row_height": 25, "baselines": (12, 27, 42)},
        "I14": {"row_height": 25, "baselines": (10, 25)},
        "K14": {"row_height": 25, "baselines": (7, 22)},
        "C15": {"row_height": 26, "baselines": (10, 25, 40)},
        "E15": {"row_height": 26, "baselines": (8, 23, 38)},
        "G15": {"row_height": 26, "baselines": (12, 27, 42)},
        "I15": {"row_height": 26, "baselines": (10, 25)},
        "K15": {"row_height": 26, "baselines": (8, 23)},
        "C16": {"row_height": 27, "baselines": (11, 26, 41)},
        "E16": {"row_height": 27, "baselines": (9, 24, 39)},
        "G16": {"row_height": 27, "baselines": (12, 27, 42)},
        "I16": {"row_height": 27, "baselines": (11, 26)},
        "K16": {"row_height": 27, "baselines": (9, 24)},
        "C17": {"row_height": 28, "baselines": (11, 26, 41)},
        "E17": {"row_height": 28, "baselines": (10, 25, 40)},
        "G17": {"row_height": 28, "baselines": (12, 27, 42)},
        "I17": {"row_height": 28, "baselines": (11, 26)},
        "K17": {"row_height": 28, "baselines": (10, 25)},
        "C18": {"row_height": 29, "baselines": (12, 27, 42)},
        "E18": {"row_height": 29, "baselines": (11, 26, 41)},
        "G18": {"row_height": 29, "baselines": (12, 27, 42)},
        "I18": {"row_height": 29, "baselines": (12, 27)},
        "K18": {"row_height": 29, "baselines": (11, 26)},
        "C19": {"row_height": 30, "baselines": (12, 27, 42)},
        "E19": {"row_height": 30, "baselines": (12, 27, 42)},
        "G19": {"row_height": 30, "baselines": (12, 27, 42)},
        "I19": {"row_height": 30, "baselines": (12, 27)},
        "K19": {"row_height": 30, "baselines": (12, 27)},
        "C20": {"row_height": 31, "baselines": (5, 20, 35)},
        "E20": {"row_height": 31, "baselines": (-2, 13, 28)},
        "G20": {"row_height": 31, "baselines": (12, 27, 42)},
        "I20": {"row_height": 31, "baselines": (12, 27)},
        "K20": {"row_height": 31, "baselines": (13, 28)},
        "C21": {"row_height": 32, "baselines": (6, 21, 36)},
        "E21": {"row_height": 32, "baselines": (-1, 14, 29)},
        "G21": {"row_height": 32, "baselines": (12, 27, 42)},
        "I21": {"row_height": 32, "baselines": (13, 28)},
        "K21": {"row_height": 32, "baselines": (14, 29)},
        "C23": {"row_height": 15, "baselines": (11,)},
        "E23": {"row_height": 15, "baselines": (13,)},
        "G23": {"row_height": 15, "baselines": (15,)},
        "C25": {"row_height": 15, "baselines": (10,)},
        "E25": {"row_height": 15, "baselines": (16,)},
        "G25": {"row_height": 15, "baselines": (21,)},
        "C27": {"row_height": 20, "baselines": (12, 27, 42)},
        "E27": {"row_height": 20, "baselines": (12, 27, 42)},
    },
}
# fmt: on
"""What sheetpeek draws for each case, by group and cell, as far as the comparison needs it (zoom
100, points): the row's height, the cell's text, the number of lines and their baselines below
the row's top, the first glyph's advance from the cell's left edge (``left``), the drawn size
over the font's in sevenths (``scale``), and each rule (kind, lines, offset from its run's
baseline, top below the row's top, start from the cell's left edge, end from its right edge).
``tests/test_suite_text_style_more.py`` derives it from sheetpeek's layout (``drawn``)."""


def _font(size: float, name: str = "Calibri", **options: Any) -> Font:
    return Font(name=name, size=size, **options)


def _vert(ws: Any) -> list[Case]:
    found = []
    row = 4
    for size in SIZES:
        for position in (None, "superscript", "subscript"):
            ws[f"C{row}"] = "HHH"
            ws[f"C{row}"].font = _font(size, vertAlign=position)
            name = f"{position or 'plain'} {size}"
            found.append(Case("vert", name, f"C{row}", ("ink", "baselines")))
            row += 1
    for row, size in zip((17, 18, 19), BOTH_SIZES, strict=True):
        for column, position in (("C", "superscript"), ("E", "subscript")):
            ws[f"{column}{row}"] = "HHH"
            ws[f"{column}{row}"].font = _font(size, vertAlign=position)
            name = f"both {size} {position}"
            found.append(Case("vert", name, f"{column}{row}", ("ink", "baselines")))
    row = 21
    for size, digits in ((11, range(8, 13)), (22, range(4, 9))):
        for position in ("superscript", "subscript", None):
            for column, n in zip("CEGIK", digits, strict=True):
                cell = ws[f"{column}{row}"]
                cell.value = int("1234567890123"[:n])
                cell.number_format = "0"
                cell.font = _font(size, vertAlign=position)
                name = f"number {n} digits {size} {position or 'plain'}"
                found.append(Case("vert", name, f"{column}{row}", ("ink",)))
            row += 1
    return found


def _rules(ws: Any) -> list[Case]:
    for column in "CEGIK":
        ws.column_dimensions[column].width = 14
    found = []
    for row, size in zip((4, 6, 8, 10), RULE_SIZES, strict=True):
        for column, kind in zip("CEGIK", RULE_KINDS, strict=True):
            ws[f"{column}{row}"] = "HHH"
            if kind == "strike":
                ws[f"{column}{row}"].font = _font(size, strike=True)
            else:
                ws[f"{column}{row}"].font = _font(size, underline=kind)
            found.append(Case("rules", f"{kind} {size}", f"{column}{row}", ("rules", "ink")))
    for first, last, kind in (
        ("C12", "E12", "singleAccounting"),
        ("G12", "I12", "doubleAccounting"),
    ):
        ws[first] = "HH"
        ws[first].font = _font(11, underline=kind)
        ws.merge_cells(f"{first}:{last}")
        found.append(Case("rules", f"{kind} merge", first, ("rules", "ink"), last))
    for column, kind in (("C", "singleAccounting"), ("E", "doubleAccounting")):
        ws[f"{column}14"] = THREE
        ws[f"{column}14"].font = _font(11, underline=kind)
        ws[f"{column}14"].alignment = Alignment(wrap_text=True)
        found.append(Case("rules", f"{kind} wrapped", f"{column}14", ("rules", "lines", "ink")))
        ws[f"{column}16"] = "aa bb cc"
        ws[f"{column}16"].font = _font(11, underline=kind)
        ws[f"{column}16"].alignment = Alignment(horizontal="distributed")
        found.append(Case("rules", f"{kind} distributed", f"{column}16", ("rules", "ink")))
    for row, height in ((18, 15), (20, 20), (22, 30)):
        ws.row_dimensions[row].height = height
        ws[f"C{row}"] = ws[f"E{row}"] = "HHH"
        ws[f"C{row}"].font = _font(11, underline="singleAccounting")
        found.append(Case("rules", f"singleAccounting fixed {height}", f"C{row}", ("rules", "ink")))
        found.append(Case("rules", f"plain fixed {height}", f"E{row}", ("ink", "baselines")))
    return found


def _steps(ws: Any) -> list[Case]:
    for column, width in zip("KMOQ", STEP_WIDTHS, strict=True):
        ws.column_dimensions[column].width = width
    found = []
    for row in range(4, 23):
        for (name, size), shrunk, reference in zip(STEP_FONTS, "CEGI", "KMOQ", strict=True):
            ws[f"{shrunk}{row}"] = ws[f"{reference}{row}"] = "H" * row
            ws[f"{shrunk}{row}"].alignment = Alignment(horizontal="left", shrink_to_fit=True)
            ws[f"{reference}{row}"].alignment = Alignment(horizontal="left")
            for column in (shrunk, reference):
                ws[f"{column}{row}"].font = _font(size, name)
            found.append(
                Case(
                    "steps",
                    f"{name} {size} scale {row}",
                    f"{shrunk}{row}",
                    ("scale", "ink"),
                    None,
                    f"{reference}{row}",
                )
            )
    return found


def _shrink(ws: Any, cell: str, text: str, size: float = 11, **options: Any) -> None:
    ws[cell] = text
    ws[cell].font = _font(size)
    ws[cell].alignment = Alignment(horizontal="left", shrink_to_fit=True, **options)


def _shrunk(ws: Any) -> list[Case]:
    found = []
    columns = "CDEFGHIJKLMN"
    for row, height in ((4, 15), (6, 30), (8, 60)):
        ws.row_dimensions[row].height = height
        k = 0
        for vertical in ("top", "center", "bottom"):
            for step, text in SHRUNK.items():
                cell = f"{columns[k]}{row}"
                _shrink(ws, cell, text, vertical=vertical)
                name = f"{vertical} {height} shrunk {step}/7"
                found.append(Case("shrunk", name, cell, ("ink", "baselines")))
                k += 1
    four = SHRUNK[4]
    for row, height, label in ((10, 26, "26"), (12, 26, "26"), (14, None, "fitted")):
        if height is not None:
            ws.row_dimensions[row].height = height
        _shrink(ws, f"C{row}", four)
        ws[f"E{row}"] = "Hg"
        found.append(
            Case("shrunk", f"{label} shrunk 4/7, row {row}", f"C{row}", ("ink", "baselines"))
        )
        found.append(
            Case("shrunk", f"{label} plain 11, row {row}", f"E{row}", ("ink", "baselines"))
        )
        if row != 10:
            ws[f"G{row}"] = "Hg"
            ws[f"G{row}"].font = _font(20)
            found.append(
                Case("shrunk", f"{label} plain 20, row {row}", f"G{row}", ("ink", "baselines"))
            )
    _shrink(ws, "C16", four)
    _shrink(ws, "E16", "H" * 11 + "g", 20)
    found.append(
        Case("shrunk", "fitted shrunk 4/7 beside 20 pt shrunk", "C16", ("ink", "baselines"))
    )
    found.append(Case("shrunk", "fitted 20 pt shrunk", "E16", ("ink", "baselines")))
    ws.column_dimensions["O"].width = 2
    for cell, label in (("C18", "default column"), ("O18", "2-character column")):
        _shrink(ws, cell, "H" * 60)
        found.append(Case("shrunk", f"nothing fits, {label}", cell, ("ink",)))
    for row, height in ((20, 15), (22, 20)):
        ws.row_dimensions[row].height = height
        for column, vertical in (("C", "center"), ("E", "bottom"), ("G", "top")):
            ws[f"{column}{row}"] = THREE
            ws[f"{column}{row}"].alignment = Alignment(wrap_text=True, vertical=vertical)
            name = f"overfull {vertical} {height}"
            found.append(Case("shrunk", name, f"{column}{row}", ("lines", "ink")))
    return found


def _wrap(ws: Any) -> list[Case]:
    for column in "LMNOPQRS":
        ws.column_dimensions[column].width = 20
    found = []
    pairs = zip(INDENT_NARROW.items(), INDENT_WIDE.items(), strict=True)
    for k, (narrow, wide) in enumerate(pairs):
        row = 4 + k
        for block, (width, text) in (("CDEFGHIJ", narrow), ("LMNOPQRS", wide)):
            for j, column in enumerate(block):
                indent, horizontal = j // 2, ("left", "right")[j % 2]
                ws[f"{column}{row}"] = text
                ws[f"{column}{row}"].alignment = Alignment(
                    horizontal=horizontal, indent=indent, wrap_text=True
                )
                name = f"wrap {width} {horizontal} indent {indent}"
                found.append(Case("wrap", name, f"{column}{row}", ("lines", "ink")))
    for row, options, label in (
        (13, {"horizontal": "justify", "wrap_text": True}, "justify wrapped"),
        (15, {"horizontal": "justify"}, "justify unwrapped"),
        (17, {"vertical": "justify"}, "vertical justify"),
        (19, {"vertical": "distributed"}, "vertical distributed"),
        (21, {"wrap_text": True}, "wrapped"),
    ):
        ws[f"C{row}"] = text_style.SPILLED
        ws[f"C{row}"].alignment = Alignment(**options)
        found.append(Case("wrap", f"fitting {label}", f"C{row}", ("lines", "ink")))
    return found


def _faces(ws: Any) -> list[Case]:
    found = []
    for row, height, text, vertical in (
        (4, 60, HXG, "top"),
        (6, 60, HXG, "center"),
        (8, 60, HXG, "bottom"),
        (10, 90, "\n".join([HXG] * 3), "top"),
        (12, 90, "\n".join([HXG] * 3), "center"),
        (14, 90, "\n".join([HXG] * 3), "bottom"),
    ):
        ws.row_dimensions[row].height = height
        lines = "one line" if text == HXG else "three lines"
        for column, (name, size) in zip("CEGIKM", FACES, strict=True):
            ws[f"{column}{row}"] = text
            ws[f"{column}{row}"].font = _font(size, name)
            ws[f"{column}{row}"].alignment = Alignment(vertical=vertical, wrap_text=text != HXG)
            label = f"{name} {size} {vertical} {lines}"
            found.append(Case("faces", label, f"{column}{row}", ("lines", "ink", "baselines")))
    return found


def _overflow(ws: Any) -> list[Case]:
    found = []
    blocks = (
        ("C", "Haa\nHbb\nHcc", "center"), ("E", "Haa\nHbb\nHcc", "bottom"),
        ("G", "Haa\nHbb\nHcc", "top"), ("I", "Haa\nHbb", "center"), ("K", "Haa\nHbb", "bottom"),
    )  # fmt: skip
    for k, height in enumerate(range(15, 33)):
        row = 4 + k
        ws.row_dimensions[row].height = height
        for column, text, vertical in blocks:
            ws[f"{column}{row}"] = text
            ws[f"{column}{row}"].alignment = Alignment(vertical=vertical, wrap_text=True)
            lines = text.count("\n") + 1
            name = f"overflow {vertical} {lines} lines {height}"
            found.append(Case("overflow", name, f"{column}{row}", ("lines", "ink", "baselines")))
    for row, size in ((23, 14), (25, 20)):
        ws.row_dimensions[row].height = 15
        for column, vertical in (("C", "bottom"), ("E", "center"), ("G", "top")):
            ws[f"{column}{row}"] = "Hg"
            ws[f"{column}{row}"].font = _font(size)
            ws[f"{column}{row}"].alignment = Alignment(vertical=vertical)
            name = f"tall {size} {vertical} 15"
            found.append(Case("overflow", name, f"{column}{row}", ("lines", "ink", "baselines")))
    ws.row_dimensions[27].height = 20
    for column, vertical in (("C", "justify"), ("E", "distributed")):
        ws[f"{column}27"] = THREE
        ws[f"{column}27"].alignment = Alignment(vertical=vertical, wrap_text=True)
        name = f"spread {vertical} 3 lines 20"
        found.append(Case("overflow", name, f"{column}27", ("lines", "ink", "baselines")))
    return found


FILLS = {
    "vert": _vert, "rules": _rules, "steps": _steps, "shrunk": _shrunk, "wrap": _wrap,
    "faces": _faces, "overflow": _overflow,
}  # fmt: skip


def fill(ws: Any, group: str) -> list[Case]:
    """Lay one group out on a worksheet; its cases."""
    return FILLS[group](ws)


def cases(group: str | None = None) -> list[Case]:
    """Every case of ``group``, or of every group, in order."""
    chosen = GROUPS if group is None else (group,)
    return [case for g in chosen for case in fill(openpyxl.Workbook().active, g)]


def _boxes(found: list[Case]) -> list[str]:
    """The cells whose boxes a group's measuring needs."""
    cells = [c for case in found for c in (case.cell, case.last, case.reference) if c]
    return list(dict.fromkeys(cells))


def references(group: str) -> list[str]:
    """The cells whose geometry is read: row 1 for each column, column A for each row, and the
    cells whose text is read."""
    columns, rows = {}, {}
    for cell in _boxes(cases(group)):
        letters, row = coordinate_from_string(cell)
        columns[f"{letters}1"] = None
        rows[f"A{row}"] = None
    return [*columns, *rows, *TEXT_CELLS.get(group, ())]


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


def build(path: Path) -> list[Case]:
    """Write each group to its own worksheet of one workbook at ``path``; every case."""
    wb = openpyxl.Workbook()
    found = []
    for k, group in enumerate(GROUPS):
        ws = wb.active if k == 0 else wb.create_sheet()
        ws.title = group.capitalize()
        found.extend(fill(ws, group))
    wb.save(path)
    return found


def run(groups: tuple[str, ...] = GROUPS) -> None:
    """Capture each group's window at its zooms, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    for group in groups:
        capture(
            lambda ws, group=group: fill(ws, group),
            cells=references(group),
            out=BUILD_DIR / group,
            markers=MARKERS[group],
            zooms=ZOOMS[group],
            name=f"text_style_more_{group}",
        )
    write()


def left_bottom(cap: Any, box: tuple[int, int, int, int]) -> int | None:
    """The bottom (end exclusive, image pixels) of the ink in the left quarter of the ink in
    ``box``: the baseline of a line that starts with "H" and descends only further right."""
    from capture.screen.ink import ink

    found = ink(cap, box)
    if found is None:
        return None
    quarter = ink(cap, (found[0], box[1], found[0] + max(1, (found[2] - found[0]) // 4), box[3]))
    return None if quarter is None else quarter[3]


def measure(cap: Any, case: Case) -> dict[str, str] | None:
    """One case's measurements at one zoom (``text_style``'s and this suite's), or ``None`` when
    it is outside the window."""
    from capture.screen.ink import bands, content_box, ink, pt, rule_rows

    complete(cap, _boxes([case]))
    found = text_style.measure(cap, case)
    if found is None:
        return None
    top = cap.edge(case.cell, "top")
    box = content_box(cap, case.cell, case.last)
    if "baselines" in case.measure:
        lines = []
        for y, n in bands(cap, box):
            bottom = left_bottom(cap, (box[0], y, box[2], y + n))
            if bottom is not None:
                lines.append(pt(cap, bottom - top))
        found["baselines"] = " ".join(lines)
    glyph = ink(cap, box)
    if "rules" in case.measure and glyph is not None:
        rules = rule_rows(cap, box, max(4, int(0.6 * (glyph[2] - glyph[0]))))
        found["rule_tops"] = " ".join(pt(cap, y - top) for y, _, _, _ in rules)
    return found


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``text_style_more.tsv``."""
    from capture.screen.run import header, load

    records = []
    first: dict[str, str] = {}
    for group in GROUPS:
        out = BUILD_DIR / group
        if not (out / "formats.tsv").exists():
            print(f"{group}: not captured yet, left out", file=sys.stderr)
            continue
        first = first or header(out)
        for cap in load(out):
            for case in cases(group):
                found = measure(cap, case)
                if found is None:
                    continue
                row = coordinate_from_string(case.cell)[1]
                records.append(
                    {
                        "group": case.group,
                        "name": case.name,
                        "zoom": str(cap.zoom),
                        "cell": case.cell,
                        "row_height": cap.value(f"A{row}", "geometry.height") or "",
                        "text": cap.value(case.cell, "geometry.text") or "",
                        **found,
                    }
                )
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after", "zooms"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: text_style_more\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cases in {argv[1]}")
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
