# Capturing how Excel lays out a sheet

Expected layouts are captured from Excel (Microsoft 365) for Mac and for Windows, never written
from memory. A Mac capture builds a workbook with openpyxl, has Excel open it through AppleScript,
and reads back what Excel produced; a Windows capture is a job in the public capture-job format
(the Windows table below). The number formats themselves are xlformat's concern and are captured there; this
project captures layout: spill, alignment, merges, overflow, indent.

## Rules for every capture

- Workbooks are written inside Excel's sandbox container
  (`~/Library/Containers/com.microsoft.Excel/Data/Documents/sheetpeek/`). Opened from anywhere
  else, macOS asks for access to each file.
- Scripts are non-interactive. `display alerts` is switched off while a script runs and restored
  on every exit path, errors included. Left off, Excel would afterwards fail to show its dialogs.
- A script opens only the workbook it built, and closes it without saving unless saving is the
  point. It stops before opening anything when a workbook of the same name is already open in
  Excel: that workbook is not the script's (it may have been left open to compare), and
  Excel would hand the script that copy instead of the file just built. It never touches another open workbook, and never runs an application-wide command such
  as `calculate`.
- Two captures never run at once.
- No sweep of crafted or extreme input runs during a capture, and a capture starts only when
  Excel answers: a machine out of memory pauses Excel, which then looks like a hang.
- A case that hangs ends the run and is reported, not retried: a retry hides what made Excel
  hang, and a hang that comes from Excel's state comes back.
- Anything that feeds crafted or extreme input to the package runs in a child process under a memory
  watchdog that samples its size and kills it at 4 GB (`tests/watchdog.py`): a time limit alone
  does not bound memory, and `ulimit -v` is not a limit on macOS. A killed run, a missing result or
  an exit status of 137 is a finding, not noise to rerun past: the same input does the same again.
- The page setup a capture needs (orientation, scale, gridlines, margins) is written into the
  workbook by openpyxl, not set through AppleScript.

## Route A: the web page export

`capture/html_probe.py` with `capture/html_export.applescript`
(`uv run --group capture python -m capture.html_probe run`). Excel's "save as web page" writes a
single-file MIME web archive whose sheet part is HTML 4 with `mso-` styles. It records:
- column widths (`mso-width-alt` is the file width × 256);
- hidden rows and columns;
- merges as `colspan`, and right spill as `colspan` marked `mso-ignore:colspan`;
- each cell's displayed text, the same as AppleScript's `string value`.

It does not record left or centered spill. It treats a neighbor holding only a fill as stopping
a spill, which the screen shows is wrong. It cuts clipped text at a word boundary. Route A is
therefore used for structure only.

## Route B: the PDF export

`capture/layout.py` with `capture/pdf_export.applescript`
(`uv run --group capture python -m capture.layout run`), read with pdfminer.six (the `capture`
dependency group; never a dependency of the package). Facts the method relies on, measured with
`capture/pdf_probe.py`:

- With the workbook's page setup at 100 % scale, one PDF point is one sheet point. A default
  column in Calibri 11 is 53 pt, and a digit is 6 pt. A row 15 pt high in the file prints 14 pt
  apart.
- Glyphs sit on whole points. Left-aligned text starts 3 pt after the column's left gridline, and
  right-aligned text ends 2 pt before its right one.
- Excel prints each gridline as a 1 pt filled rectangle, **and leaves a vertical one out wherever
  text crosses it**. So the missing gridlines of a row show exactly which column boundaries its
  text spans: the most direct evidence of spill.
- The PDF holds the full text of every cell, and clips it with clip paths that text extraction
  does not report. Clipping is therefore read from the rendered page, not from the text runs.
- Each character carries its fill color, which gives the color of a `#` run. Cell fills are
  filled rectangles with their own color.
- A filled cell hides the gridlines on its edges, so a missing gridline beside a fill is not
  evidence of spill; the text runs and the boundaries beyond the fill are.
- A row with a custom height shows only the lines that fit, so a spacer row follows each
  multi-line case.

The capture writes `tests/captures/layout.tsv`, one record per sheet row, after a header that
records the Excel version, font, column widths, date and method:
- `missing`: the boundaries without a gridline, numbered between visible columns (1 is A|B);
- `runs`: each run of text, as `'text'@x0-x1#RRGGBB` in PDF points;
- `fills`: each filled area other than white, as `x0-x1#RRGGBB`;
- `baselines`: per run, its baseline in points below the top edge of the row's top gridline,
  from the glyph's text matrix.

A column of width 0 is written into the XML after openpyxl saves, because openpyxl drops a zero
width. A second run of the capture reproduced the earlier rows exactly.

## Row heights

`capture/row_height.py` with `capture/row_height.applescript`
(`uv run --group capture python -m capture.row_height run`) reads AppleScript's `row height` and
the worksheet's `standard height`, in points, from one workbook per case. Each has a measured
font as its Normal style, a number in A1, an empty cell element in row 3, an empty `<row r="5"/>`,
no element for rows 2 and 100, and `<sheetFormatPr>` written into the XML exactly as the case
gives it, or left out. It also reads back the font Excel applied to A1, which shows that Excel
did not substitute the theme font. The capture writes `tests/captures/row_height.tsv`; three runs
reproduced every common row. Facts from it (Excel 16.115, `tests/captures/row_height.tsv`):

- **The font height.** A row without `ht` takes the Normal font's height: Calibri 9 to 16 give
  12, 14, 15, 16, 19 and 21 pt; Aptos Narrow gives 13, 14, 15, 16, 19 and 22 pt.
- **`defaultRowHeight` counts only with `customHeight="1"`.** Without it, Excel ignores the
  stored value (15, 20 and 30 all give the font height, at Calibri 11 and 14). With it, every row
  without `ht` takes the stored value exactly, rows with cells included (12.75 stays 12.75).
- **A `<row>` element matters, not a value.** Rows with an element (a value, an empty cell, or
  none) and rows without one differ in three cases; otherwise all rows are equal:
  - without `<sheetFormatPr>`, a row with an element has the font height, and a row without one
    the `standard height`, 16 pt for every font: Excel for Mac's own default, from the
    application's standard font (Aptos Narrow 12 under Microsoft 365, ARCHITECTURE §4.7), which
    the capture does not vary;
  - with `thickTop` or `thickBottom`, a row without an element is 1 pt taller per attribute;
  - with `zeroHeight`, a row without an element is hidden (0 pt), and a row with one has the font
    height, even when `customHeight="1"` stores another.
- **A row is fitted to its cells when Excel opens the file** (A1 in another font, wrapped or
  merged, in a Calibri 11 workbook): a 14 pt cell makes its row 19 pt, bold or not, even when
  the cell is empty but styled; a 9 pt one leaves it at 15; wrapped text gives 16 pt per line
  (three lines 48, six lines 96); in a merge over A:C the 14 pt font still counts, the wrapped
  lines do not; a stored `ht` of 15 stays 15.

## Route C: the screen

The PDF is Excel's print layout; the screen is the truth, and the two differ in places (a row 15
pt high in the file prints 14 pt apart). `capture/screen/` captures Excel's window:

- **Requests** go through `capture/screen/excel.applescript`, one request a call (count, open,
  close, view, scroll, window, maximize, geometry, save, read, follow), so each call does one visible thing and can
  be inspected on its own: `scroll` sets the window's top row and returns its visible range,
  `window` reads the window's zoom, view, freeze and split state without setting anything, and
  `follow` follows a cell's first hyperlink and reads where it lands. The suites that read
  Excel's own answers without a window (`package`, `package_more`, `odd_values`, `row_height`,
  `row_height_more`, `theme_colors`, `theme_colors_more`) have scripts of their own.
  A request past its time limit is killed (`osascript` waiting on Excel ignores SIGTERM) and
  evidence is saved: a `sample` of Excel and whether it still answers.
- **Safeguards:** a screen suite starts only when Excel answers, holds no workbook and shows no
  window of its own, since any other window spoils a window capture (`layout` and the suites with
  scripts of their own, above, only refuse a workbook of their own name). It opens only the workbook it built, inside Excel's sandbox
  container, keeps the display awake, and on every exit path closes it without saving, turns
  display alerts back on and deletes it.
- **The window** is found through CoreGraphics by its title (`capture/screen/windows.js`) and
  captured with `screencapture -l` at the display's scale (2 px a point on a 2x display) once it
  has settled, since Excel may still be redrawing after a zoom change (`window.settle`): shots
  back to back (about 0.2 s each), compared as bytes (a still window gives identical files); the
  window is settled when its picture has held for 0.4 s, and a run ends when it has not after
  8 s. Each capture records how it settled (`settle.<zoom>`: shots, when the last change was
  seen, how long the picture held, the longest stillness a change interrupted, and `near` when
  that came within 0.1 s of 0.4). A run's first capture is shot again 1 s later and must be
  unchanged; with `SHEETPEEK_CAPTURE_AUDIT=1` every capture is watched for 2 s after settling,
  for a suite's first run and after an update of Excel or macOS. The shots compared are of the
  window less 16 pt at its right and 72 pt at its bottom (`screencapture -R`): macOS's overlay
  scroll bars, which a view's scroll shows and which fade about a second later, lie there
  (measured in the macOS VM: the window's only change after settling, a 14 px strip down its
  right edge and one above the sheet tabs). That keeps the headings and nearly every cell; the
  window is then captured whole.
- **The window's geometry:** where a dash or a tile starts follows Excel's window on the screen
  (a dash's phase moved with the window at zoom 200), so each machine captures in one geometry,
  recorded once outside the repository (`capture/screen/geometry.py`: `$SHEETPEEK_CAPTURE_GEOMETRY`,
  else `~/.sheetpeek-capture/geometry.json`, written by `python -m capture.screen geometry
  record`): the main screen's size in points and its backing scale, Excel's window maximized
  (bounds and window state) and its usable width and height. Before its first view a capture
  maximizes the window, reads the geometry, records it in its header (`geometry`) and refuses to
  run when it differs from the recorded one in any field, or when none is recorded.
- **Colors:** each capture is converted from the display's profile to sRGB (`sips --matchTo`),
  so pixel colors compare with the file's. From a wide-gamut display the conversion leaves
  saturated colors a few units off: on a Mac with a P3 display, `border_conflicts.tsv` reads
  `00FF00` as `03FF00`, `FF8000` as `FF8002` and `A00000` as `A00001` (32 of its 484 window
  rows, at most 3 units in one channel), where the same suite run in a macOS virtual machine,
  whose display is sRGB, reads every color exactly. A test comparing a captured color allows
  for it (within 24 of the expected color, and nearer it than the other candidate).
- **In a macOS virtual machine** the suites run as on the host. A suite's first run there is
  compared with its committed capture: `border_conflicts` (16.115, all four groups) gave the
  same data lines, but for the colors above in `colors` (28 rows) and `black_white` (4 rows);
  every drawn style, offset, pattern and distance the same.
- **Markers:** two cells (B2 and one below and right of the cases) are filled magenta; their pixel
  boxes, against Excel's own `left position`, `top`, `width` and `height` of the cells, give the
  origin and the scale per axis. A marker must be a cell's full height, which keeps out the
  floating Copilot button of Microsoft 365 (a 15 × 10 px magenta-ish disk, when it is on).
- **One marker, on a small screen:** where only B2 shows, its box gives the origin and the zoom's
  nominal scale the rest: screen pixels per point at zoom 100 (1 on the Mac, 96/72 on Windows)
  times the display's scale (recorded in the header as `display_scale`) and the zoom. Wherever
  both markers show, the nominal scale is checked against them (`nominal_off`). On the Mac, at
  zoom 100 and 200, it matched every saved capture exactly, and measuring every suite's saved
  captures from one marker gave the same results; at other zooms Excel rounds rows and columns
  to whole pixels (zoom 150: 3.008 across and 3.060 down against 3.0), so a one-marker grid holds
  only where a two-marker run's check shows the nominal scale holds.
- **Gridlines** are `#D4D4D4` on the Mac and `#E0E0E0` on Windows (`windows/borders`,
  `windows/fills`). A control that finds a gridline takes its run's platform's gray
  (`window.GRIDLINE`, from the header's `platform`, `mac` for runs before the field).
- **Reads:** for every cell measured, its display format (with conditional formatting and table
  styles), its own format, its own borders and its geometry, each property in its own `try`.
- **Headers** record Excel's version and build, the edition (stated in `EXCEL_EDITION`, which
  Excel's dictionary does not give), the update channel, macOS, the appearance, the zoom, the
  color conversion, the date, the suite, the platform, the display's scale and the window's
  geometry. The headers of
  `layout` and of the suites with scripts of their own record version, build, edition and
  channel too.

Units: lengths are reported in sheet points at the capture's zoom (image pixels divided by the
scale), with image pixels beside them where a suite keeps them. Lines, dashes, pattern tiles and
outline symbols are fixed in screen units, so they are reported in screen points.

`python -m capture.screen run SUITE ...` runs suites one at a time (one Excel, never two
captures at once); `read SUITE ...` measures saved captures again from `capture/build/<suite>/`
without Excel. Each suite writes `tests/captures/<suite>.tsv`, gzipped as `<suite>.tsv.gz` when
it is large (`tables`). Every suite is in the table under **Suites** below; among the screen
suites:
- the cell-style suites `borders`, `fills` and `text_style`;
- `rotation`: angled and stacked text (which borders slant, where general text is anchored,
  fitted heights over lengths and sizes; cases measured by row height alone are read without a
  capture);
- `filters`: filter and sort buttons;
- `tables`: table styles (each table cell's display format, read exactly, and the device pixels
  across each of its edges), one workbook and window per group; its groups `unknown` and `merge`
  may be refused by Excel and run on their own (`python -m capture.tables run unknown`).

Cases added to a suite after it was captured go where the captured ones stay as they were: a
group of their own (`text_style`'s `wrap` and `lone`, run with
`python -m capture.text_style run wrap lone`), or windows of their own (`rotation`'s second set,
`python -m capture.rotation run more`). `read` leaves out what is not captured yet.

## Saving a copy

`capture/excel_saved.py` with `capture/save_copy.applescript` has Excel open a workbook of
formulas written by openpyxl, which carries no cached values, and save a copy as `.xlsx`. The copy
holds Excel's own cached results and Excel's own XML. It is kept as `tests/files/excel_formulas.xlsx`.

## What Excel does with a doubtful package

`capture/package_probe.py` with `capture/package_open.applescript`
(`uv run --group capture python -m capture.package_probe run`) opens one small workbook per case,
holding 111 in A1, each in its own `osascript` under `gtimeout`; a stopped run has its alerts
restored by `capture/restore_alerts.applescript`. A control opens first and last, and an outcome
counts only between two that open. The only DTD ever given to Excel is a minimal
`<!DOCTYPE worksheet>`, with no entity and no external reference. With alerts off, error -50 means
Excel would refuse the file or offer to repair it. Facts from it (Excel 16.115, `tests/captures/package.tsv`): both controls open
and read 111; a `<!DOCTYPE>`, a bzip2 or LZMA worksheet part, a worksheet part named twice, and a
central-directory size that is too small, too large, or that of the part's XML with its CRC, with
spaces following, all give -50.

The same run holds the numbers no column, row or font can have, read back with `column width`,
`row height` and `font size`, for columns A and B and rows 1 and 2; `ARCHITECTURE.md` §5.5 has what
Excel makes of them. A case that hangs ends the run and nothing is written.

## Suites

| suite | rows | what it covers |
|---|---|---|
| `layout` | 200 | spill reach by alignment (digits 6 to 40), centered placement, glyph text, clipping, competing spills, hidden and width-0 columns, `#` color, numbers under `@`, indent, merged text placement and fill, `fill`, `justify`, `distributed` and wrapped cells, line breaks, fills under spill, merged and center-across text wider than its area, `_xHHHH_` escapes, line feeds and control characters |
| `theme_colors` | 96 | theme indices 0 to 11 with eight tints, as fill and font color, read with AppleScript per cell |
| `merge_width` | 150 | numbers in merges of one to four columns, one with a hidden column, at six widths, read with `string value` |
| `error_width` | 56 | seven error values at eight widths, read with `string value` |
| `row_height` | 55 | row heights of rows without `ht`, per Normal font and size (Calibri and Aptos Narrow at 9 to 16 pt), without `<sheetFormatPr>`, with openpyxl's, and with `customHeight`, other stored heights, `zeroHeight`, `thickTop` and `thickBottom`; rows fitted to a cell font, a styled empty cell, wrapped text and merges |
| `borders` | 352 | route C, the screen: the 13 border line styles on four sides: device rows, dash runs, darkest ink, at zoom 100 and 200 |
| `filters` | 30 | route C, the screen: filter and sort buttons on the sheet's filter and two tables: box, gaps, glyph (bitmap id), at zoom 100 and 200 |
| `fills` | 82 | route C, the screen: the 19 pattern types in two color pairs: tile, share, extent, gridlines between patterned cells, at zoom 100 and 200 |
| `text_style` | 234 | route C, the screen: vertical and horizontal alignment, distributed included, indent, reading order; shrink to fit; underlines, strike, super- and subscript; the width a wrapped line may take; lone distributed words |
| `tables` | 2,726 | route C, the screen: table styles, all 60 presets and the rules around them: each table cell's display format, read exactly, and the device pixels across each of its edges, at zoom 100; Excel's refusal of a merge inside a table |
| `theme_colors_more` | 76 | read from Excel: stored 16-bit tints on themes 0 to 9 against their round values, and tints of a third on black and on theme 3, as fill and font color |
| `auto_colors` | 37 | route C, the screen: automatic, system and missing colors (indexed 64, 65 and 81, `auto`, no element, an element without attributes) in solid and `lightGrid` fills, fonts and medium borders: interior colors and shares, gridlines, ink, edge runs and Excel's display format, at zoom 100; Excel's refusal of a border in indexed 81 |
| `headings_grid` | 174 | route C, the screen: the grid color by `colorId` (palette entries, 64, 65, past the palette, malformed) and the edges a fill covers; the heading bands by the last row's digits (2 to 7) at zoom 50, 75, 100, 150, 200 and 400 and by Normal font (Calibri 8 to 36, Aptos Narrow 11); the headings' colors in the dark and light appearances; Excel's refusals of `colorId` 65 and 99999 |
| `spill_neighbors` | 70 | route C, the screen: the neighbors that stop a text's spill (a number, a saved `""`, a space, `=0`, a zero the sheet hides) and numbers of 7 to 12 digits in Calibri 9, 11 bold, 14 and 20 and Arial 11 at two widths: Excel's text, ink extent and row height, at zoom 100 |
| `filters_more` | 45 | route C, the screen: filter buttons on a header merged across and down, an empty header, a hidden column and a hidden table header row, criteria not applied on opening, a sort beside the sheet's filter: box and glyph at zoom 100 and 200, row heights and the hidden column's width; Excel's refusals of a left-to-right sort in the sheet's filter and in a table |
| `extensions` | 29 | route C and saves: sparklines, validation, protected ranges and ignored errors under their exact sheet extension URIs and each in lower case: the sparkline's ink, error triangles at zoom 100, and where Excel puts each extension when it saves the file |
| `freeze_view` | 90 | route C, the screen: the Normal view at manual breaks and print-area edges, the freeze line at zoom 100 and 200 in the dark and light appearances with text spilling across it, page break preview's break marks, and the zoom Excel opens and draws crafted `zoomScale` values at (window state and visible range) |
| `links` | 18 | route C, the screen: linked cells in the Normal, Hyperlink and Followed Hyperlink styles, with and without a link, `display` unlike the value, empty cells and merges, `HYPERLINK()`: text, ink color, underline rules and Excel's display format, at zoom 100; where Excel goes when it follows each link (a location with and without a sheet, from its own sheet and another, an element beside `HYPERLINK()`) |
| `merges_spill` | 140 | route C, the screen: a merge's outline from its edge cells' sides, the anchor's sides alone, neighbors competing per segment, a merge's and a covered cell's diagonals, and spilled text across borders and fills: what each edge shows and the diagonals' coverage, at zoom 100 and 200 |
| `border_conflicts` | 500 | route C, the screen: every pair of the 13 border styles meeting, both orders and orientations; pairs of colors that darkness measures order differently, equal Rec. 601 luma, theme, tinted and indexed colors, black and white against six colors; sides across hidden rows and columns; `start` and `end` with and without `left` and `right` |
| `gradients` | 72 | route C, the screen: linear gradients at degrees around the circle in wide and square cells, path gradients, first stops past 0, duplicate stop positions, degrees out of range or not finite, gradients without stops: an 8 by 3 grid and profiles across and down each cell, at zoom 100 and 200 |
| `theme_fonts_palette` | 47 | route C, the screen, and reads: scheme fonts under a Courier New and Georgia theme with Hiragino script fonts, no theme part, an empty typeface; a workbook palette recolored at 3, 10, 12 and 17 in fills, fonts and number-format colors, short, long and non-color lists: Excel's display format, ink and fill colors, widths, at zoom 100; Excel's refusals among the crafted files |
| `odd_values` | 8 | read from Excel: `textRotation` 181, 254, 256 and 1000, and `numFmtId` 170 and 300 left undefined, between two controls: the orientation, number format and text Excel shows |
| `rotation` | 157 | route C, the screen: slanted borders and the text between them, anchoring, alignment, fitted heights, turned wrapping, upright widths at 9 to 20 pt |
| `package` | 68 | a control, a minimal `<!DOCTYPE>`, bzip2 and LZMA parts, a part named twice, three kinds of false size, widths, heights, font sizes and sheet defaults no sheet can have, a control |
| `package_more` | 26 | read from Excel, between two controls: widths of 1e100, 1e200 and 1e300, `defaultColWidth` and `baseColWidth` at 255, 256, 299 and 300, font sizes 409, 409.5, 410, -2 and -10, and Normal fonts Arial 10 and Calibri 14 with widths of -1, 1e6 and 1e30: whether Excel opens each, and the widths, heights and font it shows |
| `color_kinds` | 7 | read from Excel: a fill, font and border color stating two kinds (theme with rgb, rgb with indexed, theme with indexed, auto with rgb) beside one of each kind alone: the color Excel shows |
| `runs` | 31 | route C, the screen: rich-text runs in shared strings: which font and color each word is drawn in (a run without properties, one of only bold, missing name and size under an Arial 14 Normal font, section colors, a text section adding characters, a `<t>` beside runs), widths added per run, rows and baselines of mixed sizes, wrapped and not, superscript and subscript runs, at zoom 100 |
| `tab_strip` | 106 | route C, the screen, dark and light: the tab strip's items left to right (tabs, separators, scroll arrows, new-sheet button) with each tab's size, fill and text color: visible sheets in file order, every kind of tab color inactive and active, `auto`, and tab colors carrying two kinds at once |
| `outline` | 432 | route C, the screen, dark and light: the outline gutter's bars, ticks, dots, minus and plus buttons and level buttons, per row and column, with summaries below and right and above and left, rows of 15, 30 and 60 pt, at zoom 100 and 200; `showOutlineSymbols` on the view and on `<outlinePr>`, declared depths, and a flag of `yes` |
| `text_style_more` | 547 | route C, the screen, zoom 100 and 200: superscript and subscript at 9 to 36 pt, rules at 9 to 36 pt and in merges, wrapped and distributed text, accounting baselines in fixed rows, shrink steps and baselines at other sizes and faces, overfull blocks, an indent on wrapped left and right text, justify and the vertical spreads in fitting, baselines of sizes and faces xlformat has not measured; wrapped lines that overflow rows of 15 to 32 pt, a line taller than its row, and vertical justify and distributed overflowing |
| `text_steps` | 256 | route C, the screen, zoom 100: "HHHH" at indent 0 to 5, left and right aligned, under Normal fonts Calibri 9, 11, 12, 16 and 24 and Aptos Narrow 11 and 12, and in Calibri 16 and Aptos Narrow 12 cells under Calibri 11; "H" repeated 4 to 30 times with shrink to fit in a default column beside its unshrunk reference, and unshrunk text at 1 to 11 pt in half points; under Normal fonts Calibri 16 and Aptos Narrow 11, "H" repeated 4 to 51 times shrunk beside two unshrunk references |
| `row_bottoms` | 297 | route C, the screen, zoom 100: "HHH" in Calibri 11, bottom aligned in rows fixed at 32 pt, beside a Calibri 24 cell of each kind (a value, an empty styled cell, a hidden column, 45 and 90 degrees, stacked, a merge across rows, superscript, subscript, a rich-text run) and under a row's and a column's own Calibri 24 style; the same text merged across two rows beside a neighbor in its first or last row; three lines vertically distributed and justified in 70 pt rows, beside the neighbor and alone; each with the baseline each candidate rule gives, written before the run; and a Calibri 11 probe beside superscript or subscript "HHH" at every whole and half size from 6 to 40 in 64 pt rows (the shared bottom and the glyph's baseline) |
| `diagonals` | 238 | route C, the screen, zoom 100 and 200: diagonal borders of each style up, down and both, in small and large cells, with text under them, without a style or a direction: where each line lies, its thickness, its dashes along it and what shows where it crosses the text |
| `rotation_more` | 43 | route C, the screen: angled text between slanted sides at other alignments (top, negative angles, vertical spreads), fill, justify and distributed at angles, merges, stacked text in each alignment, an angled accounting underline, the sides of slanted cells and their neighbors: ink extents, sides and rows |
| `rotation_placement` | 73 | route C, the screen, zoom 100: fill, justify, distributed and left at 30, 60, 85, -15, -75 and -90 degrees; top alignment without sides in 90 and 30 pt rows and between four sides at more angles and rows; "Abc" stacked left, right and center at 9 to 36 pt and "WiW" at 11 and 20, in a 20-character column |
| `row_height_more` | 72 | read from Excel: row heights under Normal fonts Arial, Times New Roman, Courier New and Georgia at 11 and 14, and Calibri and Aptos Narrow at 13, 15, 17 and 22 (a listed row, a row without `<row>`, three wrapped lines); under Calibri 11, a number in each of those fonts and three wrapped lines in Calibri 11, 14, 20, Aptos Narrow 11 and the four faces |
| `row_height_fitted` | 1016 | read from Excel, under Calibri 11: the rows Excel fits to text at 90 degrees at every size from 9 to 36 (three texts) and to runs of one glyph, 1 to 40, at 11, 14 and 20 and short runs at 16 to 36, and the three texts at 88, 89 and -90 at 20; and to "HHH" in Calibri at every half size from 6 to 40 and to 72, superscript, subscript, both side by side and accounting at 6 to 40, and Arial, Times New Roman, Courier New and Aptos Narrow at 6 to 40; superscript, subscript and both at every half size from 6.5 to 39.5 (group `scripts`) |
| `merge_rows` | 22 | read from Excel, under Calibri 11: the rows Excel fits around merges across rows, wrapped text (2, 3 and 6 lines over 2 and 3 rows, 14 pt, a 14 pt neighbor, a merge two columns wide) and one "H" in 14 and 20 pt over 2 and 3 rows, with one wrapped cell and a horizontal merge as controls; predictions in the module before the run, of which sheetpeek's (which then counted a merge's font in its first row) held for 18 rows: a merge's font fits no row either |
| `wrap_rows` | 96 | read from Excel, under Calibri 11: the rows Excel fits to text of one to six lines, wrapped, justified, distributed and distributed with wrap off, as a word a line or broken by line feeds, in Calibri 11, Calibri 16 and Arial 11: the alignment does not change the row, which is the lines' (k lines of Calibri 11 take 16k pt) |
| `rules_below` | 63 | route C, the screen, zoom 100: "HHH" under double accounting (11, 20, 36 pt), single accounting and single (36 pt) underlines and plain controls in rows fixed at 12, 15 and 20 pt, bottom, center and top aligned, each over an empty 60 pt row: whether a line past its row is drawn on into the row below (`below`, the ink bands under the cell's bottom edge) |
| `font_cut` | 50 | route C, the screen, zoom 100: Arial 24 "Hxg" probes, each a red of its own, at font-list entries 990 to 1,009 and as rich-text runs at font-table entries 995 to 1,004 (after the list, in cell order), behind filler fonts on a second sheet, under Calibri 11 Normal and theme fonts; and at list entries 995 to 1,004 under a theme of Arial body and Aptos Narrow heading fonts with a Times New Roman 10 or 16 Normal font: the fitted row, the font Excel reports, the ink's height, color and glyph starts, and Excel's standard font (`standard_font`, recorded with the theme groups) |

**Captures from Excel for Windows** (`tests/captures/windows/`), each read from a capture job in
[sheetmill's job format](https://github.com/sheetmill/sheetmill/blob/3adb865500f4d2223e5c1a758b466e081d3fdc79/docs/capture-jobs.md) by `capture/job.py` (`python -m capture.job write GROUP JOB_ID OUT`
makes the job; `read GROUP JOB_DIR RESULT_DIR` writes the suite files from its result). Each file
has its Mac suite's schema, under a header of sheetpeek's own lines (`suite`, `platform`,
`excel_build`, `job`, `job_id`, `captured`, and `zooms` and `gridline` where the suite has them)
and the result's header whole as one line (`job_header`), whose fields that page defines; the
job's manifest lies beside it as `<group>.job.json`. Excel 16.0.20527.20196 (Microsoft 365), Windows at 96 dpi, en-US.

| file | rows | what |
|---|---|---|
| `windows/cells` | | the job of the group `cells`: `capture/cell_text.py`'s workbook opened, and each cell's displayed text read (`read` of `text`) |
| `windows/font_cut` | 50 | as the Mac's `font_cut`, all four groups, by the job of the same name (themes rewritten before the copy, the displayed font and its color read; a rich-text cell's font, which Windows returns as nothing, an empty field) |
| `windows/font_aptos` | 11 | `font_cut`'s group `faces_aptos` by the job `font_aptos`: the references and probes of `windows/font_faces` under a theme of Aptos Narrow body and Aptos Display heading fonts with an Aptos Narrow 11 Normal font |
| `windows/merge_rows` | 22 | the suite `merge_rows`'s workbook by the job of the same name: each case row's height and column A's applied font, read from Excel (`read` of `row_height`, `read_format` of `font.name` and `font.size`), in the Mac suite's columns |
| `windows/row_height_fitted` | 483 | `row_height_fitted`'s groups `styled` and `scripts` by the job `fitted_rows`: each case row's height and column A's applied font |
| `windows/wrap_rows` | 96 | the suite `wrap_rows`'s workbook by the job `fitted_rows`: each case row's height and column A's applied font |
| `windows/fitted_rows` | | the job of the group `fitted_rows`: the workbooks of `row_height_fitted`'s groups `styled` and `scripts` and of `wrap_rows`, each opened, and each case row's height and column A's applied font read (`read` of `row_height`, `read_format` of `font.name` and `font.size`) |
| `windows/font_faces` | 22 | `font_cut`'s groups `faces_calibri` and `faces_times` by the job `font_faces`: "Hxg" in Arial, Aptos Narrow, Segoe UI, Tahoma and Calibri 11 at font-list entries 1 to 5 beside Arial 24 probes at entries 997 to 1,002, under a Calibri 11 or Times New Roman 10 Normal font and openpyxl's theme |
| `windows/line_budget` | 136 | `text_style`'s group `line_budget` alone, captured at zoom 100 by the job of the same name: two words the cell less 6, 5 and 4 px; an "i"'s dots on Windows are a line band of their own (1 or 2 px), left out when lines are counted |
| `windows/merge_width` | 150 | as the Mac's `merge_width`: numbers alone and merged across two to four columns, one hidden, at six widths |
| `windows/error_width` | 56 | as the Mac's `error_width`: the seven error values at eight widths |
| `windows/borders_fills` | | the job of the group `borders_fills`: the windows of `borders` and `fills`, each opened, its cells' geometry and displayed fill read, and Excel's window captured at zoom 100 and 200 (`capture_window`) |
| `windows/borders` | 176 | as the Mac's `borders`, measured on the job's pictures at their device pixels a point; the header's `gridline` is the gray measured on the control cell |
| `windows/fills` | 82 | as the Mac's `fills`; tiles in device pixels (`tile_unit` `device_px`), `display_pattern` Windows' `XlPattern` number |
| `windows/text_steps` | 256 | as the Mac's `text_steps`: a workbook a group under its Normal font, captured at zoom 100 by the job of the same name; the gray measured on D1 |
| `windows/shrink_fonts` | 65 | `text_steps`'s group `shrink_fonts` alone, captured at zoom 100 by the job of the same name: shrink to fit under five cell fonts, with each font's unshrunk references |
| `windows/text_style` | 226 | as the Mac's `text_style`: a workbook a group at its zooms, captured by the job of the same name with the displayed font and alignment read (`underline`, `h_align` and `v_align` Windows' numbers); cells outside the window at zoom 200 left out |
