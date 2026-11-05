# Capturing how Excel lays out a sheet

Expected layouts are captured from Excel 16 for Mac, never written from memory. Each capture
builds a workbook with openpyxl, has Excel open it through AppleScript, and reads back what Excel
produced. The number formats themselves are xlformat's concern and are captured there; this
project captures layout: spill, alignment, merges, overflow, indent.

## Rules for every capture

- Workbooks are written inside Excel's sandbox container
  (`~/Library/Containers/com.microsoft.Excel/Data/Documents/sheetpeek/`). Opened from anywhere
  else, macOS asks for access to each file.
- Scripts are non-interactive. `display alerts` is switched off while a script runs and restored
  on every exit path, errors included. Left off, Excel would afterwards fail to show its dialogs.
- A script opens only the workbook it built, and closes it without saving unless saving is the
  point. It stops before opening anything when a workbook of the same name is already open in
  Excel: that one is someone's (left open to compare, say), and Excel would hand the script that
  copy instead of the file just built. It never touches another open workbook, and never runs an application-wide command such
  as `calculate`.
- Two captures never run at once.
- No sweep of crafted or extreme input runs during a capture. Before a capture, check that Excel
  answers and that no sweep is still running: a machine out of memory pauses Excel, which then looks
  like a hang.
- A case that hangs ends the run. It is reported and looked into, not retried.
- Anything that feeds crafted or extreme input to the package runs in a child process under a memory
  watchdog that samples its size and kills it at 4 GB (`tests/watchdog.py`). A time limit alone does
  not count, and `ulimit -v` is not a limit on macOS. A killed run, a missing result or an exit status
  of 137 means stop and find out why, not run it again.
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
reproduced every common row. Facts from it (Excel 16.114):

- **The font height.** A row without `ht` takes the Normal font's height: Calibri 9 to 16 give
  12, 14, 15, 16, 19 and 21 pt; Aptos Narrow gives 13, 14, 15, 16, 19 and 22 pt.
- **`defaultRowHeight` counts only with `customHeight="1"`.** Without it, Excel ignores the
  stored value (15, 20 and 30 all give the font height, at Calibri 11 and 14). With it, every row
  without `ht` takes the stored value exactly, rows with cells included (12.75 stays 12.75).
- **A `<row>` element matters, not a value.** Rows with an element (a value, an empty cell, or
  none) and rows without one differ in three cases; otherwise all rows are equal:
  - without `<sheetFormatPr>`, a row with an element has the font height, and a row without one
    the `standard height`, 16 pt for every font: Excel for Mac's own default, Calibri 12,
    probably the application's standard font setting, which the capture does not vary;
  - with `thickTop` or `thickBottom`, a row without an element is 1 pt taller per attribute;
  - with `zeroHeight`, a row without an element is hidden (0 pt), and a row with one has the font
    height, even when `customHeight="1"` stores another.
- **A row is fitted to its cells when Excel opens the file** (A1 in another font, wrapped or
  merged, in a Calibri 11 workbook): a 14 pt cell makes its row 19 pt, bold or not, even when
  the cell is empty but styled; a 9 pt one leaves it at 15; wrapped text gives 16 pt per line
  (three lines 48, six lines 96); in a merge over A:C the 14 pt font still counts, the wrapped
  lines do not; a stored `ht` of 15 stays 15.

## The screen

The PDF is Excel's print layout. For Excel-faithful output the truth is the screen, and print and
screen differ in places: a row 15 pt high in the file prints 14 pt apart. The next step is an
automated screen capture of the test sheets, a window capture at 100 % zoom with pixel analysis of
the gridlines and text extents, checked against the PDF route on a sample.

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
Excel would refuse the file or offer to repair it. Facts from it (Excel 16.114): both controls open
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
| `package` | 68 | a control, a minimal `<!DOCTYPE>`, bzip2 and LZMA parts, a part named twice, three kinds of false size, widths, heights, font sizes and sheet defaults no sheet can have, a control |
