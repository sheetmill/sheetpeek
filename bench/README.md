# Benchmarks

Scripts that rerun the measurements in `ARCHITECTURE.md` §5 (reading) and §7.3 (the HTML
output). Each writes its workbooks into a directory given as an argument, by default a temporary
one, never into the repository.

- `load_modes.py`: time and peak memory of reading a large workbook through xlio's cell stream
  (what sheetpeek's reader runs) and sheetpeek's package check, with openpyxl's modes for
  comparison.
- `large_workbook.py`: writes a workbook of 14.8 million cells over six sheets and 264,408 rows,
  most of them styled and empty, the shape of the largest workbooks met in practice (an export
  with a style applied to whole columns); a scale multiplies its rows.
- `read.py`: time and peak memory of reading every sheet of a workbook in a `recalc_on_load`
  mode, and with `--trace` the peak of Python's live allocations on a sheet (§5.2, §5.4, §5.7
  give the figures for `large_workbook.py`'s workbook).
- `layout.py`: time and peak memory of laying out every sheet of a workbook in `excel` mode,
  read time apart (§5.2 gives the figures for `large_workbook.py`'s workbook).
- `date_fidelity.py`: how often a date's conversion through `datetime` (openpyxl's) changes what a
  format displays, which is why the reader keeps serials.
- `html_size.py`: size and time of the HTML output for a 5,000 × 20 sheet, in two cases:
  `plain` (formatted numbers, dates, text that spills, a bold filled header) and `features`
  (every cell uses a drawn feature: borders and diagonals, pattern and
  gradient fills, underlines, strikethrough, superscript and subscript, rich-text runs, wrapped,
  shrunk, distributed and angled text, links, outlines, merges, a grid color).

## Large workbooks

`python bench/large_workbook.py DIR SCALE` writes `DIR/large-SCALE.xlsx`; then
`python bench/layout.py DIR/large-SCALE.xlsx LABEL`, in a fresh process under
`tests/watchdog.py`, best of three. Releases are compared on it: each version from its own clone
(`PYTHONPATH=<clone>/src`, printing `sheetpeek.__file__`), with the dependency versions it was
published with, alternating. Scales 0.25 and 0.5 show how time and memory grow with the rows
(`ARCHITECTURE.md` §5.2).

## HTML output

`python bench/html_size.py DIR 5000 20 plain|features`, each run in a fresh process under
`tests/watchdog.py`, best of three; `... features frozen` freezes the sheet at B2. It prints the
page's size and the time and peak memory of reading, layout and rendering. Times vary by about
5 % between runs and from one day to the next, and peak RSS by about 5 MB, so two versions
compare only when measured together, each from its own clone, alternating.

The sizes it gives are in `ARCHITECTURE.md` §7.3. The measurement of record for reading is the
one in `ARCHITECTURE.md` §5.7, with its commit and environment.
