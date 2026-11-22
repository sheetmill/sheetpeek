# Benchmarks

Scripts that rerun the measurements in `ARCHITECTURE.md` §5 (reading) and §7.3 (the HTML
output). Each writes its workbooks into a directory given as an argument, by default a temporary
one, never into the repository.

- `load_modes.py`: time and peak memory of reading a large workbook through xlio's cell stream
  (what sheetpeek's reader runs) and sheetpeek's package check, with openpyxl's modes for
  comparison.
- `date_fidelity.py`: how often a date's conversion through `datetime` (openpyxl's) changes what a
  format displays, which is why the reader keeps serials.
- `html_size.py`: size and time of the HTML output for a 5,000 × 20 sheet, in two cases:
  `plain` (formatted numbers, dates, text that spills, a bold filled header) and `features`
  (every cell uses a feature 0.2.0 did not draw: borders and diagonals, pattern and
  gradient fills, underlines, strikethrough, superscript and subscript, rich-text runs, wrapped,
  shrunk, distributed and angled text, links, outlines, merges, a grid color).

## HTML output

`python bench/html_size.py DIR 5000 20 plain|features`, each run in a fresh process under
`tests/watchdog.py`, best of three; `... features frozen` freezes the sheet at B2. It prints the
page's size and the time and peak memory of reading, layout and rendering. Times vary by about
5 % between runs and from one day to the next, and peak RSS by about 5 MB, so two versions
compare only when measured together, each from its own clone, alternating.

The sizes it gives are in `ARCHITECTURE.md` §7.3. The measurement of record for reading is the
one in `ARCHITECTURE.md` §5.7, with its commit and environment.
