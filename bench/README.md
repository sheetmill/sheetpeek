# Benchmarks

Scripts that rerun the measurements in `ARCHITECTURE.md` §5 (openpyxl) and §7.3 (the HTML
output). Each writes its workbooks into a directory given as an argument, by default a temporary
one, never into the repository.

- `readonly_features.py`: what openpyxl exposes in normal and read-only mode.
- `load_modes.py`: time and peak memory of reading a large workbook in each of openpyxl's modes,
  and of xlio's cell stream (what sheetpeek's reader runs) and sheetpeek's package check.
- `date_fidelity.py`: how often openpyxl's date conversion changes what a format displays.
- `html_size.py`: size and time of the HTML output for a 5,000 × 20 sheet, in two cases:
  `plain` (formatted numbers, dates, text that spills, a bold filled header) and `features`
  (every cell uses a feature 0.2.0 did not draw: borders and diagonals, pattern and
  gradient fills, underlines, strikethrough, superscript and subscript, rich-text runs, wrapped,
  shrunk, distributed and angled text, links, outlines, merges, a grid color).

## HTML output

`python bench/html_size.py DIR 5000 20 plain|features`, each run in a fresh process under
`tests/watchdog.py`, best of three; Python 3.11, openpyxl 3.1.5, xlformat 0.3.0, Apple M4 Pro,
macOS. Times vary by about 5 % between runs and from one day to the next, and peak RSS by about
5 MB, so rows compare only when measured together: each version below was measured beside the
next. 0.2.0 is
the release from PyPI, which draws none of the newer features: on the `features` sheet it shows
what the same cells cost without them. `2421aed` is an intermediate version, measured beside 0.2.0 and beside the next.

| case | version | HTML | read | layout | render | peak RSS |
|---|---|---|---|---|---|---|
| plain | 0.2.0 | 11.1 MB | 0.55 s | 1.70 s | 0.33 s | 237 MB |
| plain | `2421aed` | 11.1 MB | 0.60 s | 1.76 s | 0.28 s | 248 MB |
| plain | now | 11.1 MB | 0.60 s | 1.76 s | 0.28 s | 246 MB |
| features | 0.2.0 | 21.6 MB | 0.54 s | 2.03 s | 0.50 s | 285 MB |
| features | `2421aed` | 33.5 MB | 0.66 s | 2.60 s | 0.62 s | 356 MB |
| features | now | 33.8 MB | 0.67 s | 2.69 s | 0.63 s | 358 MB |

With a frozen pane (`... 5000 20 features frozen`, the sheet frozen at B2, measured beside the
`features` case on the same version): 36.2 MB against 33.8 MB, render 0.73 s against 0.66 s, best
of three; the copies hold only the frozen rows', columns' and corner's cells, fills and borders.

**Against 0.2.0**, on the plain sheet layout is 4 % slower, about the runs' variation, and render
is 13 % faster, after removing per-cell overhead the newer features had added: a second
measurement of every line's width, a `Font` hashed several times a cell (fonts, row heights and
line metrics are now found by the identity of the style's font), the bidirectional check on
ASCII text, the general baseline routine for a single bottom-aligned line, and in HTML the link
and rule bookkeeping for cells and runs that have none. Read time grew with the structure scan,
which now reads every cell's type and the rich-text runs (in `load_modes.py` on 100,000 × 20
cells: 2.46 s with 0.2.0, 3.13 s now); the scan is replaced when sheetpeek moves to the shared
xlsx reader.

**Against `2421aed`**, the plain sheet is byte-identical at the same times. On the `features`
sheet the wrap rule (a word joins a line only if the space after it fits too) breaks the wrapped
cells of one column into one more line: 4,999 more runs, 1 % more HTML, 3 % more layout time;
peak memory is within the run-to-run spread.
