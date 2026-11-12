# Benchmarks

Scripts that rerun the measurements in `ARCHITECTURE.md` §5 (openpyxl) and §7.3 (the HTML
output). Each writes its workbooks into a directory given as an argument, by default a temporary
one, never into the repository.

- `readonly_features.py`: what openpyxl exposes in normal and read-only mode.
- `load_modes.py`: time and peak memory of reading a large workbook in each mode, and of
  sheetpeek's structure scan and package check.
- `date_fidelity.py`: how often openpyxl's date conversion changes what a format displays.
- `html_size.py`: size and time of the HTML output for a 5,000 × 20 sheet, in two cases:
  `plain` (formatted numbers, dates, text that spills, a bold filled header) and `features`
  (every cell uses a feature 0.2.0 did not draw: borders and diagonals, pattern and
  gradient fills, underlines, strikethrough, superscript and subscript, rich-text runs, wrapped,
  shrunk, distributed and angled text, links, outlines, merges, a grid color).

## HTML output

`python bench/html_size.py DIR 5000 20 plain|features`, each run in a fresh process under
`tests/watchdog.py`, best of three; Python 3.11, openpyxl 3.1.5, xlformat 0.3.0, Apple M4 Pro,
macOS. Times vary by about 5 % between runs. 0.2.0 is the release from PyPI, which draws none of
the newer features: on the `features` sheet it shows what the same cells cost without them.

| case | version | HTML | read | layout | render | peak RSS |
|---|---|---|---|---|---|---|
| plain | 0.2.0 | 11.1 MB | 0.55 s | 1.70 s | 0.33 s | 237 MB |
| plain | before the per-cell overhead was removed | 11.1 MB | 0.60 s | 2.02 s | 0.44 s | 246 MB |
| plain | now | 11.1 MB | 0.60 s | 1.77 s | 0.29 s | 246 MB |
| features | 0.2.0 | 21.6 MB | 0.54 s | 2.03 s | 0.50 s | 285 MB |
| features | before the per-cell overhead was removed | 33.5 MB | 0.70 s | 3.08 s | 0.70 s | 361 MB |
| features | now | 33.5 MB | 0.70 s | 2.74 s | 0.65 s | 362 MB |

On the plain sheet layout is 4 % slower than 0.2.0, about the runs' variation, and render is
13 % faster. What
was removed: a second measurement of every line's width, a `Font` hashed several times a cell
(fonts, row heights and line metrics are now found by the identity of the style's font), the
bidirectional check on ASCII text, the general baseline routine for a single bottom-aligned
line, and in HTML the link and rule bookkeeping for cells and runs that have none. Read time
grew with the structure scan, which now reads every cell's type and the rich-text runs (in
`load_modes.py` on 100,000 × 20 cells: 2.46 s with 0.2.0, 3.13 s now); the scan is replaced
when sheetpeek moves to the shared xlsx reader.
