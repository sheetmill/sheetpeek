# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- `ReadLimits`, and `limits=` on `read`, `open` and `render`: `max_parts`, `max_part_bytes` and
  `max_total_bytes` for the package, and `max_cells` for the grid area of what is read and laid out.
  `sheetpeek view` and `serve` take `--max-parts`, `--max-part-bytes`, `--max-total-bytes` and
  `--max-cells`.
- `PackageError`, a `ValueError`, and `LimitError`, a `PackageError`.
- `Sheet.max_cells`, which a layout of the sheet keeps to.

### Changed
- Every file is checked before anything parses it. These are now refused: a part declaring a DTD,
  in any encoding, under any part name; a stated size that is not the part's size; a compression
  method other than stored or deflate; an encrypted part; a part named twice. Excel opens none of
  those it was tried with (all but encryption).
- Requires xlformat 0.3. Merged widths come from `column_metrics_from_points`, and text widths are
  whole points, as Excel draws them.
- `Book` opens a path once, so a book is for one thread at a time.
- A file openpyxl cannot read raises `PackageError`, with openpyxl's error as its cause, in place of
  whatever openpyxl raised: `BadZipFile`, `KeyError`, `OSError`, `ParseError`,
  `InvalidFileException`. A path that cannot be opened still raises its `OSError`.
- The file's extension is no longer checked: a valid workbook reads under any name. `.xls` and
  `.xlsb` files keep their advice to convert them, recognized by their content, then their name.
- `sheetpeek serve` has lower limits, refusing for example a sheet of more than 1,048,576 cells
  unless `?ref=` asks for less. A refused file is a 422.
- `sheetpeek serve` reads two workbooks at a time, and answers 503 with `Retry-After` when no slot
  frees up within 2 seconds.
- Layout no longer formats the empty columns it may add for spilled text, so a long text in a sheet
  of many rows costs much less.

### Fixed
- The documentation names the formula engine, `xlrecalc`, now that it is published.
- OOXML's `_xHHHH_` escapes in text are decoded, as Excel decodes them: `a_x000D_b` holds a carriage
  return, not the literal text. Two known deviations come from openpyxl, in shared strings:
  `_x005F_x0041_` shows `A` where Excel shows `_x0041_`, and whitespace without `xml:space` stays.
- Control characters are not drawn, as in Excel, and do not count toward a text's width. `Cell.value`
  keeps them.

### Security
- `sheetpeek serve`, bound to a loopback address, answers only to `localhost`, `127.0.0.1`, `[::1]`
  and that address, with 403 otherwise. Any host name was answered, so DNS rebinding could let a web
  page read the folder's sheets.
- `sheetpeek serve` no longer shows absolute paths: the index shows the folder's name, and errors
  from openpyxl, which name the full path, drop the folder.
- A crafted file or `?ref=` could use memory without bound; see the limits above. In particular a
  small file could ask for a dense grid of 17 billion cells, `ZipFile.read` could decompress 2 GB
  or more from a part stating 1,000 bytes, and a large entity could expand within expat's limit.

## [0.1.0] - 2026-10-31

First release.

### Added
- **Reading.** `read` and `open` read a worksheet into a `Sheet`: cached values (dates as the
  serial numbers Excel stores, error values as `xlformat.ExcelError`), number formats, fonts,
  fills, alignment, column widths, row heights, hidden rows and columns, merged ranges and
  hyperlinks, in constant memory while the file is read. `ref=` reads part of a sheet.
- **Layout.** `Sheet.layout` lays a sheet out as Excel for Mac shows it, in one of three modes:
  - `excel`: Excel's geometry in points, with column and row positions (`x_pt`, `width_pt`,
    `y_pt`, `height_pt`), each text as a `TextRun` (position, baseline, width, clip window,
    anchor, font, color) and the boundaries text crosses (`RenderedRow.crossed`);
  - `characters`: whole-character columns, as Excel's Column Width dialog gives them for the
    workbook's Normal font, each cell recording the characters drawn in its column;
  - `none`: every value in full, without widths.

  Number formats come from xlformat, at the column's width: `#` for a number that does not fit,
  `General` reduced as Excel reduces it. Text spills into empty neighbors as Excel spills it,
  and past the last used column into added columns (`spill_edge="extend"`); merges, center
  across selection, indent, wrapping, hidden and width-0 columns, and fills follow Excel. Row
  heights follow Excel: the Normal font's height unless `<sheetFormatPr customHeight>` says
  otherwise, and a row fitted to its cells' fonts and wrapped lines when it has no height of
  its own.
- **Backends.** `render` with `text`, `rich`, `html`, `markdown`, `csv` and `tsv`;
  `register_backend`, and any object with a `render` method, for others. The text, rich and
  HTML backends take `column_headings` and `row_headings`.
- **HTML.** `HTMLBackend` writes a static page or fragment from `excel` mode: spilled text over
  its neighbors' own fills, clipping where Excel clips, competing spills overlapping, geometry in
  whole CSS pixels, every text anchored where Excel anchors it, and Calibri text spaced to
  Excel's widths from a bundled table of Carlito's advance widths. `color_scheme` (`"light"` or
  `"dark"`) colors the headings. Its CSS custom properties (`--sheetpeek-*`) and class names
  (`sheetpeek`, `sp-*`) are public interface: changes to them are listed here like API changes.
- **Command line.** `sheetpeek view` shows a sheet in any backend, and `sheetpeek serve` serves a
  folder of workbooks as HTML pages, read-only, bound to 127.0.0.1 by default and reading only
  `.xlsx` files inside the folder.
- **Captures.** The layout rules are tested against Excel 16.114 for Mac: its PDF export for
  spill, clipping, alignment, merges, fills and baselines (179 rows), its scripting interface
  for row heights (55 workbooks), theme colors with tints (96 cells), merged widths (150 cells)
  and error values (56 cells). Every test run reports the known deviations by count.

### Known limits
- Excel for Mac only: widths, glyph advances and row heights are measured there, not on Windows.
- Cached values only: a formula without a cached value shows as empty, with a warning.
  Recalculation through a formula engine, now in development, is planned.
- Calibri and Aptos Narrow are measured; other fonts are measured as Calibri at the same size,
  with a warning. Aptos Narrow text is not spaced in HTML.
- The gridline and heading colors, the heading sizes, and the text insets and baselines come
  from Excel's print layout until they are measured on screen.
- Borders, conditional formatting, charts, images and comments are not shown; rich text within
  a cell takes the cell's font.

[Unreleased]: https://github.com/sheetmill/sheetpeek/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/sheetmill/sheetpeek/releases/tag/v0.1.0
