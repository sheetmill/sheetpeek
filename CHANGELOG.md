# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

The first release, 0.1.0, is being prepared.

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

[Unreleased]: https://github.com/sheetmill/sheetpeek/commits/main
