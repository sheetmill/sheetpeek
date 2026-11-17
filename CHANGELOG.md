# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Notices: nothing a workbook uses is dropped silently. `Sheet.notices` and
  `RenderedSheet.notices` hold a `Notice` for each kind of thing the file uses that sheetpeek does
  not read yet, or that the layout approximates, and `sheetpeek.notices(source, to=...)` adds what
  that output cannot show, without rendering. A notice has a dotted code (`FEATURES` lists them),
  an outcome (`omitted` or `approximated`), a reason, the cells and the output.
- Reported when a file uses them, until they are drawn: phonetic guides, rich-text runs in inline strings (which openpyxl does not keep), row and column formats, sheet views (hidden zeros, right to left, formulas,
  page layout, split panes), conditional formatting, filter buttons,
  tables, sparklines, slicers, timelines, pictures, charts, shapes, embedded objects, form
  controls, background pictures, notes and threaded comments, tab colors (drawn only on the workbook page), dialog sheets, macro sheets, and any sheet element or extension sheetpeek
  does not recognize. A file is reported for what it uses, not what it contains: margins, page
  setup, protection and validation, which draw nothing, are never reported. A linked picture is
  reported as refused: Excel fetches the linked file, sheetpeek never does.
- `render(on_notice=...)`: a callable that receives each notice, or `"ignore"`. Without it,
  `render` issues one `NotShownWarning` when there is something to report.
- A backend may declare `support`, a mapping from feature code to `Support`; one without it has
  every feature reported as not shown. `Book.notices` lists chartsheets.
- `HTMLBackend(notices=True)` lists what the page does not show after the sheet; `sheetpeek serve`
  turns it on.
- `sheetpeek view --notices summary|full|none`, `--strict` (exit status 3 when anything is not
  shown) and `--suppress CODE`.
- `CellRange.intersection`.
- Hyperlinks are drawn. HTML wraps a linked cell's text in an anchor, which adds no color or
  underline of its own, since the cell's format draws a link as in Excel; its `title` is the
  file's tooltip. A link to a cell on the page goes to that cell (the cell gets the id `sp-` and
  its coordinate). `HTMLBackend(new_tab=True)` opens external links in a new tab. `rich` draws
  terminal hyperlinks, and Markdown `[text](url)` links; plain text, CSV and TSV keep the text.
- By default only `http`, `https` and `mailto` targets become links. `javascript:`, `data:` and
  any other scheme are never linked: drawn as plain text, never written into the output, and
  reported (`link.scheme`).
- `local_links=True` on `HTMLBackend`, `RichBackend` and `MarkdownBackend` also links `file:`
  URLs and relative, UNC and drive-letter paths, which Excel follows; such links point into the
  viewer's own files and reveal the paths the workbook names, and a target starting with `//`
  resolves to a network host when a page is served, as an `http` link would. Without it they are
  drawn as text and reported by each output (`link.local`).
- A `HYPERLINK()` formula whose first argument is a string links its cell. One with a computed
  target, or in a shared or array formula, is reported (`link.computed`); a cell with both a
  hyperlink element and a `HYPERLINK()` formula takes the element's (`link.conflict`). A link on
  an empty cell is reported (`link.empty_cell`), and so is a link to another sheet or a defined
  name, which HTML shows as a link that goes nowhere (`link.elsewhere`).
- Outlines (grouped rows and columns). `Row` and `Column` carry `outline_level` and `collapsed`,
  and `Sheet.outline` an `Outline` (summaries below or above, right or left, whether the file
  shows the symbols, the depths). The layout derives the groups and gives
  `RenderedSheet.outline`, the bars, ticks, dots and buttons of each row and column
  (`RenderedOutline`, `OutlineMark`). The HTML backend draws Excel's outline gutter beside the
  headings, at Excel's screen sizes in CSS pixels, with level buttons when both headings are
  drawn; `rich` and plain text draw it in characters; Markdown, CSV and TSV report it as not
  shown. A range that cuts a group reports the button it leaves out (`outline.cut`). A level is read as
  Excel reads it: modulo 8, and one that is not decimal digits as none.
- `buttons="static" | "none"` on `HTMLBackend`, `RichBackend` and `TextBackend`: whether to draw
  Excel's buttons, which do nothing in a static output; for now the outline gutter. New CSS custom
  properties `--sheetpeek-outline-paper`, `--sheetpeek-outline-ink`,
  `--sheetpeek-outline-button`, `--sheetpeek-outline-button-paper`, and classes `sp-ol`,
  `sp-olbutton`, `sp-ollevel`.
- Pattern fills: the 17 patterns are drawn in HTML as Excel draws them, from tiles captured on
  the screen, anchored at the sheet's origin, a tile pixel a whole screen point at any zoom.
- Gradient fills: linear gradients at Excel's angle, which is taken in the unit cell, and path
  gradients in rectangular bands.
- Borders: the four sides in all 13 styles and their colors, and diagonals, drawn in HTML as a
  layer of lines in CSS px. Where two cells' sides meet, the layout keeps the one Excel draws
  (the stronger style, then the darker color), keeps each cell's own sides around a merge, and
  drops the vertical sides that spilled text crosses. `slantDashDot` is drawn as
  `mediumDashDot`, and reported.
- `CellStyle.fill_spec` (a `Pattern` or a `Gradient`) and `CellStyle.borders` (a `Borders` of
  `Side`s), and `RenderedSheet.borders`, a `RenderedBorders` of the resolved lines in
  `excel` mode. New public types: `Pattern`, `Gradient`, `Fill`, `PatternKind`, `Borders`, `Side`,
  `BorderStyle`, `RenderedBorders`, `Diagonal`.
- Census codes `fill.pattern`, `fill.gradient`, `border`, `border.diagonal` and
  `border.slant_dash_dot`, declared by every output: `rich` shows a pattern as one blended color
  and a gradient as its midpoint color; borders are drawn only in HTML so far.
- The HTML classes `sp-fills` and `sp-borders`.
- Text styles are drawn in HTML (`excel` mode), from captures of Excel for Mac: superscript and
  subscript (two thirds of the size, raised or lowered, rows fitted taller); accounting
  underlines (under text across the cell, under a number across the number, 3 pt lower, rows 3 pt
  taller); shrink to fit (the largest seventh of the size that fits, in place of `####`; shrunk
  text never spills); vertical justify and distributed (lines spread from the top to the bottom,
  text wrapped); indent on horizontally distributed text; reading order (`general` text that
  starts right-to-left aligns right, and is drawn as a right-to-left paragraph unless the order
  is left to right).
- `Font.vert_align`; `CellStyle.shrink`, `CellStyle.reading_order` and
  `CellStyle.justify_last_line`; `TextRun.scale`, `TextRun.rules` and `TextRun.direction`; and
  `TextRule`, an underline or strikethrough of a run.
- The HTML classes `sp-rule` and `sp-double`: underlines and strikethrough.
- Census codes `font.vert_align`, `text.shrink` and `text.reading_order`, which every output
  declares: `rich` and plain text clip shrunk text at full size and keep `#` for a number too
  wide, and draw superscript and subscript at full size (`rich`) or plain (text); Markdown draws
  them as plain text.
- Theme fonts: a font of the theme's body or heading scheme is drawn in the theme's typeface,
  whatever name it stores, as Excel draws it. `Font.name` is that typeface; the new
  `Font.scheme` and `Font.stored_name` keep what the file says. The Normal font follows the theme
  too, so column widths take the theme font's digits; a theme typeface xlformat has no glyph
  table for is measured as Calibri and reported (`font.no_metrics`).
- The grid color: `Sheet.grid_color` reads `<sheetView defaultGridColor="0" colorId>` through
  the workbook palette. HTML draws the gridlines in it, one device pixel wide as Excel does;
  `rich` draws the column separators in it.
- `format.palette`: a number format's `[ColorN]` that the workbook's palette recolors is reported,
  since the layout draws section colors from Excel's default palette.
- Rich-text runs are drawn, from captures of Excel for Mac. A run without properties takes the
  cell's font, and a run with properties takes nothing from it. In `excel` mode each run is
  measured in its own font for spill, `####`, wrapping and shrink to fit, and each line takes the
  height of its largest run. HTML draws a nested span per run on one baseline, superscript and
  subscript runs raised and lowered; Markdown with `emphasis=True` marks each run's bold and
  italic.
- `Cell.runs`, the runs of a cell's text, and `TextRun.spans`, each piece of a line with its
  font, width, size and baseline shift. New public type: `TextSpan`.
- Census code `text.runs`, which every output declares: `rich` draws runs in the cell's font,
  plain text, CSV and TSV leave them out, and Markdown shows bold and italic runs with
  `emphasis=True`. A cell with runs is also counted under each font feature its runs use.
- Rotated and stacked text are drawn in HTML (`excel` mode), from captures of Excel for Mac.
  Angled text is one line turned about the start of its baseline, placed in the cell by its
  turned box (general text left for positive angles below 90 and at -90, right otherwise),
  spilled into empty neighbors in the direction it runs, clipped at a custom-height row, with
  fitted rows grown to it; an angled value shows in full, never `#` or shortened. Stacked text
  draws one character a line, centered. `rich`, plain
  text, Markdown and CSV draw both horizontally and report them. Wrapped angled text wraps along
  the turned direction, as Excel draws it, and leaves its row as it is; wrapped stacked text is
  drawn on one line and reported (`text.stacked_wrapped`). The side borders of an angled cell slant with
  the text into a parallelogram and the text sits between them, as Excel draws them
  (`SlantedBorders`, `RenderedBorders.slanted`); a merge's borders stay upright and its text is
  placed as without borders, reported (`border.slanted`).
- `CellStyle.rotation` and `CellStyle.stacked`; `TextRun.angle`.
- Census codes `text.rotation`, `text.stacked` and `border.slanted`, which every output declares.
- `Book.sheets`: every sheet in workbook order as a `SheetInfo` (name, kind, state, tab color).
  Kinds come from the workbook's relationship types: `worksheet`, `chartsheet`, `dialogsheet`,
  `macrosheet`. Tab colors are `RRGGBB` through the palette and theme; `auto` is `None`.
- `sheetpeek.render_workbook()`: every sheet on one static HTML page, with Excel's tab strip
  below the sheets, or above them with `tabs="top"` (visible sheets in file order, tab colors as
  Excel draws them; hidden sheets with `hidden_sheets=True`, hidden rows and columns with
  `show_hidden=True`). Links to a cell of another sheet on the page go there. Chartsheets and dialog
  sheets get a tab and a line saying they are not shown.
- `sheetpeek view --all-sheets` writes that page (`-t html` only; `--hidden-sheets` for hidden
  sheets).
- `NotAWorksheetError`, a `ValueError`, raised for a chartsheet or a dialog sheet.
- The notice `sheet.hidden`: hidden sheets the workbook page leaves out.
- The sheet view and page layout are read: `Sheet.view` (a `SheetView`: the saved
  zoom, whether zeros show, a frozen pane's rows and columns), `Sheet.print_areas` (the print
  area's ranges; `Sheet.print_area_problem` says why a part was left out: `#REF!`, a formula,
  another sheet), and the manual page breaks, `Sheet.row_breaks` and `Sheet.column_breaks`.
- HTML draws the zoom saved with each sheet, as Excel opens it; `HTMLBackend(zoom=...)` (10 to
  400), `sheetpeek view --zoom` and `serve`'s `?zoom=` override it. Cells, text and headings
  scale; lines, pattern tiles and outline symbols keep their screen size, as in Excel.
- HTML draws a frozen pane: the freeze lines, and on a whole page the frozen rows and columns held
  in place while the page scrolls, through CSS `position: sticky` (no script).
  `HTMLBackend(freeze="panes" | "static")` chooses; a fragment draws only the lines by default.
  Every row and column of the range is drawn; the saved scroll position is not followed.
- `render(print_area=True)`, `sheetpeek view --print-area` and `serve`'s `?print_area=1` show the
  sheet's print area instead of its used range, as Excel prints it: each of its ranges as a
  block, in the file's order, whole columns or rows cut to the used range. Off by default.
- `HTMLBackend(page_breaks=True)` and `sheetpeek view --page-breaks` mark manual page breaks
  with a solid blue line, as Excel's page break preview draws them. Off by default.

### Changed
- Requires xlformat 0.4. Through it, a workbook whose Normal font is 18 pt or more gets Excel's
  wider column padding: its columns without a stored width are wider (2 pt at Calibri 20).
- `Cell.value` of a whole number is an `int` only below 2**53 in magnitude, a `float` above:
  it comes from the stored double, not the file's digits, so `9995806932108624` (one value in
  a corpus of 15,966 workbooks) is now `9995806932108624.0`. The rendered text is the same.
- HTML draws each sheet at the zoom saved with it, as Excel opens it, where 0.2.0 drew every
  sheet at 100 %: a sheet saved at another zoom renders at another size. Pass
  `HTMLBackend(zoom=100)` (or `sheetpeek view --zoom 100`) for 0.2.0's size.
- The HTML backend writes a page faster: 13 % on a plain 100,000-cell sheet
  (`bench/html_size.py`), with layout within 4 % of 0.2.0 on it; a sheet pays for the newer
  features only where it uses them.
- Tints stored as 16-bit fractions (0.6 is stored as 0.59999389...) give the colors Excel draws:
  the tinted lightness is truncated after a small tolerance. Some fills and font colors with such
  tints were a lightness level or two off.
- `Cell.hyperlink` is a `Hyperlink` (`target`, `location`, `tooltip`, `display`, `source`), or
  `None`, in place of a string that joined the target and the location with `#`. `Hyperlink` is
  exported from `sheetpeek`.
- A sheet with grouped rows or columns whose file shows outline symbols starts with the outline
  gutter in HTML, `rich` and plain text output; `buttons="none"` leaves it out.
- `sheetpeek view` reports on standard error after the output, one line per kind with its code
  first (`sheetpeek: border: borders not shown (3 cells): not implemented yet`), in place of
  `sheetpeek: warning: ...` lines before it.
- `Sheet` and `RenderedSheet` take `notices` in place of their `warnings` field; `warnings` is a
  read-only property.
- openpyxl's warning that it drops a sheet extension (conditional formatting, sparklines,
  slicers and others it does not read) is no longer passed on to the caller while sheetpeek reads a
  sheet; those extensions are reported as notices instead.
- `CellStyle.fill` of a pattern fill is the blend of its two colors by the pattern's share,
  not its foreground: a `gray0625` fill, 6.25 % black dots, was shown as solid black. A
  gradient's is its midpoint color, where it was no fill.
- A solid fill in indexed color 64 is no fill, and a fill index past the palette is black, as
  Excel draws them; an index past the palette was no fill.
- Underlines and strikethrough in HTML are 1 px rules at Excel's offsets, not the browser's text
  decorations, so they stay one pixel thick at any page scale; a single or double underline runs
  1 pt past its text, as Excel 16.115 draws it.
- Text whose first strong character is right-to-left (Hebrew, Arabic) with `general` alignment
  is right-aligned in every output, as Excel aligns it.
- Vertically justified text in `characters` mode is placed at the top, and distributed text
  centered or spread, instead of at the bottom; both wrap.
- HTML headings take Excel's sizes: the column band 21 pt and the row band 16 pt plus 3 pt a digit
  of the last row, at Calibri 11, growing with the Normal font's size. Their colors are Excel's:
  dark `#161616`, `#ffffff` and `#565656`; light `#f7f7f7`, `#5e5e5e` and `#e0dfdd`.
- A workbook palette replaces the default entry by entry; a `<colors>` element without
  `indexedColors` keeps the default palette instead of drawing indexed colors black.
- A row fitted to a Calibri 20 cell is 26 pt, as Excel fits it, instead of Calibri 16's 21.
- A row Excel fits grows to the lines of horizontally distributed text, which wraps when drawn
  (two lines of Calibri 11 take 32 pt); it was not fitted to them.
- A dialog sheet is no longer drawn as a blank worksheet: `Book.sheet()` raises
  `NotAWorksheetError` for it, as for a chartsheet, and `sheet.dialogsheet` is a workbook-level
  notice in `Book.notices`.
- The default sheet skips an active chartsheet or dialog sheet: it is the first visible worksheet
  or macro sheet.
- `sheetpeek view --list` shows each sheet's state and kind when it is not a visible worksheet
  (`Chart  (chartsheet, not shown)`, `Data  (hidden)`), where it printed the name alone.
- `serve` marks a chartsheet or dialog sheet in its sheet list instead of linking to an error.
- Wrapped text breaks where Excel breaks it (`excel` mode, so HTML and fitted rows): a word joins
  a line only if the line fits with the space after the word. "shrink and wrap both set here"
  in a default column takes 4 lines, as in Excel, where it took 3.
- Distributed and justified lines share the room among their gaps as Excel does, counting the
  space that ends a broken line: a one-word line inside distributed text sits at the left inset,
  where it was centered; a one-word last line stays centered. Distributed text with an indent
  breaks between the indents at both ends.

### Deprecated
- `Sheet.layout(spill_edge=...)`: use `range_edge`, the same option under a name that also covers
  objects reaching past the range. `spill_edge` still works, with a `DeprecationWarning`, and is
  removed in 1.0; `SpillEdge` is an alias of `RangeEdge`.
- `Sheet.warnings` and `RenderedSheet.warnings`, now views of the notices that were warnings
  before. Use `notices`.

### Fixed
- Built-in number formats are Excel's (xlformat's table, captured from Excel for Mac), no
  longer openpyxl's: a typed date, which Excel stores as format 14, shows as `1/2/24` where it
  showed `01-02-24`; formats 44 and 47 are Excel's, and the built-in ids from 50 to 163, read as
  General before, are the formats Excel gives them.

## [0.2.0] - 2026-11-04

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
- `sheetpeek serve` has limits, lower than the library's, where it had none: it refuses for example
  a sheet of more than 1,048,576 cells unless `?ref=` asks for less. A file it refuses or cannot read
  is a 422; some of those were a 404.
- The documentation names the formula engine, `xlrecalc`, now that it is published.
- The documentation and `sheetpeek serve --help` say that `serve` is not for production.
- `sheetpeek serve` reads two workbooks at a time, and answers 503 with `Retry-After` when no slot
  frees up within 2 seconds.
- Numbers 0.1.0 took as given follow Excel: a negative row height is its magnitude, and a font size
  of -1 is size 1. A font size of 0 or 410, a `baseColWidth` of 300 and a `defaultRowHeight` of NaN or
  infinity are refused, as Excel refuses them.
- Layout no longer formats the empty columns it may add for spilled text: 300 rows with one long text
  take 2.2 s and 293 MB in `excel` mode, from 5.0 s and 793 MB, and a little less in `characters` mode.

### Fixed
- A number no column, row or font can have no longer raises a raw `ValueError`, `OverflowError` or
  `ZeroDivisionError` from the layout: a column width or `defaultColWidth` of NaN, infinity, or a
  negative or enormous number; a row height, or a `defaultRowHeight` with `customHeight`, of NaN,
  infinity or an enormous number; a font size of NaN, infinity, an enormous number, or a fraction such
  as 0.5 or -0.5. Each is taken as Excel takes it, or refused where Excel refuses it.
- A column in `characters` mode is at most 255 characters, Excel's own limit. A width of 1e4 made a
  column of 9,999 characters, which cost 936 MB at 10,000 rows.
- OOXML's `_xHHHH_` escapes in text are decoded, as Excel decodes them: `a_x000D_b` holds a carriage
  return, not the literal text. Two known deviations come from openpyxl, in shared strings:
  `_x005F_x0041_` shows `A` where Excel shows `_x0041_`, and whitespace without `xml:space` stays.
- Control characters are not drawn, as in Excel, and do not count toward a text's width. `Cell.value`
  keeps them.

### Security
- A cell's text could put control characters into the output of every backend: DEL and the C1 controls
  (U+0080 to U+009F, among them U+009B, the one-character form of `ESC [`), which XML can hold. They
  are removed from what a cell shows. ESC and the other C0 controls, apart from tab and line breaks,
  could not get there: XML cannot hold them, and `_x001B_` stayed literal text. Decoding such an
  escape makes them reachable, and they are removed too.
- Markdown output shows a cell's text literally. Only `\\` and `|` were escaped, so a workbook could
  put HTML, links and images in the output through a cell such as `<SCRIPT>alert(2)</SCRIPT>` or
  `[a](javascript:alert(1))`, and a line break of its own into the table.
- A theme color is used only if it is one. Any text in a crafted theme's `srgbClr` or `sysClr` went
  into the HTML as it stood, inside a `style` attribute, so a workbook could put markup and scripts
  in the page, and it crashed the `rich` backend.
- `sheetpeek serve`, bound to a loopback address, answers only to `localhost`, `127.0.0.1`, `[::1]`
  and that address, with 403 otherwise. Any host name was answered, so DNS rebinding could let a web
  page read the folder's sheets.
- The command line prints no control characters or line breaks from a workbook, a file name or a
  client: sheet names from `--list`, file names in messages, the sheet and font names in HTML, and
  `sheetpeek serve`'s log lines, which any HTTP client could write to. `Book.sheet_names` and
  `--sheet` keep the name as the file has it.
- `sheetpeek serve` no longer shows absolute paths: the index shows the folder's name, and errors
  from openpyxl, which name the full path, drop the folder.
- A width such as `99999999999`, on a column or as `baseColWidth`, made `characters` mode allocate
  its width in spaces, and a merge as large as the sheet made any layout of the sheet walk 17 billion positions, whatever range
  was read.
- A crafted file or `?ref=` could use memory without bound; see the limits above. In particular a
  small file could ask for a dense grid of 17 billion cells, a 1 MB part stating a far smaller size
  could take over 2 GB as `ZipFile.read` decompressed it, and a large entity could expand within
  expat's limit: a 12 KB file gave one cell a string of 755 million characters.

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

[Unreleased]: https://github.com/sheetmill/sheetpeek/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/sheetmill/sheetpeek/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/sheetmill/sheetpeek/releases/tag/v0.1.0
