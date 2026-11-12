# Architecture

Rules about Excel's layout cite the capture that pins them (§4); a rule no capture has pinned yet
is marked **unverified**. Facts about openpyxl (§5) were measured
with openpyxl 3.1.5 on Python 3.11, macOS, Apple M4 Pro, with the scripts in `bench/`
(`readonly_features.py`, `load_modes.py`, `date_fidelity.py`), which rerun them;
`bench/html_size.py` measures the HTML output (§7.3).

## 1. Scope and principles

sheetpeek reads an xlsx with openpyxl and shows a sheet the way Excel shows it, in four output
formats: HTML, Markdown, `rich` (terminal) and plain text (fixed-width and delimited). A small
CLI views a file and serves a folder of workbooks as HTML.

- **One model, thin backends.** Reading and layout produce a public, immutable render model.
  Backends only translate it. Anything Excel decides is decided once, in the model: what text a
  cell shows, where it is aligned, how far text spills, which cells a merge covers.
- **Cached values.** sheetpeek shows the values stored in the file. A formula without a cached value
  (a file written by openpyxl and never opened in Excel) shows as empty and is counted in the
  sheet's warnings. Recalculation through `xlrecalc`, an optional sibling, is planned.
- **Excel's number formats through xlformat**, a required dependency. It is pure Python with no
  dependencies of its own, and the package's purpose is showing what Excel shows.
- **Degrade, never fail, never silently.** A backend that cannot show a feature leaves it out or
  approximates it, as Markdown does with merges and delimited text with widths, and says so: every
  feature a file uses and an output does not show is reported as a notice (§3.7). A number format
  Excel would refuse falls back to `General`. A Normal font without glyph tables is approximated.
  The exception is a package refused as unsafe or a request over a limit (§5.5): those raise
  `PackageError` or `LimitError`, since nothing can be shown of a file that may not be read.
- **Excel is the reference.** Correct means what Excel for Mac (Microsoft 365) shows; each
  capture records its build and update channel. ECMA-376 is secondary; LibreOffice is a
  lower-trust secondary reference.
- **Scope grows by feature groups.** Each group is read, drawn in Excel-mode HTML, and given a
  stated judgment per other output, with notices for the rest. In scope and built: cell styles
  (fonts, fills, borders, alignment, rich-text runs), hyperlinks, outlines, theme fonts and the
  workbook palette, gridline colors and headings, sheet kinds, states and tab colors. In scope
  and not built yet: rotated text, tables, filter buttons, conditional formatting, notes and
  comments, pictures, panes and other sheet-view parts, print settings, pivot tables; charts
  possibly as a package of their own (§11). Out of scope: editing and any interactive UI.

## 2. Pipeline

```
xlsx ──read──▶ Sheet ──layout(mode)──▶ RenderedSheet ──backend──▶ text stream
   openpyxl read-only     source facts,      display text, overflow,     HTML, Markdown,
   + structure scan       layout-free        spill, visibility           rich, text
```

A **`Sheet`** holds what the file says, independent of how it will be shown: raw cached values,
number formats, resolved styles and geometry. **Layout** turns a `Sheet` into a `RenderedSheet`
for one width mode (§2.1): it calls xlformat, applies Excel's layout rules (§4) and computes
spill. A `RenderedSheet` keeps a reference to its `Sheet`, so a backend that wants a different
mode re-runs only the layout. Layout costs microseconds per cell. Reading the file is the
expensive step (§5).

### 2.1 Width modes

Each backend declares the width mode it needs, and the mode decides what a cell's `text` is.

| mode | width of a column | a number that does not fit | backends |
|---|---|---|---|
| `excel` | Excel's column width, measured in points with xlformat's glyph tables | as many `#` as Excel shows | HTML |
| `characters` | a whole number of characters (§2.2) | `width` × `#` (xlformat character mode) | `rich`, fixed-width text |
| `none` | none | never overflows: `full_text` | Markdown, delimited text |

`none` gives what copy and paste from Excel gives. That is the right content for formats whose
columns have no fixed width.

**Geometry in `excel` mode** (§3.6). The `excel` render model exposes absolute geometry in
points: each visible column's x position and width, each row's y position and height, and each
text as a `TextRun` with its position, baseline, width and clip window, as Excel lays them out.
Geometry backends (PDF, SVG, images) and other libraries can draw from the model without
measuring anything themselves. `characters` mode keeps its whole-character grid.

### 2.2 Column widths in characters

A terminal column holds whole characters. Excel's width is in units of one digit of the Normal
font, the number its Column Width dialog shows. xlformat's `file_column_metrics` turns the file's
`<col width>` into that number for Excel for Mac, given the Normal font: a column without `<col>`
takes `defaultColWidth` as a file width, and without that it is `baseColWidth` units (8, 53 pt at
Calibri 11), not 8.43. A column gets `round(units × scale)` characters, minimum 1 and at most 255, Excel's own limit,
capped by `max_width` (`Sheet.layout` takes both). At Calibri 11:

| file width | dialog units | points | characters |
|---|---|---|---|
| none | 8 | 53 | 8 |
| 20 | 19.17 | 120 | 19 |
| 9.140625 | 8.33 | 55 | 8 |
| 12.71 | 11.83 | 76 | 12 |
| 3.71 | 2.83 | 22 | 3 |

Digits in Calibri are tabular, so a column of width `w` holds `w` digits in both Excel and the
terminal. Text holds more in Excel, because most letters are narrower than a digit. A Normal font
xlformat does not measure is measured as Calibri at the same size, with a warning (§6.2).

## 3. The render model

All classes are frozen dataclasses with slots, importable from `sheetpeek`. Coordinates are
1-based, as in openpyxl and Excel.

### 3.1 What a cell carries

Read into `Cell` (in `Sheet`) and kept in `RenderedCell` (in `RenderedSheet`), which adds the
layout fields marked *layout*. The style fields (`number_format` to `indent`) live in one
`CellStyle` per distinct style in the file, shared by every cell that has it, and resolved once.

| field | type | source | notes |
|---|---|---|---|
| `row`, `column` | `int` | the cell's position | |
| `value` | `None \| bool \| int \| float \| str \| ExcelError` | `cell.value` in read-only mode with date conversion off (§5.3) | raw cached value; dates stay serial numbers; `xlformat.ExcelError` (a `str` enum, compared with `==`) for `data_type == "e"`, or `str` when the file's text is not shaped like an error code |
| `kind` | `"empty" \| "number" \| "text" \| "bool" \| "error"` | `cell.data_type` | a date is a `number` whose format is a date format; `d` cells (ISO 8601, rare) become serials |
| `number_format` | `str` | `cell.number_format` | |
| `font` | `Font(name, size, bold, italic, underline, strike, color)` | `cell.font`; a missing name or size inherits the Normal style's (openpyxl returns `None` for a font record without them, measured) | `color` is resolved `RRGGBB`, or `None` for automatic (§3.3) |
| `fill` | `str \| None` | `fgColor` of a `solid` pattern fill, resolved | other patterns map to their foreground color, a documented approximation; gradients are ignored |
| `fill_spec` | `Pattern \| Gradient \| None` | `cell.fill` (§4.12) | the fill itself: a pattern's kind and two colors, or a gradient's kind, angle, rectangle and stops; `fill` is its one color |
| `borders` | `Borders \| None` | `cell.border` (§4.12) | four `Side`s (style, color) and the diagonal with its two flags; `start` and `end` read as `left` and `right` |
| `h_align` | `"general" \| "left" \| "center" \| "right" \| "fill" \| "justify" \| "distributed" \| "center_across"` | `cell.alignment.horizontal` | as stored; `centerContinuous` is `center_across` |
| `v_align` | `"top" \| "center" \| "bottom" \| "justify" \| "distributed"` | `cell.alignment.vertical`, default `bottom` | |
| `wrap` | `bool` | `cell.alignment.wrap_text` | |
| `indent` | `int` | `cell.alignment.indent` | one indent level is one digit's width, 6 pt at Calibri 11 (§4.6) |
| `shrink`, `reading_order`, `justify_last_line` | `bool`, `0 \| 1 \| 2`, `bool` | `cell.alignment.shrinkToFit`, `readingOrder` (anything but 1 or 2 is 0), `justifyLastLine` | §4.13 |
| `rotation`, `stacked` | `int` (−90 to 90), `bool` | `cell.alignment.textRotation`: 0 to 90 as stored, 91 to 180 as 90 − value, 255 stacked | §4.17 |
| `font.vert_align` | `"superscript" \| "subscript" \| None` | `cell.font.vertAlign` | drawn at 2/3 size, raised or lowered (§4.13) |
| `runs` | `tuple[TextSpan, ...] \| None` | the shared string's runs (§4.16) | each run's text and resolved `Font`; `None` without runs or when every run is in the cell's font; `value` is their text joined |
| `hyperlink` | `Hyperlink \| None` | the structure scan's `<hyperlinks>` and `HYPERLINK()` formulas (§4.10) | `target`, `location`, `tooltip`, `display` (never drawn), `source` (`element` or `formula`) |
| *layout* `text` | `str` | xlformat `FormattedCell.text` at the mode's width | what the cell itself shows: `####`, reduced `General`, `*` fills applied; never clipped or spilled |
| *layout* `full_text` | `str` | `FormattedCell.full_text` | width-less; what copy and paste gives |
| *layout* `overflow` | `bool` | `FormattedCell.overflow` | `text` is a run of `#` |
| *layout* `align` | `"left" \| "center" \| "right" \| "fill" \| "justify" \| "distributed"` | `h_align`, with `general` resolved (§4.3) | the effective alignment |
| *layout* `format_color` | `str \| None` | `FormattedCell.color` through `DEFAULT_PALETTE` | the number format's section color; it replaces the font color, as in Excel |
| *layout* `merge` | `CellRange \| None` | §4.4 | the merged range the cell belongs to; `is_merge_anchor` tells the anchor |
| *layout* `spill` | `tuple[int, int]` | §4.1 | visible columns the cell's own text reaches to the left and to the right |
| *layout* `lines` | `tuple[str, ...]` | §3.5 | what is drawn **in this column**: one string of the column's width per line of the row (`characters`), or the cell's text (`none`) |
| *layout* `source` | `(row, column) \| None` | §3.5 | the cell whose text `lines` shows: this cell, a spilling neighbor, or a merge anchor |
| *layout* `fill` | `str \| None` | the cell's own fill; the anchor's for a merged cell | the background of this column, also under text spilled into it |
| *layout* `font`, `color` | | the source cell's font, and its section color else its font color | the style of the text drawn here |

**Borders and fills** are read into `CellStyle.borders` and `CellStyle.fill_spec` (§4.12), and
text styles into `Font` and `CellStyle` (§4.13). Rich-text runs inside one cell are read into
`Cell.runs` (§4.16).

### 3.2 What a sheet carries

| field | source |
|---|---|
| `name`, `state` (`visible`, `hidden`, `very_hidden`) | `ws.title`, `ws.sheet_state` |
| `used_range` | the structure scan: bounding box of every `<c>` element, from `<row r>` and `<c r>` (positions inferred when `r` is omitted, as the schema allows), widened to cover every merged range, since a merge occupies its cells even where the file has no `<c>` for them; `<dimension>` is never trusted |
| `ref` | the requested range, or `used_range` |
| `merges` | scan: `<mergeCells>` |
| `columns`: per column `width`, `hidden`, `outline_level`, `collapsed` | scan: `<cols>` (each `<col>` covers `min` to `max`) |
| `rows`: per row `height` in points, `hidden`, `outline_level`, `collapsed` | scan: `ht`, `hidden`, `outlineLevel`, `collapsed` of `<row>` |
| `outline` (`Outline`, or `None` when no row or column has a level) | scan: `<outlinePr summaryBelow summaryRight>`, `showOutlineSymbols` of the first `<sheetView>`, else of `<outlinePr>`; the deepest row and column levels in the sheet; the levels of the rows and columns just outside `ref` (§4.11) |
| `default_column_width`, `default_row_height` | scan: `<sheetFormatPr>` (`defaultColWidth`, `baseColWidth`, `defaultRowHeight`) |
| `show_gridlines` | scan: `<sheetView showGridLines>` |
| `grid_color` | scan: `<sheetView defaultGridColor colorId>`, through the workbook palette (§4.14) |
| `normal_font` | the `Normal` entry of `wb._named_styles`, else `wb._fonts[0]`; both are present in read-only mode |
| `date1904` | `wb.epoch == CALENDAR_MAC_1904` |
| `warnings` | reading and layout: formulas without cached values, approximated fonts, refused number formats, the date fallback (§5.3) |
| `cells` | a dense row-major grid over `ref` |

Column widths are kept as the file states them (file units, §6.4). Hidden rows and columns stay in
the model, marked hidden, with their cells, because spill and merges cross them (§4.5). Backends
leave them out by default and show them with an option.

### 3.3 Colors

openpyxl returns a font or fill color as a `Color` whose `type` is `rgb`, `theme` (with `tint`),
`indexed` or `auto`. Reading `.rgb` on a theme color returns the text
`"Values must be of type <class 'str'>"` instead of raising (measured), so the reader branches
on `type`.

- `rgb`: `ARGB`, alpha ignored.
- `indexed`: the workbook's custom palette if `styles.xml` has one, else the default 64-entry
  palette (`_colors.DEFAULT_PALETTE`), entry by entry (§4.14). Index 64 is the system foreground
  color, which means automatic.
- `theme`: the `clrScheme` of the theme part (`wb.loaded_theme`, raw XML), six hex digits or the
  theme is treated as unreadable, since its text would reach CSS as it stands, then `tint` applied
  as ECMA-376 §18.8.19 describes. **Captured** (suite `theme_colors`: theme indices 0 to 11 with
  tints from −0.5 to +0.8 as fills and font colors, 96 cells read back through AppleScript's
  `color of interior object` and `color of font object`): Excel's indices swap the first two
  pairs relative to the XML order (0 is `lt1`, 1 `dk1`, 2 `lt2`, 3 `dk2`), and the tint works on
  the lightness in Windows' integer HLS on a scale of 240 (`ColorRGBToHLS`), the result
  truncated, with grays rounded back. Floating-point HLS was off by one or two in a channel in 43
  of the 96 cells; the integer arithmetic reproduces all of them.
- `auto` or missing: `None`, the backend's default foreground. A dark terminal then shows light
  text, deliberately, because black text in a workbook means default text.
- A font color of theme 1 without tint is also `None`. Excel's own default font is stored as
  theme 1 ("Text 1"), so in practice it means ordinary text, not a chosen black.

The number format's `[ColorN]` uses `xlformat.DEFAULT_PALETTE`. A workbook with a custom indexed
palette recolors those in Excel; sheetpeek reports each one it changes (`format.palette`, §4.14).

### 3.4 What each backend shows

| feature | HTML | rich | fixed-width text | Markdown | delimited |
|---|---|---|---|---|---|
| width mode | `excel` | `characters` | `characters` | `none` | `none` |
| column widths | `<col>` widths, `table-layout: fixed` | column widths | padded columns | ignored | ignored |
| `####` | yes | yes | yes | no (full text) | no (full text) |
| merges | `colspan`/`rowspan` | anchor text across the merged area, placed by its alignments, on the anchor's fill; no separators inside | as rich | anchor in its cell, covered cells empty | anchor value, covered cells empty |
| spill | runs positioned from the model's geometry over the neighbors' own fills (§7.3) | text drawn across the empty neighbors, each on its own fill | as rich, across the separators | no | no |
| alignment | each run anchored where the model puts it (§7.3) | horizontal and vertical, per cell | as rich | per column in the separator row, from the first body row | no |
| font, colors, fill | CSS | `Style` | no | `**`/`*` for bold and italic, off by default | no |
| hidden rows and columns | omitted (option to show) | omitted | omitted | omitted | omitted |
| row heights | `height` on `<tr>` | lines as wrapped text needs | as rich | no | no |
| headings (row numbers, column letters) | `column_headings`, `row_headings` | the same | the same | `header="letters"` | no |
| pattern and gradient fills (§4.12) | SVG tiles and CSS or SVG gradients, in a layer under the text | one color: the pattern's blend, the gradient's midpoint | no | no | no |
| borders (§4.12) | a layer of lines in CSS px, conflicts resolved in the layout | no | no | no | no |
| hyperlinks (§4.10) | `<a>` around the cell's runs; links to shown cells go to their id | terminal hyperlinks (OSC 8), external targets only | text only | `[text](url)`, external targets only | text only |
| outline symbols (§4.11) | gutter cells beside the headings, `buttons` | a gutter of characters, `buttons` | as rich | no | no |
| superscript, subscript, shrink to fit (§4.13) | drawn | full size, shrunk text clipped | as rich | plain | no |
| underline kinds, strikethrough (§4.13) | 1 px rules at Excel's offsets | the terminal's underlines | no | no | no |
| rich-text runs (§4.16) | a nested span per run | the cell's font | no | bold and italic per run, with `emphasis` | no |

Spill in HTML is computed in the model, not left to the browser, and the browser's font must not
draw past it. Whatever the technique, the HTML backend shows the same spill as every other
backend.

### 3.5 Spill is recorded per column

Spilled text is drawn over each neighbor's own fill, and the spilling cell's fill stays in its
own column (checked by eye, §4). So the model does not record a spill as one span. In
`characters` mode the layout places every text on a character canvas per row, then cuts the
canvas back into columns.

- Each `RenderedCell` gets the characters drawn in its column (`lines`), the cell they come from
  (`source`), that cell's `font` and `color`, and its own `fill`.
- Between two visible columns, a `Gap` records the characters of any text crossing it, or
  `None` where the backend draws its separator. Its fill is the fill both neighbors share, so a
  filled header row stays continuous.

Any backend can then draw column by column: the background from the column, the text style from
the source. The `rich` backend does this per character run. HTML cannot use a single `colspan`
for spilled text, because a `colspan` cell has one background; it draws `excel` mode's runs
instead (§7.3).

A `RenderedRow` holds its visible cells, the gaps between them and its `height` in lines (more
than 1 when wrapped text needs it). `RenderedSheet` holds the visible `columns` (with their width
in characters), the rows, and the sheet's warnings followed by the layout's.

### 3.6 Geometry in `excel` mode

Everything is in points, in **sheet coordinates**: x to the right of the first visible column's
left gridline, y down from the first visible row's top gridline.

- **A gridline is the first point of its column or row.** Column `p` covers `[x, x + width)`;
  its gridline is `[x, x + 1)` and its interior `[x + 1, x + width)`. x positions are running
  sums of `width_pt` (captured: the `layout` gridlines are 53 pt apart for 53 pt columns, and the
  runs of rows 130 to 136 put the 22 pt column N right after the 53 pt column M).
- **Text edges:** left text starts at `x + left_inset` (3 pt) plus one `digit_width` per indent
  level; right text ends `right_inset` (2 pt) before the next gridline; centered text sits at
  `x + 1 + ceil((interior − width) / 2)`, the odd point going to its right. The insets are
  xlformat's `ColumnMetrics`, provisional from print.
- **A fill covers its cell and its far gridlines**, `[x, x + width + 1) × [y, y + height + 1)`
  (captured from the PDF's fill rectangles).
- **Rows** are laid out by §4.7; **columns** by `file_column_metrics` with the Normal font.

| field | in `excel` mode | elsewhere |
|---|---|---|
| `RenderedColumn.x_pt`, `width_pt` | x of the column's gridline, its width in points (0 for a hidden column shown with `show_hidden`) | `None` |
| `RenderedColumn.width` | `None` | characters |
| `RenderedRow.y_pt`, `height_pt` | y of the row's top gridline, its height in points | `None` |
| `RenderedRow.height` | lines of text the row needs | the same |
| `RenderedRow.runs` | every text drawn from the row's cells (a merge's from its first visible row) | `()` |
| `RenderedRow.crossed` | per boundary between two visible cells, whether text or a merge crosses it (no vertical gridline there) | `()` |
| `RenderedRow.gaps` | `()` | characters mode's gaps |
| `RenderedCell.lines`, `source`, `font`, `color` | empty or `None`: the runs carry what is drawn | as §3.1 |

A **`TextRun`** is one line of a cell's text: `text`, `x` (the first glyph's advance), `y` (the
**baseline**), `width` (`xlformat.text_width`), `clip` (left, top, right, bottom of what shows;
text outside is cut through a glyph), `align` (the point a different font keeps fixed: `left`,
`center` or `right`), `source`, `font` and `color` (the section color, else the font color).
An angled run has an `angle` (degrees counterclockwise) and is turned about `(x, y)`, the start
of its baseline, with its rules; its `clip` stays upright (§4.17).

- **Baselines** (provisional, from print at Calibri 11; `layout` baselines): a line box is the
  font's row height (15 pt at Calibri 11) with its baseline 11 pt below its top. A bottom-aligned
  line sits 3 pt above the block's bottom gridline, a top-aligned one 11 pt below its top, a
  centered one at `floor((H − n × line) / 2) + 11`. Lines that do not fit the block start at the
  top, so a custom-height row shows its first line (captured, rows 161 to 168).
- **Glyphs** sit at `x` plus the running sum of `xlformat.glyph_widths`, in whole points, as Excel
  draws them: outside the competing spills, all 3,480 glyphs of 185 `layout` runs are within
  0.02 pt of the PDF's. (The other 35 runs have nothing to compare: lines a fixed row height hides,
  and clipped text the PDF interleaves with a neighbor's.) Layout decisions (fit, spill, placement)
  use `text_width` and match the capture.
- **Spill** (§4.1) is a run whose clip reaches past its own column; the origin cell's `spill`
  counts the columns. **Competing spills** are both recorded, with overlapping clips (decided),
  as Excel draws both; `characters` mode keeps first-wins.
- **Merges:** the anchor is formatted at the merged width, one column of the sum of the visible
  columns' points (`xlformat.column_metrics_from_points`), and its runs clip to the merged
  interior.
- **Wrapped, `justify` and `distributed` text** is broken at spaces by `text_width`, one run per
  line; `justify` spreads every line but the last, `distributed` every line, a lone word centered
  between the insets (row 162). `fill` repeats whole copies of the text that fit between the
  insets (row 159).
- **Text styles** (§4.13): a run's `scale` is its drawn size over `font.size` (superscript,
  subscript, shrink to fit), its `rules` its underlines and strikethrough, its `direction` `rtl`
  for a right-to-left paragraph. Vertically justified and distributed lines are spread.
- **Rich text** (§4.16): a cell with runs is measured run by run, and its `TextRun`s carry
  `spans`, each piece of the line in its own font with its width, size and shift.
- `scale` and `max_width` raise `ValueError` in `excel` mode.

`tests/test_captures.py` lays the `layout` workbook out in `excel` mode too: every row's crossed
and fill-covered boundaries against the missing gridlines, every captured run's start (exact) and
end (within a point), color and baseline.

### 3.7 Notices: what is not shown

Nothing a workbook uses is dropped silently. What a rendering leaves out or approximates is a
`Notice` (`_notices.py`): a dotted code registered in `FEATURES`, an outcome (`omitted` or
`approximated`), a reason (`not_implemented`, `not_representable`, `needs_dependency`, `option`,
`unsafe`, `file`), the sheet (or `None` for the workbook), the cells as at most 1,000 coalesced
ranges with an exact count, the output that decided it (or `None` for the model), and a context
mapping holding the message's values. Notices are aggregated: one per code, sheet, output, outcome,
reason and context, never one per cell.

Three producers, one for each stage:
1. **The reader** (`Sheet.notices`, `_detect.py`) reports what the file uses that the model does
   not carry. Styles are checked once per style, by each aspect in `_style/` (section colors
   against the workbook palette, rotation); the structure scan marks rich-text and phonetic
   cells, uncached formulas, outline levels outside 0 to 7, row and column styles, and links no
   output can follow (§4.10); the sheet's other children, its drawing, VML, comments and
   threaded comments parts, and the workbook's sheet kinds and print names give the rest.
   Notices on cells a layout hides are dropped there (`_notices.visible`).
2. **The layout** (`RenderedSheet.notices`) adds what every output approximates: fonts without
   glyph tables, refused formats, unshowable values,
   hidden rows and columns kept by `show_hidden`, and outline groups whose summary the range cuts
   off (§4.11).
3. **Each output**: the layout's census (`RenderedSheet.features`, `_census.py`) records the
   features the model carries that the laid-out cells use; each backend declares a `Support` per
   census code (`backends/_support.py`, the six tables side by side), and its notices are the
   census filtered through that table. A code a backend does not declare is reported as omitted,
   so a third-party backend without a table reports everything.

**A file is reported for what it uses, not what it contains.** A detector fires only when the
element states something Excel draws differently from its default:
- a style counts only through a cell that uses it, and only for what differs from the default
  (a border with a side style; a diagonal only with its flags; a row or column style only with a
  fill or a border);
- children that draw nothing on the sheet are never reported: `dimension`, `sheetCalcPr`,
  `sheetProtection`, `protectedRanges`, `scenarios`, `sortState`, `dataConsolidate`,
  `customSheetViews`, `phoneticPr`, `dataValidations`, `printOptions`, `pageMargins`, `pageSetup`,
  `headerFooter`, `customProperties`, `cellWatches`, `ignoredErrors`, `smartTags`,
  `webPublishItems`, `legacyDrawingHF`, `drawingHF`;
- window state that only chooses which part of the sheet is seen (zoom, `topLeftCell`, the
  selection, the active tab) is not reported: a page shows the whole range;
- `autoFilter` counts with a `ref`, `conditionalFormatting` with a `cfRule`, `tableParts` with a
  `tablePart`, page breaks only manual ones, outline symbols only while the file shows them,
  a pane when frozen or split, a tab color other than automatic; drawings, embedded objects and controls are not reported under
  `workbookPr showObjects="none"`;
- a linked picture (`r:link` without `r:embed`) is reported as `unsafe`: Excel fetches the linked
  file on open, and sheetpeek never fetches;
- sheet extension URIs are compared as Excel compares them, exactly. For conditional formatting
  this is captured: Excel reads the extension only under `{78C0D931-6437-407d-A8EE-F0AAD7539E65}`
  and ignores the upper-case form printed in [MS-XLSX], so that form is not reported. The other
  known URIs (sparklines, slicers, timelines; and validation, protected ranges, ignored errors and
  web extensions, which are not reported) are matched ignoring case until captured (unverified),
  so that none is missed;
- anything else among a sheet's children, or an unknown extension URI, is reported as
  `sheet.unknown_element` or `sheet.unknown_extension`;
- dialog sheets, which openpyxl loads as worksheets, and macro sheets are told apart by their
  relationship type from the workbook and reported.
- a chartsheet or dialog sheet is reported once for the workbook (`Book.notices`), and the
  workbook page reports the hidden sheets it leaves out (`sheet.hidden`, §4.15).

`tests/test_notices.py` holds a workbook using every registered feature and checks that each
output shows or reports each one, that every built-in output declares every census code, and that
plain workbooks from openpyxl, XlsxWriter and Excel give no notice.

**Delivery.** `sheetpeek.notices(source, to=...)` gives the three producers' notices without
rendering. `render` issues one `NotShownWarning` per call when there is something to report,
unless `on_notice` takes each notice or is `"ignore"`. The precedents: Python's `warnings`
categories and filters, docutils' report and halt levels, pandoc's "Not rendering" messages and
`--fail-if-warnings`, Sphinx's typed warnings and `suppress_warnings`, pydantic's error codes and
locations, and pandas' `on_bad_lines` callable. `sheetpeek view` writes one line per notice to
standard error after the output, the code first; `--notices`, `--suppress` and `--strict` (exit
status 3) control it. `HTMLBackend(notices=True)` lists them in a collapsed
`<details class="sp-notices">` after the sheet; `serve` turns it on.

## 4. Excel layout rules the model reproduces

Three sources of evidence, all Excel 16 for Mac with Calibri 11 (openpyxl's default workbook);
`CAPTURING.md` describes the methods:

- **Route A** (`capture/html_probe.py`, 2026-10-27): Excel's "save as web page" export plus
  `string value` of the cells. The export records the cell structure: merges and right spill as
  `colspan`, hidden rows and columns, each cell's displayed text. It does not record left spill or
  what is painted, and its clip points are simplified.
- **Checked by eye** (the route A workbook open in Excel, 2026-10-28).
- **Route B** (`capture/layout.py`, suite `layout`, 200 rows; later runs reproduced the earlier
  rows exactly): Excel's PDF export,
  read with pdfminer.six. Excel prints each gridline as a 1 pt rectangle and leaves a vertical
  one out wherever text crosses it, so the missing gridlines of a row show exactly which
  boundaries its text spans; the text runs give where the glyphs sit and their color. Clipping
  is done with clip paths that the text runs do not show, so clipped text was checked on the
  rendered page.

`tests/test_captures.py` lays the `layout` workbook out in `characters` mode, compares the
reach of every spill of digits with the capture, checks the other captured facts, and reports
the known deviations by count on every run. It compares `excel` mode with every row (§3.6).

**The PDF is a stepping stone:** it is Excel's print layout, and the screen is the truth
(`CAPTURING.md`, "The screen").

### 4.1 Text spills into empty neighbors

**Captured** (`layout`, rows 1 to 119; a default column is 53 pt, a digit 6 pt):
- Left-aligned text starts 3 pt after the column's left gridline and spills into the next column
  only when its end passes that column's gridline: 8 digits stay in A, 9 reach B, and 17 digits,
  ending exactly at C's gridline, still stop there. **A spill extends one whole column at a time,
  only as far as the text needs.**
- Right-aligned text ends 2 pt before the right gridline and spills left by the same rule, with a
  one-point asymmetry: text that would start exactly at the far edge of a gridline already counts
  as spilling (26 digits right-aligned reach three columns where 26 left-aligned reach two).
- Centered text is centered on the cell's interior; an odd remainder puts the extra point on the
  right (`LONG`, 215 pt, starts 107 pt left of the center and ends 108 pt right of it). **When it
  spills, it takes a column on both sides at once**: 9 to 26 digits span one column each way, 27
  or more two each way, even where the characters on one side would fit.
- Text of other glyphs follows the same rules by glyph width (rows 114 to 119).
- **A hidden column is crossed whatever it holds** (rows 127 to 129: a value in the hidden column
  changes nothing).
- **Two spills competing for one empty cell both take it**, and Excel draws the two texts over
  each other (rows 123 to 126, rendered page). Each still stops at a cell holding a value.
- **Clipping** (rows 120 to 122, rendered page): text blocked by a value is cut at the cell's
  edge, through a glyph. Left text keeps its start, right text its end, centered text its middle.
  There is no ellipsis.

Also captured by route A and checked by eye:
- A number, a formula returning `""`, a single space or a formula returning `0` in the neighbor
  stops the spill. A neighbor holding only a fill does not: the text is drawn over the neighbor's
  own fill, and the spilling cell's fill stays in its own column.
- Wrapped text does not spill; Excel grew the row to 96 pt. A cell in a hidden row still spills.

**How the `characters` layout follows this:** a spill takes whole columns as far as the text needs,
centered text takes a column on each side at once, and hidden columns are crossed. Excel's points
and characters differ in three of the 105 captured digit rows, listed in `tests/test_captures.py`:
a 53 pt column holds 8.83 digits, not 8, and the one-point asymmetry has no counterpart in
characters. Where two spills compete, characters cannot be drawn over each other, so the first
spill placed, left to right, takes the cell (an alternative is in §9). Text is clipped at a
character.

**Captured as well** (`layout` rows 156 to 175, rendered page):
- **A centered text blocked on one side** is clipped at that side; the other side takes only what
  its own text needs (12 centered digits blocked on the left take one column to the right).
- **`fill`, `justify`, `distributed` and wrapped cells do not spill.** `fill` repeats whole copies
  of the text (`ab` four times in a default column) and clips a text longer than the column.
  `justify`, `distributed` and wrapped text wrap at spaces; `distributed` also spreads the words
  across the width, which the `characters` layout does not (it places them as `left`).
- **A line break in a cell without wrapping is not drawn**: the lines are joined without a space
  (`first linesecond line`), and the joined text spills like any other.
- **Spilled text is drawn over a filled neighbor's fill, and a filled cell's fill stays in its
  own column** (rows 174 and 175; checked by eye earlier).
- A filled cell hides the gridlines on its edges, so a missing gridline next to a fill is not
  evidence of spill.

**Captured** (`layout` rows 180 to 200): **OOXML's `_xHHHH_` escapes are text** (§5.6), and:
- **A lone carriage return is not a line break.** It draws nothing, in wrapped text too, where
  `first_x000D_second` wraps as the one word `firstsecond`, broken by width (row 182). A line feed
  breaks a wrapped line, with or without a carriage return before it (rows 186, 188).
- **No control character is drawn**: ESC, tab and BEL between two letters leave `ab` (rows 198
  to 200). The layout removes them, and the C1 controls, from what a cell shows; `Cell.value` keeps
  them. They must not reach a terminal, where they would be commands.
- **Known deviation:** a line feed adds a little to the width that decides spill. Rows 180 and 184
  draw the same 49.79 pt text, and only the one holding a line feed spills; rows 195 and 197 show
  the addition is under 2 pt. The layout drops the line feed without measuring it.

**Captured** (`layout` rows 176 to 179): **merged text does not spill**, left, right or centered;
it is cut at the merge's edge. **A center-across run spills like centered text**, from the run's
area to both sides, a blocked side cut.

**Past the edge of the range**, spilled text runs on into empty columns that the layout adds at
the default width and marks `added` (`spill_edge="extend"`, the default; the alternatives are in
§9). Columns are added only past the used range, where the sheet is known to be empty; where an
explicit `ref` cuts through the data, the spill is clipped at the edge, because the cells beyond
it were not read.

Layout computes spill in the mode's units: points with xlformat's glyph measure in `excel` mode
(§3.6), characters in `characters` mode.

### 4.2 `####` for values that do not fit

**Captured** (route A, `string value`; `layout` rows 130 to 136):
- **Every error value overflows to `#` like a number** at every width from 1 to 8 (suite
  `error_width`, 56 cells, no deviation in `excel` mode). At width 3, `#DIV/0!`,
  `#N/A`, `TRUE` and `1234567` all show `##`, as xlformat formats an
  `ExcelError`.
- **A `#` run has the section's color** (`[Red]` gives red `###`, `[Blue]General` blue); an
  error's run has the font's color. xlformat returns the section color with an overflowed text,
  so the layout already draws it so.
- **A `#` run is centered in its cell**, whatever the value's alignment: `###` in the 22 pt
  column N starts at its interior plus 3 pt for numbers and the error alike (rows 130 to 136).
- **12345 under `@` overflows to `###`** in a 22 pt column, as xlformat renders it. 1234567890
  under `@` in a default column shows `1.23E+09`, left-aligned: General's reduction applies.
- A General number reduces before it overflows (`1.23E+12` at the default width), and
  `#,##0.00` of 123456789 shows nine `#`.

Unverified: cells whose font differs from the Normal font (bold, other size). xlformat
measures them in their own font against the Normal font's column (§6.4); the layout capture has
none.

### 4.3 Default alignment by value type

**Captured** (route A; `layout` rows 137 and 138): numbers right, text left, `TRUE` center,
`#DIV/0!` center, as xlformat aligns them.

**12345 stored as a number with the `@` format is left-aligned** (checked by eye; `layout` row 137
puts it 3 pt from the left gridline). xlformat's `FormattedCell.align` is `left` for a number
under any format with only a text section, such as `@` or `@" units"`, and the layout takes it
as it is (§6.4).

Explicit alignments override the default. `center_across` centers the text across the following
empty cells that share that alignment (route A; `layout` row 155 for a centered merge).

### 4.4 Merged cells

**Captured:**
- **Width-dependent formatting uses the merged width.** 123456789 under `#,##0.00` shows
  `#########` in one default column but `123,456,789.00` merged over A:C (route A).
- A merge across a hidden column (D13:F13 with E hidden) spans the visible columns (route A).
- **The anchor's text sits by its vertical alignment**: top, center and bottom as set, and **at the
  bottom by default** (`layout` rows 152 to 154). Horizontally it is placed across the whole merge
  (row 155).
- openpyxl drops covered cells' values on load: they come back as `MergedCell` with value `None`
  (measured).

The merged width for xlformat is the sum of the merge's visible columns, and the gaps between
them in `characters` mode. **Captured** (suite `merge_width`, 150 cells: five numbers alone and
merged over two to four columns, or over three with the middle one hidden, at six column
widths, read with `string value`): `excel` mode shows what Excel shows in every cell.

**Captured** (`layout` rows 172 and 173): **the whole merged area shows the anchor's fill, and a
covered cell's own fill is ignored**, also when the anchor has none. (openpyxl's normal mode drops
a covered cell's style on load; the read-only reader keeps it, and the layout ignores it.)

**Assumed:** a merge whose anchor lies outside the range read shows no text (there is nothing to
capture: the anchor is in the sheet, only not read). Wrapped text in a vertical merge flows over
the merged rows; if it needs more lines, the last row grows.

### 4.5 Hidden columns and rows

**Captured:** hidden columns take no width; spill and merges cross them, and a value in a hidden
column does not stop a spill (`layout` rows 127 to 129). Hidden rows are not shown.

**Captured** (`layout` rows 169 to 171; the width written into the XML, since openpyxl drops a
0): **a column of width 0 without the hidden flag behaves as hidden**. Its own text is not shown,
and spill crosses it whatever it holds.

### 4.6 Indent

**Captured** (`layout` rows 139 to 151): **one indent level is 6 pt, one digit's width**, for left
and right alignment alike, and indented text spills as usual. The `characters` layout uses one
character per level.

### 4.7 Wrapped text and row heights

**Captured** (`layout` rows 161 to 168): wrapped, `justify` and `distributed` text is broken at
spaces, and an explicit line break starts a new line. **A row with a custom height
(`customHeight`) shows only the lines that fit**; the capture's rows are 15 pt, so only the first
line shows.

**Known deviation:** the `characters` layout ignores row heights and grows every row to the most
lines any cell needs, as Excel's row autofit would. Following a custom height means turning points
into lines, a choice listed in §9; it waits for `excel` mode, which has the points.

**Captured** (suite `row_height`, `CAPTURING.md`): **a row without `ht` takes the Normal font's
height, and `defaultRowHeight` counts only with `customHeight="1"`.** `excel` mode lays a row out
at, in order:

1. its `ht` when the file stores one (marked provisional until the screen capture: rows 15 pt in
   the file print 14 pt apart);
2. `0` when `<sheetFormatPr zeroHeight="1">` and the row has no `<row>` element;
3. `defaultRowHeight` when `<sheetFormatPr customHeight="1">` and `zeroHeight` is not set;
4. 16 pt when the file has no `<sheetFormatPr>` and the row has no `<row>` element;
5. otherwise the Normal font's height from the captured table (Calibri 12, 14, 15, 16, 19, 21 pt
   at 9 to 16; Aptos Narrow 13, 14, 15, 16, 19, 22), plus 1 pt each for `thickTop` and
   `thickBottom` when the row has no `<row>` element.

**Rows are fitted to their cells** when Excel opens a file (captured, suite `row_height`): a
row with a `<row>` element and no `ht`, in a sheet without `customHeight`, takes the largest of
those heights and of each cell font's height from the table, merged cells included, and of its
wrapped text, 16 pt per line at Calibri 11 (the line height plus 1; other sizes unverified),
merged cells excluded. `thickTop` or `thickBottom` together with `customHeight` was not
captured. A font not in the table uses Calibri's height at the nearest measured size, as xlformat scales
widths; that, and sizes between the measured ones, are **unverified**.

Excel ignores a stored `defaultRowHeight` without `customHeight` (15, 20 and 30 captured): with
openpyxl's 15, a Calibri 14 workbook shows 19 pt rows and a Calibri 9 one 12 pt rows, so that
default would be right only for Calibri or Aptos Narrow 11. The structure scan therefore records
which rows have a `<row>` element, whether the sheet has `<sheetFormatPr>`, and its
`customHeight`, `zeroHeight`, `thickTop` and `thickBottom` attributes.

In `characters` mode the break uses Python's `textwrap` (long words broken), by characters; `excel`
mode breaks by `text_width` (§3.6).

Superscript, subscript and accounting underlines grow a fitted row, and shrunk text fits it at its
unshrunk size (§4.13).

### 4.8 How Excel saves formula results

**Captured** (`capture/excel_saved.py`; the saved file is `tests/files/excel_formulas.xlsx`):
Excel writes a formula returning `""` as `<c t="str"><f>""</f><v/></c>`, `t="str"` with an empty
value element. The structure scan treats that form as the empty string, which stops a spill;
openpyxl alone reads it as an empty cell. A formula without a cached value (as openpyxl writes it)
has no `t`. The same file confirms Excel stores `=DATE(1900,2,29)` as serial 60, which the reader
keeps (§5.3).

### 4.9 Captures

Captured so far: `layout` (§3.6, §4.1 to §4.7), `theme_colors` (§3.3), `merge_width` and `error_width`
(§4.2, §4.4, no deviation), `row_height` (§4.7) and `package` (§5.5). The outline rules (§4.11)
rest on captures on Excel 16.114 that are not yet in `tests/captures/`. What waits for the screen capture: the gridline and heading colors, the heading sizes, the text insets
and baselines, which come from print, and the 15 pt rows that print 14 pt apart.

### 4.10 Hyperlinks

**Unverified until recaptured on Microsoft 365.** The rules below were captured once, on Excel 16
for Mac with Office 2021, before the switch to Microsoft 365; they are marked unverified until
that capture is repeated.

- **A link has no look of its own.** A linked cell in the Normal style is drawn as plain text;
  one in the Hyperlink or Followed Hyperlink cell style takes that style's font color (the file's
  theme `hlink` or `folHlink`) and single underline. A cell in the Hyperlink style without a
  link looks the same. So the cell's format draws a link, and every output adds nothing for it.
- **The text is the cell's value.** `display` is not drawn, and an empty linked cell shows
  nothing; it gets no anchor and is reported (`link.empty_cell`). A merge's covered cells are not.
- **`HYPERLINK()` from a file** is drawn as its cached value, in the cell's own format: Excel adds
  no style when it loads a file.

What sheetpeek decides (not Excel's behavior):
- **Reading.** A `<hyperlink>` gives its `ref` range, the target of its relationship, `location`,
  `tooltip` and `display`; a relationship that is not `External` points into the package and is
  refused (`link.scheme`). A `HYPERLINK()` cell is linked only when the formula is one call whose
  first argument is a string literal (doubled quotes undone); a computed target, a shared or an
  array formula is reported (`link.computed`). An element and a formula on one cell: the element
  wins and the conflict is reported (`link.conflict`); which one Excel follows is unverified.
- **Schemes:** an allow-list, `http`, `https` and `mailto`, compared lower-cased after what a
  browser drops (C0 controls and spaces at the ends, tabs and line breaks anywhere, so
  `java\tscript:` is `javascript:`), with `urllib.parse.urlsplit`. Anything else (`javascript:`,
  `data:`, `vbscript:`, `file:`, UNC and relative paths, unparseable text) is drawn as plain text
  and reported once by the reader (`link.scheme`, `unsafe`), never per output, and the refused
  target is never written into any output. Controls left in an allowed URL are percent-encoded,
  so none reaches a terminal.
- **Places in the workbook.** `location` names a cell (`A1`, `Sheet!A1`, `'My sheet'!B2:C3`,
  first cell taken; sheet names compared ignoring case) or a defined name. A reference without a
  sheet is taken as the link's own sheet (unverified). A link to a cell that is laid out goes
  there in HTML (`link.internal`); a link to another sheet, a defined name or a cell not laid out
  is an anchor that goes nowhere (`link.elsewhere`, approximated). Defined names are not resolved.
- **The census** (§3.7) counts `link` (an allowed URL), `link.internal` and `link.elsewhere` on
  cells that show text; each output declares what it does with them (`backends/_support.py`).
- **On the workbook page** (§4.15) a link to a cell of another sheet on the page goes there, and
  counts as `link.internal`.

### 4.11 Outlines

**Captures** (Excel 16 for Mac, Office Home & Business 2021, dark appearance; one run
of a sheet with rows 4 to 17 grouped up to level 3 and columns D to J up to level 2, at zoom 100
and 200, and one with summary columns on the left). Every rule here stays **unverified until it
is recaptured after the move to Microsoft 365**; the tests that rest on them say so.

**Groups from levels** (`_pass_outline.py`). For each level `L` from 1 to the deepest, a group is
a maximal run of rows whose `outline_level` is at least `L`; its summary row is the row after the
run (`summaryBelow`, the default) or the row before it. Its button is − while any of its detail
rows is visible, and + when all are hidden or the summary row has `collapsed` (captured: rows 11
and 17 show only a +, rows 7, 8 and 12 a −). Hidden detail rows draw nothing. Each visible
detail row gets a dot in the lane after its own level, unless a button is there (captured: rows 4,
5 and 6). Columns follow the same rule with `summaryRight` (captured mirrored for summaries on the
left: the button on the summary column, the tick at the far end of the last detail column).
Summary rows above are taken as mirrored the same way (unverified). The result is
`RenderedSheet.outline`, a `RenderedOutline` of `OutlineMark`s (index, lane, kind) for the rows
and columns laid out.

**What the reader takes from the file.** `outlineLevel` and `collapsed` of `<row>` and `<col>`;
`summaryBelow` and `summaryRight` of `<outlinePr>`; `showOutlineSymbols` of the first
`<sheetView>`, which wins over `<outlinePr>`'s. **The depth is the deepest level found in the
whole sheet** (captured: without `outlineLevelRow` Excel draws the same gutter as with it), so the
declared `outlineLevelRow` and `outlineLevelCol` are not read. Office allows levels 0 to 7; a
deeper level is read as 7, and a level that is not a whole number from 0 up as 0, reported as
`outline.level` (what Excel does with them is not captured). With `showOutlineSymbols="0"` nothing
is drawn and nothing is reported (taken as given, unverified).

**A range that cuts a group** still draws the bars of its rows inside the range. The model keeps
the levels of the row and column just outside the range (`Outline.rows_beyond`,
`columns_beyond`), so a run that continues past the edge draws no tick there, and a summary just
inside the edge gets its button. A group whose summary is outside the range draws no button,
reported as `outline.cut` (decided).

**Geometry** (captured; drawn by the HTML backend, `backends/html/_outline.py`). **Every size is
screen-fixed**: identical at zoom 100 and 200, so one screen point is one CSS pixel at any page
scale (decided), like the gridlines.
- The row gutter is left of the row headings, `2 + 12 × (depth + 1) + 3` px wide (53 at depth 3);
  the column gutter is above the column headings, by the same rule (41 at depth 2). Lane `n`
  starts at `2 + 12 × (n − 1)`.
- A bar is 2 px wide, 3 px into its lane, and starts 2 px below the first detail row's top (or
  past the first detail column's left edge), at any row height, with a 6 × 2 px tick toward the
  headings; it runs into the summary's button. With summaries above or left, it ends 2 px short
  of the last detail row's far edge, where the tick is.
- − and + buttons are 12 × 12 px, filling the lane, rounded, white, with a 1 px border; their
  strokes are 6 px long and 1 px thick. **A button is centered on its row**, its top at
  `1 + ⌈(h − 1 − 12) / 2⌉` for a row `h` px high (2 in 15, 10 in 30, 25 in 60), so it does not
  follow the row height.
- A dot is 2 × 2 px, centered on its row at the bar's x.
- The level buttons are 12 × 12 px, one per lane: the row levels across the column-heading band,
  centered in its height; the column levels down the row-heading column, centered in its width.
- Colors in the dark appearance: gutter 85, bars 153, button border 51 (gray levels). The light
  appearance is not captured.

Not captured: the light appearance's colors, the exact end of the tick on the column side, the
gutter with the headings off, and summaries above.

**Other outputs.** `rich` and plain text draw a gutter of characters, one per lane before each
row and one line per lane above the columns: `│` and `─` for bars, `┌`, `└` or `┐` for a bar's
tick end, `-` and `+` for the buttons, `·` for dots; no level buttons (approximated). Their
`buttons="none"` leaves it out, as in HTML. Markdown, CSV and TSV report the outline as not
shown.

### 4.12 Fills and borders

**Unverified until recaptured on Microsoft 365.** Every rule below was captured once, on the
screen of Excel 16 for Mac with Office 2021 at zoom 100 and 200, before the switch to Microsoft
365; the tests that rest on them say "capture on Excel 16.114". The model is in
`_fill_border_model.py`, the reading in `_style/fill.py` and `_style/border.py`, the layout in
`_pass_borders.py`, the drawing in `backends/html/_patterns.py` and `backends/html/_borders.py`.

**Pattern fills.** `CellStyle.fill_spec` is a `Pattern` (`solid` or one of 17 patterns, `fg`,
`bg`) or a `Gradient`; `CellStyle.fill` stays one representative color for outputs that show
one.
- Each pattern is a bitmap tile in two colors only: `fgColor` draws the set pixels, `bgColor` the
  rest (`TILES`, one tile pixel a whole screen point; `gray0625` is 8 × 4 with two set pixels).
  The tiles are fixed in screen points, the same at zoom 100 and 200, and anchored at the sheet's
  origin, not at each cell (Excel 16.115, screen suite `fills`; 16.114 drew a tile pixel as one
  device pixel, half a screen point).
- Colors: `solid` draws `fgColor`; a missing `fgColor` is black, so `solid` with only `bgColor`
  draws black; a pattern without colors, with automatic ones, or with indexed 64 and 65, is black
  on white; one color given, the other takes its default (fg black, bg white). A `solid` fill in
  indexed 64 is no fill; a fill index past the palette (81) is black; `patternType="none"` or none
  at all draws nothing, whatever its colors. The alpha byte is ignored. openpyxl reads a missing
  color as `00000000`, so a background of `00000000` is read as missing (white).
- **The representative color** of a pattern is the blend of its two colors by the foreground's
  share of the tile (6.25 % for `gray0625`, 43.75 % for `lightGrid`); of a gradient, its color at
  the midpoint.

**Gradient fills.** `linear` runs along `degree`, clockwise from left to right, **taken in the
unit cell**: on a `w × h` box the drawn angle is `90° + atan2(w sin d, h cos d)` in CSS terms, and
the corners take the end stops, as CSS `linear-gradient` normalizes them; colors interpolate
linearly in sRGB. Past the last stop the last color holds. One case had its first stop at 0.25:
before it the color ran from the last stop's color at 0 (taken as the rule, unverified). `path`
draws rectangular bands from the rectangle `left`, `right`, `top`, `bottom` (in the first stop's
color) out to the cell's edges; all four 0 put the first stop at the top-left corner. A degree
that is not finite is 0, rectangle values are clamped to 0 to 1, stops are sorted by position,
and a gradient without stops is no fill (decided, not captured). Duplicate stop positions make
openpyxl refuse the whole stylesheet, a `PackageError` until the reader move.

**Fills against gridlines.** A pattern or gradient covers its cell and the gridlines after its
right and bottom edges, like a solid fill (§3.6); the cell to the right or below paints over that
gridline with its own fill. A pattern shows its pattern on the gridlines it covers.

**Borders.** `CellStyle.borders` is a `Borders` of four `Side`s (13 styles, a color or `None`
for automatic, which is black), and a diagonal with its `diagonal_up` and `diagonal_down` flags.
`start` and `end` are read as `left` and `right`. A diagonal is read only with a style and a
direction; Excel draws nothing otherwise.
- **Lines are fixed in screen points**, the same at every zoom (CSS px in HTML): thin 1 on the
  gridline; medium 2 from 1 before it; thick 3 from 1 before it; double 1, a white 1, 1, from 1
  before it; hair 1 on 1 off; dotted 2 on 2 off; dashed 3 on 1 off, all in whole screen points
  with no stagger between device rows (Excel 16.115, screen suite `borders`; 16.114 drew hair as a
  device-pixel checkerboard, dotted 1/1 staggered and dashed 1.5/0.5); dashDot 9, 3, 3, 3; dashDotDot 9, 3, 3, 3, 3, 3;
  the medium dashes the same at 2 wide; slantDashDot 2 wide, 11, 1, 5, 1 with slanted
  antialiased gaps. **Every gap is white**, also over a fill: no gridline shows under any border.
  Sides of one cell meet square.
- **Conflicts:** where two cells' sides meet, one is drawn whole, in its own style and color,
  never both and never mixed. The stronger style wins: **double > thick > medium > mediumDashed >
  mediumDashDot > slantDashDot > mediumDashDotDot > thin > dashed > dotted > dashDot > dashDotDot >
  hair** (73 pairs, in both orders and orientations, no contradiction). Of one style, **the darker
  color wins, by Rec. 601 luma** (11 pairs, both orders; Rec. 709 fits them as well), never the
  position. Two colors of equal luma: the larger `RRGGBB` (unverified). Colors are compared after
  resolution to RGB (unverified for theme and indexed colors).
- **Merges:** borders stay per cell. The merge's outline is the outer sides of its edge cells,
  each on its own segment, the anchor's only on its own; sides inside the merge are not drawn;
  a neighbor's side competes with each segment by the rule above. A merge's diagonals are its
  anchor's, across the whole merge; a covered cell's diagonal is not drawn.
- **Spilled text removes every vertical border it crosses**, over the whole row height, on
  either cell, while the horizontal sides of the cells it crosses stay. Where the spill stops at
  a value, that cell's left side is drawn. So the borders pass runs after the cells and reads
  `RenderedRow.crossed` (§3.6).
- **Hidden rows and columns** contribute no sides; the visible neighbors' facing sides meet on
  the one edge and compete.
- Diagonals run corner to corner of the cell or merge, antialiased; hair diagonals are drawn as
  thin, dashed ones as short blocks; text and the line both show.

**The layout** (`excel` mode only) resolves one side per edge of the laid-out grid into
`RenderedSheet.borders`, a `RenderedBorders` of horizontal and vertical lines and the
diagonals; no output compares two cells' sides. A range read cuts off the sides of the cells
beyond it (not read).

**HTML** draws two layers of absolutely positioned boxes before the table's text, in sheet
coordinates: the fill layer (`sp-fills`: SVG tiles drawn `pixelated` and anchored at the sheet's
origin, CSS linear gradients, SVG path gradients in four bands) and the border layer
(`sp-borders`: one box per run of equal sides along a line, reaching over the bands of the lines
across its ends; dashes as SVG tiles on white; diagonals as SVG lines). Both stay in CSS px at any
page scale. Tiles and dashes are in whole screen points, one to a CSS px, so a 1× display draws
them exactly too. Where Excel starts a dash period is not captured; a run starts its own.
**slantDashDot is drawn as mediumDashDot with square gaps** and reported
(`border.slant_dash_dot`).

Not captured: a dash's phase, the light and dark appearances' effect (none expected), double
lines meeting at a corner (drawn crossing), what Excel does with a gradient without stops, and
the zoom-200 oddities (a 3 × 1 merge's diagonal stopping at its middle; two linear gradients
ending 1 pt short of the right gridline).

**Other outputs.** `rich` shows a pattern as its blended color and a gradient as its midpoint
color (approximated) and no borders (`not_implemented`: an option could later draw box lines);
plain text, Markdown, CSV and TSV show neither fills nor borders.

### 4.13 Text styles

**Captures** (Excel 16 for Mac, Office Home & Business 2021, Calibri 11 at zoom 100
and 200, measured on the screen image at 2 image pixels a screen point). Every rule here stays
**unverified until it is recaptured after the move to Microsoft 365**; the tests that rest on
them (`tests/test_text_styles.py`) say so. Sizes other than those captured are scaled from them,
unverified. The rules live in `_text_styles.py` and `_excel.py`.

**Superscript and subscript** (`Font.vert_align`). The whole cell's text is drawn at 2/3 of its
size (cap heights 0.64 to 0.68 of the plain text's at 11 and 24 pt), its baseline raised by
`round(0.48 × size)` or lowered by `round(0.18 × size)` from the cell's baseline (5 and 2 pt at
11 pt, 12 and 4 pt at 24 pt). It is measured at the drawn size, for wrapping, `####` and spill
as for its runs (unverified for `####`). A fitted row grows by 2 pt at 11 pt (17 pt) and 6 pt at
24 pt (37 pt against 31), linearly between and beyond and never below 0 (unverified); a row
holding both kinds grows 1 pt more at 11 pt (18 pt), half the one-kind growth at other sizes
(unverified). `TextRun.scale` carries the drawn size over `font.size`.

**Underlines and strikethrough** are lines **1 screen pt thick at any zoom**, as gridlines and
outline symbols are: one CSS pixel at any page scale (decided), at whole-pixel positions. The
model gives each run its `TextRule`s (`TextRun.rules`): where a line starts and ends and the top
of the first line, in points, offset from the run's baseline by the 11 pt captures scaled with
the drawn size (unverified at other sizes). A second line is 2 screen pt below the first.
- Single: 1 pt below the baseline; double: 0 pt, and a second line. Both span the text.
- Single accounting: 3 pt below the baseline; double accounting a second line. **On text they
  span the cell**, from 3 pt after its left gridline to 2 pt before the next one, even under
  two letters in a 213 pt cell; a merge spans the merge, wrapped text gets one line per line,
  distributed words one line under the whole line (both unverified). **On a number, boolean or
  error they span the value**, like a single underline (Excel for Windows is said to span
  numbers too; unverified). They add 3 pt under the text: a fitted row is 18 pt, not 15, and
  a bottom-aligned baseline sits 6 pt above the cell's bottom, also in a row of fixed height
  (unverified).
- Strikethrough: 3 pt above the baseline, the text's width.

**Outline, shadow, condense and extend** are drawn as plain text by Excel for Mac (the same ink
at both zooms): they are not read, and nothing is reported.

**Shrink to fit** (`CellStyle.shrink`). Text too wide for its cell is drawn at the largest `k/7`
of its size, `k` from 1 to 7, that fits between the insets, less the indent: a fixed screen step
(1/14 at zoom 200), so the layout, made at zoom 100, uses sevenths. 18 of the 19 captured lengths
(4 to 22 Hs in a default column) take Excel's step; ten Hs take 4/7 where Excel draws 5/7 (their
width at 5/7 measures 50 pt in a 48 pt budget), a known deviation. Captured at Calibri 11 only;
the step at other sizes is unverified, and when even 1/7 does not fit the text is clipped
(decided). **It replaces `####`**: a value whose format overflows shows its full text, shrunk
(`RenderedCell.text` is that text). **Wrap overrides it**, as do horizontal and vertical justify
and distributed, which wrap. **Shrunk text never spills**, in every mode. **Rows fit the unshrunk
size** (captured at 20 pt). A merge with the flag on its anchor is not shrunk, clipped at the
merge's edge (captured once, doubtful). The baseline stays where unshrunk text would sit
(unverified).

**Vertical justify and distributed** wrap text whose wrap flag is off (captured: the file still
says it does not wrap). One line sits at the top (justify) or the center (distributed). Several
lines spread from the top position to the bottom one with an even pitch: the first baseline 11
pt below the top, the last 3 pt above the bottom at Calibri 11. Lines that do not fit start at the
top, as other wrapped text. Characters mode spreads them in whole lines, and does what Excel does
with one line.

**Horizontal distributed** spreads the words between the insets, a single word centered
(§3.6), and wraps. **Indent moves both ends in** (captured: 12 pt for indent 2). `justifyLastLine`
is read (`CellStyle.justify_last_line`) and changes nothing: distributed already spreads the last
line (captured). The row fitted for distributed text (48 pt for two drawn lines) is not
reproduced: fitting counts wrapped cells only (open).

**Indent** is 6 pt a level (§4.6) on the side the alignment resolves to, `general` included:
right-aligned text and numbers end `2 + 6n` pt before the next gridline (captured as ink 3, 9
and 21 pt for indents 0, 1 and 3, and 15 pt for a `general` number with indent 2).

**Reading order** (`CellStyle.reading_order`: 0 context, 1 left to right, 2 right to left).
**`general` text whose first strong character is right-to-left aligns right whatever the
order**, in every mode (captured: Hebrew with order 0 and with order 1). Its runs are a
right-to-left paragraph (`TextRun.direction`, `rtl`), unless the order is 1, which makes the
paragraph left to right: the "!" after Hebrew moves to the right end (captured). Order 2 on Latin
text changed nothing (captured once, doubtful), so it is not applied. A file's value other than
0, 1 or 2 is read as 0.

**Drawn in HTML** (`backends/html/_text.py`): a run's `font` size is `font.size × scale`; its
rules are `<i class="sp-rule">` elements beside it in its clip window, 1 px high, a double
underline one `sp-double` element 3 px high with a top and a bottom border; a right-to-left run
has `direction: rtl`.

**Other outputs** (decided; each declared in `backends/_support.py`):

| feature | `rich` | text | Markdown | CSV, TSV |
|---|---|---|---|---|
| superscript, subscript | approximated: full size on the baseline | omitted, as other font styles | approximated: plain text | omitted |
| accounting underlines | approximated: the terminal's underline | omitted | omitted | omitted |
| shrink to fit | approximated: text clipped at full size, a number too wide shows `#` | the same | shown: the full text, nothing to shrink | shown |
| vertical justify, distributed | spread in whole lines (as `align.vertical`) | the same | omitted (as `align.vertical`) | omitted |
| reading order | approximated: the terminal's own text direction | the same | omitted | omitted |

### 4.14 Theme fonts, palette, grid color and headings

**Captures** (Excel 16.114.4 for Mac, Office Home & Business 2021, dark appearance,
zoom 100 and 200), to be recaptured on Microsoft 365, like §4.13. The tests that rest on them
(`tests/test_theme.py`) say so.

**Theme fonts.** A font with `scheme="minor"` or `"major"` is drawn in the theme's Latin typeface
(`a:minorFont/a:latin`, `a:majorFont/a:latin`) whatever name it stores: Calibri-minor and
Times-New-Roman-minor both draw the theme's Courier New, at its width, and Calibri-major draws
Georgia; a font without a scheme keeps its own name. The reader resolves this when it reads the
font (`_style/font.py`): `Font.name` is the typeface Excel draws, `Font.scheme` and
`Font.stored_name` keep what the file says. **The Normal font follows the theme too**, so the
default column takes the theme font's digits (captured: 61 pt for Courier New, against 53 pt for
Calibri 11), and so does every width in `excel` mode. A typeface without xlformat glyph tables is
measured as Calibri at the same size, with `font.no_metrics` (§6.2), so such a theme's widths are
approximate. Without a theme, or with an empty typeface, the stored name is drawn (unverified).
- **Other scripts:** East Asian text in a minor font is drawn in the theme's script font (`Jpan`,
  captured), not the system fallback. sheetpeek reads the Latin typeface only, so a page draws
  such text in the browser's fallback for that typeface. A `@font-face` family per theme
  collection with a `unicode-range` per script would draw it exactly (measured equal to a
  fallback list in WebKit); it needs the theme's script fonts in the model and a table from ISO
  15924 codes to ranges, and is not built.

**The workbook palette** (`_colors.workbook_palette`). `styles.xml`'s `indexedColors` replaces
the default palette entry by entry: indices 0 to 7 no longer repeat 8 to 15 (captured: a written
palette that recolors index 10 leaves fill index 2 red). Fill and font colors follow it (captured
for indices 10 and 17). Excel caps the list at 64 entries; a shorter list keeps the default for
the rest, and an entry that is not a color reads as the default's (both unverified). A
`<colors>` without `indexedColors` keeps the default palette (openpyxl reads it as empty).
Automatic colors (captured): a fill in indexed 64 is no fill, a fill index past the palette (81)
is black, and a font in `auto` or indexed 64 is the automatic black.

**Number-format colors.** `[ColorN]` draws the workbook palette's index `N + 7`; the eight named
colors (`[Red]`, `[Blue]`) draw their default colors whatever the palette (captured; Excel's
`display format` reports palette indices 2 and 4 for `[Red]` and `[Blue]` instead, so the screen
was taken as the truth). The layout draws section colors from xlformat's default palette, so a
`[ColorN]` that the workbook palette recolors is reported as `format.palette` (approximated), for
cells with a value, by the number format's style aspect (`_style/number_format.py`). Quoted text
and escaped characters are not section colors.

**The grid color** (`Sheet.grid_color`). `<sheetView defaultGridColor="0" colorId>` draws the
gridlines in the workbook palette's entry `colorId` (captured: teal for index 10 of a palette
that makes it teal); 64, the default, is the automatic gray (captured). An index past the palette
or a malformed one is automatic (unverified). The colored gridline is **1 device pixel wide**,
half the default line's 2, at any zoom; a fill covers it as it covers the gray one (captured); the
headings keep their colors (captured). The first `<sheetView>` is read, as for gridlines.
- **HTML:** a generated class `sp-g-RRGGBB` on the wrapper sets `--sheetpeek-grid` and draws the
  cells' gridlines 0.5 CSS px wide: one device pixel on a 2× display, as Excel draws it; a 1×
  display rounds it (decided). Measured in WebKit's snapshots: at 2× the line is one device pixel
  in the full color, against the gray line's two; at 1× both are one pixel; over 30 columns every
  line stays within one device pixel of the gray grid's, with no drift. WebKit draws it on the
  first of the gray line's two pixels, Excel on the second, at zoom 100. A sheet that hides its
  gridlines draws no grid color.

**Headings.** Excel's heading bands are sheet units: they scale with the zoom, so the page draws
them in points like the cells (`backends/html/_headings.py`).
- At Calibri 11 the column band is 21 pt high and the row band `16 + 3 × digits` pt wide, the
  digits of the last row in the window (captured with 2, 3 and 4 digits: 22, 25 and 28 pt). The
  page uses the last row it draws (decided: a page shows every row).
- The bands grow with the Normal font: at Calibri 20 the column band is 25 pt and the row band 36
  pt with 2 digits (captured). Other sizes follow the line through the two (unverified), which
  stays positive at every size Excel accepts (1 to 409 pt).
- Colors: the dark appearance's are captured (paper 22, text 255, separators 86); the light
  appearance's were measured on Microsoft 365 (paper 247, separators 224,223,221, text the
  darkest pixel of its glyphs, 94) and stay provisional until the light appearance is captured
  with the rest. The size of the headings' text is not captured.
- The headings options do not follow the file's `showRowColHeaders` (`headings.hidden`, §3.7).

**Other outputs** (decided; each declared in `backends/_support.py`):

| feature | `rich` | text | Markdown | CSV, TSV |
|---|---|---|---|---|
| theme fonts | as any font name: omitted | omitted | omitted | omitted |
| grid color | approximated: the column separators in the grid color | omitted | omitted | omitted |
| `[ColorN]` from the workbook palette | reported by the reader (`format.palette`), as for HTML | the same | the same | the same |

### 4.15 Sheets and the workbook page

**Captures** (Excel 16.114.4 for Mac, Office 2021, dark appearance, zoom 100, measured
on the screen image): which tabs the strip shows and how a tab color is drawn. They stay
**unverified until recaptured after the move to Microsoft 365**; the tests that rest on them
(`tests/test_workbook.py`) say so. The rules live in `_sheets.py` and
`backends/html/_workbook.py`.

**The sheet list** (`Book.sheets`, a `SheetInfo` per sheet, in workbook order):
- **Kind** from the relationship type of the workbook's `<sheet>`, never from openpyxl, which
  branches only on chartsheets and loads a dialog sheet as an empty worksheet: `worksheet`,
  `chartsheet`, `dialogsheet`, or `macrosheet` (Microsoft's `xlMacrosheet` and
  `xlIntlMacrosheet`). A type that names none is a worksheet.
- **State** from `sheet@state`: `visible`, `hidden`, `very_hidden`; anything else is visible.
- **Tab color** from `sheetPr/tabColor`, read from the head of the sheet's part (openpyxl's
  read-only sheets do not parse `sheetPr`); the parse stops at the end of `sheetPr` or at the
  first other child, so no cell is read for it. Resolved like any color (§3.3); `auto` and no
  color are `None`, since `auto` draws an uncolored tab (captured). When several of `indexed`,
  `theme`, `auto` and `rgb` are given, they are taken in openpyxl's order, as every other color
  is (which one Excel takes is unverified). An index or theme number is an `unsignedInt`
  (`[0-9]{1,9}`, else none); a tint that is not a finite number is 0, and one past ±1 is
  clamped.
- Sheets whose part is missing are left out, as openpyxl leaves them out, so `Book.sheets` and
  `Book.sheet_names` agree; a second sheet of one name is left out too.

**Which sheets read.** Worksheets and macro sheets; a macro sheet's cells are cells, and it is
reported (`sheet.macrosheet`, approximated). A chartsheet or dialog sheet raises
`NotAWorksheetError` (a `ValueError`) and is reported at workbook level in `Book.notices`
(`sheet.chartsheet`, `sheet.dialogsheet`). Hidden and very hidden sheets read by name.
**The default sheet** is the active one (the first `workbookView`'s `activeTab`, an index into
the workbook's own `<sheet>` list) when it is a visible worksheet or macro sheet; else the first
visible one; else the first readable one, hidden or not; with none, `NotAWorksheetError`.

**The tab strip** (captured): Excel shows the visible sheets only, in file order; hidden and very
hidden sheets have no tab. An inactive colored tab is filled solid with the color, with white
text; the active tab is a vertical gradient from near white to a light tint of the color (two
samples fit 10 % and 40 % of the color mixed with white, an inference), with bold text in
Excel's own green (`#278452`), not the file's color. Each tab is 72 × 20 pt, 7 pt apart.
`showSheetTabs`, `firstSheet` and a tab's platform shape are window chrome with no counterpart
on a page and are not read.

**The workbook page** (`render_workbook`, `sheetpeek view --all-sheets`): one HTML page, an
option; one sheet per page stays the default everywhere.
- A `<section>` per tab, then a `<nav>` of tabs, one per visible sheet in workbook order (with
  `hidden_sheets=True` hidden and very hidden ones too, in italics, titled with their state). A
  worksheet's section is the `html` backend's sheet; a chartsheet's or dialog sheet's says that
  it is not shown. The strip is below the sheets, as in Excel's window; `tabs="top"` draws it
  above them.
- **Ids by position, not by name:** a sheet's section is `sp-<n>`, its position in the workbook
  from 1, and a cell a link goes to is `sp-<n>-<cell>`, so ids are unique on the page and no
  file string reaches an id or an `href`. A one-sheet page keeps `sp-<cell>`.
- **Without script.** A tab is a link to its section. In a browser with CSS's `:has()`, a
  section shows when it is the URL's target or holds it, and the active sheet's when none is;
  the shown section's tab takes the active style. Elsewhere every section shows, one after
  another, and the active sheet's tab is the active one. The page's first section is the active
  sheet's (`activeTab`) when it is a visible sheet on the page, else the first tab's.
- **Links across sheets work:** a `location` naming a cell of another sheet on the page, drawn
  there (not in a hidden row or column, inside the used range), goes to `#sp-<n>-<cell>` (a
  merge's first visible cell), and the census counts it as `link.internal`. A link to a sheet
  not on the page, a cell not drawn or a defined name stays `link.elsewhere`.
- **Generated classes** (`sp-f-`, `sp-b-`) carry a per-sheet tag (`sp-f-a-`, `sp-f-b-`, ...), so
  two sheets' rules do not collide in the one stylesheet.
- **Notices:** each sheet's as for `render`, without `sheet.tab_color` (drawn); then the
  workbook's: the chartsheets and dialog sheets on the page, and `sheet.hidden` (omitted, option)
  for the hidden sheets left out, with how many are very hidden. The default is one
  `NotShownWarning` per sheet with something to report and one for the workbook.
- **Cost:** every sheet is read and laid out before the page is written, since a link gives an
  id to a cell of another sheet; `max_cells` bounds the page's cells, every sheet's together.
- **Options stay orthogonal:** `hidden_sheets` (`--hidden-sheets`) is about sheets, and
  `show_hidden` (`--hidden`) about rows and columns, as for `Sheet.layout`.

**Other outputs** (decided): text, `rich`, Markdown, CSV and TSV have no workbook page and draw
one sheet; `--all-sheets` refuses any `-t` other than `html`. In a one-sheet render of any
output a tab color is `sheet.tab_color`, omitted with reason `option`: only the workbook page
draws tabs. `sheetpeek view --list` prints each sheet's name, and its state and kind when it is
not a visible worksheet; `serve` stays one sheet per page and marks a chartsheet or dialog sheet
in its sheet list instead of linking it.

### 4.16 Rich-text runs

**Captures** (Excel 16 for Mac, Office Home & Business 2021, Calibri 11, zoom 100,
measured on the screen image), like §4.13's: every rule here stays **unverified until it is
recaptured after the move to Microsoft 365**, and the tests that rest on them
(`tests/test_runs.py`) say so. The code is `_read.py` (reading), `_runs.py` and `_excel.py`
(layout), `backends/html/_text.py` and `backends/markdown.py`.

**Reading.** A cell's string can hold runs, each with its own font (`<r>` with `<rPr>`). openpyxl
gives them only with `load_workbook(rich_text=True)`, and in read-only mode only for shared
strings, which is where Excel writes them: there the value is a `CellRichText`, a run without
properties a `str` and the others `TextBlock(InlineFont, text)`. Its read-only sheet parser is
built without `rich_text`, so the runs of an **inline string** (what openpyxl itself writes) are
lost; the structure scan marks those cells and they are reported as `text.runs_lost`. The reader
always passes `rich_text=True`: on 300 workbooks drawn at random from a corpus of 15,935 (41 of
them with runs), reading every cell took 7.13 s against 7.18 s without it, and the summed peak
of traced memory was 310.6 MB against 307.6 MB (1.0 %; 1.9 % on the files with runs). A shared
string holding both a `<t>` and runs, which the schema allows and Excel does not write, comes
back as its `<t>` alone (what Excel shows is unverified).

**The model.** `Cell.runs` holds the runs as `TextSpan`s (text and fully resolved `Font`), and
`Cell.value` their text joined; escapes (§5.6) are decoded in each run. Adjacent runs in one font
are joined, empty ones dropped, and a text whose every run is in the cell's font has no runs
(`None`). **A run without properties takes the cell's font; a run with properties takes nothing
from it** (captured, R4: in a 16 pt red cell, a run whose properties hold only bold is drawn
11 pt and black): what its properties leave out is the default, and a missing name or size the
Normal font's (unverified: the capture's Normal font is Calibri 11, which is also the fixed
default).

**Layout** (`excel` mode). The text is measured run by run, each in its own font at its drawn
size, for spill, `####`, wrapping, shrink to fit and alignment:
- **Widths add per run** (captured, R2: a bold 16 pt run in an 11 pt text ends 15 pt further,
  166 pt past the cell's right edge against 151, and spills one column more).
- **Each line takes the height of its largest run** (captured, R1: one 20 pt run in an 11 pt
  cell fits the row to 26 pt, as a 20 pt cell, and the small runs sit on the large run's
  baseline). Wrapped lines are each their largest run's row height plus 1 (captured, R3:
  "aaa" / "BBB" / "ccc ddd" with "BBB" at 20 pt is 16 + 27 + 16 = 59 pt; the 27 for a 20 pt line
  is inferred from this one capture). Calibri 20's 26 pt joined the row-height table (§4.7) from
  R1. The baseline at 20 pt follows §3.6's scaled rule, 6 pt above the cell's bottom where the
  capture shows 5, as it does for a 20 pt cell.
- A fitted row counts its runs' fonts in place of the cell's (unverified when every run is
  smaller than the cell's font).
- **Superscript and subscript runs** follow §4.13 (captured, R5: rows of 17 pt, as whole cells),
  each run raised or lowered on its own; the cell's own font counts only through the runs that
  take it.
- Underlines and strikethrough are drawn per run; an accounting underline under text spans the
  cell once per line, whichever runs have it.
- `fill` repeats the runs; justified and distributed words keep the runs they cut.
- The number format's section color replaces each run's color, as it replaces a cell's
  (unverified). A text section that adds characters (`@" more"`) shows other text, which is drawn
  in the cell's font (unverified).

A `TextRun` of a cell with runs carries `spans`: each piece of its line as a `TextSpan` with its
width, drawn size (`scale`) and baseline shift; the run's own `scale` is then the step of shrink
to fit alone.

**HTML.** A run with spans is the positioned run of §7.3, in the font of its largest span, with a
nested span per run. **Each nested span has `line-height: 0`**: measured in WebKit, a 20 pt
nested span in the default way grows the line box and moves the baseline down 10 px, and with
`line-height: 0` the baseline stays where §7.3 puts it, which is what Excel draws (R1).
Superscript and subscript spans are moved by a relative offset in whole pixels, never
`vertical-align`, which moves the baseline too. The Calibri `letter-spacing` is computed per
span, from its own font and Excel's width. A rule takes its run's color. A run with spans always
keeps its clip window.

**Other outputs** (decided; each declared in `backends/_support.py`). The census counts a cell
with runs as `text.runs` and under each font feature its runs use, so an output that leaves out
bold reports a bold run too.

| `rich` | text | Markdown | CSV, TSV |
|---|---|---|---|
| approximated: drawn in the cell's font, since characters mode keeps one font for each cell's text | omitted | with `emphasis=True` each run's bold and italic as `**` and `*` (shown); otherwise omitted | omitted |

**Phonetic guides** (`rPh`, `phoneticPr`) stay out of scope and reported (`text.phonetic`):
openpyxl drops them in every mode.

### 4.17 Rotated and stacked text

`textRotation` is 0 to 90 degrees counterclockwise, 91 to 180 for −1 to −90 (stored as
90 − angle), and 255 for stacked text (ECMA-376 §18.8.1). It is read into `CellStyle.rotation`
(−90 to 90) and `CellStyle.stacked`. The rules are captured by the screen suite `rotation`
(`tests/captures/rotation.tsv`, Excel 16.115; Calibri 11 unless a case sets a size, zoom 100,
"Rotated text here" 77 pt wide upright, 53 pt columns).

**Angled text** is one line turned about the start of its baseline (`TextRun.angle`). Excel's ink
boxes match a rigid rotation of the upright box (captured: 45°, 58.5 × 57.5 pt against 59.7 ×
59.7 computed).
- **Placement:** the line's box (its width, from the descent to the ascent of the line metrics of
  §3.6), turned, is placed by its upright bounding box. General text goes left for 1 to 89
  degrees and −90, and right for −1 to −89 and 90 (captured at 15, 30, 45, 60, 75, 80, 85, 89 and
  their negatives, 90 and −90: the switch is exactly at ±90). Left and right use insets of 3 and 4
  pt, fitted to the capture: every captured ink edge is within 1.1 pt. Center centers the box on
  the cell (captured at 45°); right puts it against the right inset (captured at 45°). An indent
  moves the box in by 6 pt a level (captured: 12 pt for 2, left and right). Fill is placed as
  left (captured).
- **Vertical:** bottom puts the box's lowest corner on the cell's bottom, top its highest on the
  top, center centers it (captured in a 90 pt row at 45°). A box taller than its row
  stays bottom-aligned and is clipped at the row; nothing is drawn in the row above (captured,
  a 15 pt row at 45° and 90°).
- **Spill:** angled text spills into empty neighbors as horizontal text does (§4.1), by its box's
  extent: right for positive angles, left for negative ones, both when centered (captured).
- **Fitted rows:** a row Excel fits grows to `W·|sin θ| + S·|cos θ| + 7` pt, to the nearest
  point, with `W` the text's width and `S` its font size. Captured over three lengths, four angles
  and four sizes: within 2 pt at 9 and 11 pt (at 90° Excel's row is exactly `W + 7`), but Excel
  measures larger angled text wider than xlformat's widths, by about 0.13 pt a character at 14 pt
  and 0.5 at 20 pt, so long text at 20 pt fits rows up to 19 pt taller (a known gap). A merge does
  not count, as for wrapped text.
- **Wrapped angled text** (wrap, justify or distributed) wraps along the turned direction: a line
  holds what fits between the row's top and bottom along the text, and the lines step
  perpendicular to it; the row is not fitted. At 45° in a 15 pt row every line holds one
  character, so the characters stand about 21 pt apart across the row and beyond the cell
  (captured). sheetpeek draws such text on one line and reports it (`text.rotation_wrapped`).
- **Values show in full:** an angled number is neither `#` nor General shortened to the column
  (captured: 12345678901 at 45° and 90° in a 53 pt column); the layout formats an angled cell
  without its column's width.

**Stacked text** draws one upright character a line, a font line apart (captured: 15 pt at
Calibri 11), centered for general text (captured); left and right sit at the insets
(**unverified**). The lines are placed by the vertical alignment as wrapped lines are, clipped at
the cell. A fitted row grows to a line a character plus 2 pt (captured: nine cases of 3 to 8
characters at 9, 14 and 20 pt, all exact).

**Borders of angled text slant.** Excel slants the left and right borders at the text's angle
into a parallelogram whose bottom edge is the cell's: at a height `y` above the bottom, a border
lies `y·cot θ` right of its upright place (captured at 15°, 45° and −45°, at 25 % and 75 % of the
row; 90° and −90° stay upright). Any side border makes the text move as with all four, top and
bottom alone included; a fill alone does not (captured). The text moves right by about
`15·cot θ − 4` pt toward the slant (captured: 52.5 pt at 15°, 11 at 45°, −11 at −45°; the rule
is fitted to three angles). sheetpeek draws borders upright, places the text as without borders,
and reports every angled cell with a side border (`border.slanted`, approximated).

**HTML** turns the run with `transform: rotate()` about the start of its baseline, and its
underlines and strikethrough with it; the clip window stays upright. Measured in WebKit with neither
Calibri nor Carlito installed, the captured cases' ink lies within 2.5 pt of Excel's at both
ends; the larger differences are at the far end, where the fallback font's width differs (§7.3). **Other outputs** draw
angled and stacked text horizontally (`text.rotation`, `text.stacked`, approximated in `rich` and
text, omitted in Markdown and CSV, which show the full text).

## 5. openpyxl facts that constrain the design

All measured with openpyxl 3.1.5 (`bench/`).

### 5.1 What read-only mode gives up

`load_workbook(read_only=True)` returns `ReadOnlyWorksheet`, which has **no** `merged_cells`,
`column_dimensions`, `row_dimensions`, `sheet_format`, `dimensions`, `freeze_panes`,
`sheet_view` or `conditional_formatting` (each an `AttributeError`). **Styles survive**:
`ReadOnlyCell` has `number_format`, `font`, `fill`, `alignment` and `border`, identical to normal
mode in the probe. Workbook-level facts survive too: `epoch`, sheet states, the Normal style and
the theme XML. `max_row` and `max_column` come from the file's `<dimension>` element. Empty
positions come back as `EmptyCell`, which has no `coordinate` and no style. `iter_rows` with
explicit bounds pads leading and inner empty rows but stops after the last row holding data,
whatever `max_row` says, so the reader pads the remaining rows itself.

Two writer facts matter for tests:
- openpyxl writes an uncached formula as `<f>…</f><v></v>`, an empty value element rather than
  none.
- openpyxl drops a column width of 0, treating it as unset.

### 5.2 Memory and time

A workbook of 100,000 rows × 20 columns is 2,000,000 cells: floats with a number format, strings,
datetimes and integers, with bold and a fill on every 50th row. It is 19.7 MB on disk and
108.7 MB of XML. The table gives peak RSS of a fresh process that reads every cell (`import
openpyxl` alone is 28 MB):

| mode | load | iterate | total | peak RSS |
|---|---|---|---|---|
| normal | 6.86 s | 0.62 s | 7.49 s | 931 MB |
| normal, `data_only` | 6.74 s | 0.62 s | 7.36 s | 933 MB |
| normal, `data_only`, styles read per cell | 6.67 s | 3.96 s | 10.63 s | 931 MB |
| read-only, `data_only` | 1.71 s | 4.86 s | 6.58 s | 37 MB |
| read-only, styles read per cell | 1.78 s | 6.21 s | 7.99 s | 38 MB |
| read-only, date conversion off | 1.71 s | 4.44 s | 6.16 s | 38 MB |

The first 100 rows in read-only mode take 1.76 s, most of it loading the shared strings (500,000
distinct strings in this file). In normal mode, every request pays the full load. Normal mode
costs about 465 bytes per cell. Read-only mode's memory does not grow with the sheet.

### 5.3 Dates lose information in openpyxl's conversion

openpyxl turns a number under a date format into a `datetime`, `time` or `timedelta`, rounded to
the millisecond (`openpyxl.utils.datetime.from_excel`):

- Serials 59 and 60 both become `1900-02-28`. Excel shows serial 60 as `1900-02-29`, and xlformat
  reproduces that (quirk `leap_year_1900`) only if it gets the serial.
- 106 of 200,000 random serials shown with `hh:mm:ss` display a different second after the round
  trip through `datetime` (`28139.378153931324`: `09:04:32` from the serial, `09:04:33` from the
  `datetime`). The rate is 42 for `hh:mm:ss.000` and 110 for `[h]:mm:ss`; `h:mm` showed none.

The conversion happens in openpyxl's worksheet reader when the cell's style id is in
`wb._date_formats` (`wb._timedelta_formats` selects `timedelta`). In read-only mode, worksheets are
parsed lazily, so emptying both sets after `load_workbook` returns and before iterating gives the
raw serials (measured).

**Run-time detection and fallback.** The reader checks that both attributes exist and are sets
before relying on them. If not, it reads openpyxl's converted values and turns each `datetime`,
`date`, `time` or `timedelta` back into a serial (`openpyxl.utils.datetime.to_excel`), and the
sheet records one warning that dates may be off by a millisecond and that serial 60 shows as
February 28. The tests pin both paths:
- the fast path returns serial 60 and `28139.378153931324` unchanged;
- the fallback, forced by hiding the attributes, returns the converted serials and the warning.

openpyxl has no upper version bound. One CI job runs the tests against the latest openpyxl
release.

### 5.4 Reading strategy

**Decided:** cells are read in read-only mode with date conversion off. A structure scan reads what
read-only mode drops. Normal mode was rejected: it costs about 465 bytes per cell, loses the leap
day and some seconds, and makes a 100-row preview of a large file as slow as reading the whole
file.

The scan is one pass of `xml.etree.ElementTree.iterparse` over the worksheet part. It records:
- `<sheetView>` and `<sheetFormatPr>`;
- `<cols>`;
- each `<row>`'s `r`, `ht` and `hidden`;
- each `<c>`'s `r`, for the used range;
- `<mergeCells>`;
- through features' own handlers, `<hyperlinks>` and the formulas that call `HYPERLINK()`
  (`_scan_links.py`, §4.10).

Cells are handled at the end of their row, after which the row is cleared. A row's column extent
comes from its first and last cell, since the schema orders cells by column; each cell is
touched only to notice a formula without a cached value. The worksheet part's path comes from
the package itself, not from openpyxl's private `_worksheet_path`: `xl/workbook.xml` maps sheet
names to relationship ids, and `xl/_rels/workbook.xml.rels` maps those to parts.

Measured on the 2-million-cell file:
- The scan takes 2.40 s at 30 MB peak.
- `sheetpeek.read` of the whole sheet takes 11.3 s at 309 MB peak. Most of that memory is the
  model's own grid of cells. Normal mode would need 931 MB before the model is built.
- A 100-row preview (`ref="A1:T100"`) takes 4.3 s at 38 MB, made up of the shared strings
  (1.8 s), the scan (2.4 s) and almost nothing for the cells.

The scan still tokenizes the whole part, so a faster preview would need to stop it early. It
cannot stop at the end of `<sheetData>` because `<mergeCells>` and `<hyperlinks>` follow it.
A preview doesn't need this yet. A test reads
files whose `<dimension>` is missing, too small and too large, and checks the used range and that
no cell is dropped. The read-only iterator is given the scan's bounds explicitly.

### 5.5 Untrusted files

**Refusals come before anything parses the file**, in `_package.open_package`, which `Book` calls
before `load_workbook`. openpyxl opens its own `ZipFile`, reads most parts whole and offers no
hook, so guards around sheetpeek's own reads would not cover it. It does use the same `zipfile`
on the same bytes, so a check of every entry covers whatever it reads, including what it reads
lazily and whichever part it decides is the workbook.

**The check**, `_package.check_package`:
1. From the central directory, with nothing decompressed: the number of parts, each part's size
   and the total against `ReadLimits`; a name used twice; a method other than stored or deflate;
   encryption.
2. One streaming pass over every entry, whatever its name: inflate it through `zlib` from its
   local header and refuse it unless it produces exactly the stated size; hand its first bytes to
   `_package._Prolog`.

**Why the sizes must be made true.** `zipfile` stops its output at the central directory's size,
on Python 3.11 to 3.14 and for all four methods, and a CRC of the prefix makes that read succeed.
But `ZipFile.read` decompresses first and cuts afterward: a 1 MB deflated part stating 1,000 bytes
peaks at 2.07 GB in 0.3 s, up to the 2 GiB `ZipExtFile.MAX_N` allows, and bzip2 and LZMA, which
take no output limit at all, do the same without that bound. `ZipFile.open` with a bounded read
stays at 18 MB, but openpyxl calls `read` for most parts. Measured with `bench/load_modes.py` on
the workbook of §5.2 (108.7 MB of XML), the whole check takes 0.08 s at 31 MB peak, against 4.3 s
for a preview. Once it has passed, every later read, openpyxl's included, can rely on the sizes.

**A DTD is refused, in every part, under any name.** expat itself reads the prolog
(`StartDoctypeDeclHandler`, stopping at the first start element, within 1 MiB), so a comment
before the DTD, UTF-16 and any other encoding the parsers accept are seen as the parsers will see
them; a part's name decides nothing, because `[Content_Types].xml` and the relationships name the
parts. Without a DTD there are no entities, which removes:
- expat's own defense as a dependency. Amplification limits exist only from expat 2.4: under
  2.2.8 a billion-laughs part was still growing when stopped at 4 GB after 60 s. The parser also
  differs by what is installed, since openpyxl switches to lxml or defusedxml when they are
  present;
- what that defense allows. It is a ratio, not a cap: a 36 KB file with one large entity used 90
  times expands to a 3 GiB string, 5.9 GB peak.

External entities were never fetched (file and HTTP, general, parameter and external subset, with
and without lxml and defusedxml), and a refused DTD keeps it so. A part expat cannot read at all,
such as an image, passes: no parser that would expand an entity accepts it (lxml reads UTF-32,
which expat cannot, and leaves its entities unexpanded). An encoding expat refuses by name is
refused here too, because lxml may accept it.

**What Excel does** (suite `package`, `CAPTURING.md`): it opens none of these. A `<!DOCTYPE>`, a
bzip2 or LZMA part, a name used twice and all three kinds of false size give error -50 with alerts
off, between two controls that open. So a refusal never turns away a file Excel opens without a
repair. A DTD stays refused whatever Excel does.

**Numbers no column, row or font can have** (suite `package`, 68 cases) reach no arithmetic.
`_structure.py` and `size` in `_style/font.py` take them as Excel does, and refuse what Excel refuses:

- **`<col width>`** is kept in points modulo 65,536, at 6 pt a digit: 256 and 1e4 are as usual,
  1e6 is 6036.5 units, -1 is 10921, -100 is 10822 (a negative one with one more point). From 1e9 to
  1e30 it is one point (0.09 units). NaN, either infinity and 1e308 are 0, a hidden column.
- **`<row ht>`** is its magnitude: -0.5 is 0.5, -100 is 100. Past 409.5 it wraps in twips modulo
  65,536 and is capped there: 410, 1e6, 1e9 and 1e30 are 409.5, and 1e4 is 169.5. NaN and either
  infinity are ignored, and 1e308 is 0.
- **Font size** is its magnitude from 1 to 409: -1 is size 1. 0, 0.5, -0.5, 410, 1e6, 1e30, 1e308,
  NaN and infinity are refused.
- **`defaultColWidth`:** NaN and infinity are ignored, and 1e308 makes every column 0. Negative
  numbers, 300, 1e30 and `99999999999` are refused.
- **`baseColWidth`:** 300 is refused; -1 and `99999999999` are ignored.
- **`defaultRowHeight`:** NaN and infinity are refused. Without `customHeight` a number is ignored,
  as ever (§4.7). With it, -1 is 1 pt, and 1e308 acts as `zeroHeight`.

**Unverified:** the 6 pt digit, since every case has Calibri 11 as its Normal font; where 0 starts
between 1e30 and 1e308; the edges at 255 to 300 (`defaultColWidth`, `baseColWidth`) and 409 to 410
(font size); and whether a font size below -1 is its magnitude, as -1 is. Known deviation: a row of
font size 1 is 4 pt in Excel, and sizes under 9 are outside the row-height table (§4.7).

**A column in `characters` mode is at most 255 characters**, Excel's own limit, whatever its width
in points. Each character costs about 10 bytes in every row while the rows are built (the canvas),
so without the cap a tiny file with a width of -1 needed 1.1 GB per 10,000 rows, and about 111 GB
for a full column; with it, a full column peaks at 3.6 GB. `excel` mode keeps Excel's width, which
costs nothing per point.

**`ReadLimits`** has `max_parts`, `max_part_bytes`, `max_total_bytes` and `max_cells`:

| limit | library, `view` | `serve` |
|---|---|---|
| `max_parts` | 100,000 | 10,000 |
| `max_part_bytes` | 4 GiB | 256 MiB |
| `max_total_bytes` | 16 GiB | 1 GiB |
| `max_cells` | 100,000,000 | 1,048,576 (one full column) |

**`max_cells` is a grid area**, not a count of populated cells. It is enforced before anything is
allocated, on what is actually read: `ref` when given, otherwise the used range from the structure
scan, never `<dimension>`. Without it, a 4.6 KB file with a merge `A1:XFD1048576`, a cell at
`XFD1048576` or a row past Excel's last one reaches 6 GB in 30 s, and so does
`?ref=A1:XFD1048576` on an ordinary file. It applies again to the layout, through `Sheet.max_cells`:
one 32,767-character string, which Excel allows, spills over 3,710 columns, so `A1:A1000` of that
sheet is 3.7 million rendered cells, 2.3 GB. Empty columns past the range are therefore neither
formatted nor counted until the layout knows which ones spilled text reaches.
The layout's map of merges is clipped to the range read, because a merge can cover the sheet.

**What the limits leave.** Parts are not equally expensive, and the costs are proportional to
what the limits allow:

| part, sizes honestly declared | on disk | time | peak |
|---|---|---|---|
| `sharedStrings`, 256 MiB of 24-byte entries | 0.6 MiB | 33 s | 2.5 GB |
| `sharedStrings`, 1 GiB of 24-byte entries | 2.5 MiB | 131 s | 7.2 GB |
| `sharedStrings`, one 1 GiB string | 1.0 MiB | 1.9 s | 2.4 GB |
| styles or theme, 256 MiB of elements, parsed whole | 0.3 MiB | 29 s, 10 s | 6.5 GB |
| worksheet, 1 GiB of one-cell rows | 3.0 MiB | 169 s | 2.6 GB |

A lower limit only for parts parsed whole was ruled out: which parts those are is openpyxl's
internal detail (chartsheets and drawings among them), and it finds the workbook part differently
from sheetpeek, so a split could be evaded. The library's defaults stop what could hardly be read
anyway and do not make it safe on a small machine: 100 million rendered cells take tens of GB. A
service lowers them, as `serve` does.

**Who reaches `serve`.** It has no upload and answers only GET and HEAD. A crafted file arrives
through whoever can write to the folder, or a file the operator saved there; the 6.5 GB above
needs that. Any HTTP client chooses the file, `sheet` and `ref`, with no authentication, and is
stopped by the cell limit: a full column of one number costs 6.6 s and about 1 GB. Two reads run at
once, each holding its slot through layout and rendering. A request that gets no slot within 2 s
is answered 503 with `Retry-After`, not queued, since blocked threads would be unbounded too. Bound
to a loopback address, the server answers only to `localhost`, `127.0.0.1`, `[::1]` and that
address, with 403 otherwise: without this, DNS rebinding would let any web page read the folder's
sheets. Only the name is compared, as Django's `ALLOWED_HOSTS` and Vite's `server.allowedHosts`
do, because a tunnel brings its own port. Bound to anything else it answers to any name and anyone
who can reach it can read the folder; it has no authentication.

**What reaches a terminal.** A control character in output is a command to the terminal: ESC, and
U+009B, the one-character form of `ESC [`, which XML can hold. Cells lose them in the layout (§4.1).
Everything else the command line prints that a workbook, a file name or a client can influence
goes through `_model.one_line`, which also drops line breaks, so a name cannot forge a line: every
message on standard output and error, sheet names from `--list`, the sheet name and font names in
HTML, and `serve`'s log lines, since a request line is anyone's and overriding
`log_message` gives up `http.server`'s own escaping. The model keeps the original, so `--sheet`
matches the name as the file has it. In `serve`, sheet names in pages are only escaped for HTML.

**What reaches markup.** The HTML backend escapes cell text and sheet and font names; a theme color
is used only if it is six hex digits, since it goes into a `style` attribute. Markdown escapes what
acts as markup in a table cell, which holds inline content only: `\\`, `` ` ``, `*`, `_`, `[`, `]`,
`<`, `&`, `~`, `|`, `$` and `@`, and the first character after a bare link's scheme or `www`, with no
other change, so numbers stay readable. Checked with cmark-gfm, markdown-it and mistune, with HTML
allowed. A lone carriage return is dropped, since it would end the row. An email address remains a
link in cmark-gfm whatever is escaped: the same text, as `mailto:`.

**One open file.** `Book` opens a path once, for the check and for every read after it, so a path
swapped after the check is not what gets read. A file overwritten in place during a read still
could be, and `Book` is for one thread at a time. openpyxl's check of the file's extension is
therefore not reached; `_package.open_package` gives the same advice for `.xls` and `.xlsb`,
looking at the content first and the name second.

**Errors.** `PackageError` is a `ValueError`, and `LimitError` a `PackageError`, as in xlrecalc;
the classes are separate, since xlrecalc is optional. Whatever openpyxl raises on a file it cannot
read becomes a `PackageError` with the original as its cause, and its message drops the file's
folder, since openpyxl's own names the full path and `serve` shows messages to clients. A path
that cannot be opened keeps its `OSError`. **Measured for no false refusals:** 15,935 workbooks in
two public corpora all pass the check, at about 1 ms each; `sheetpeek.open` and openpyxl agree on
every one (2 files openpyxl cannot read, for a print title of `#N/A`).

### 5.6 Text as openpyxl returns it

openpyxl's read-only mode leaves OOXML's `_xHHHH_` escapes undecoded in cell text, inline strings
and cached text results. Excel decodes them (`layout` rows 190 to 193, and xlrecalc's capture of
the same rule): `_x0041_` is `A`, hex digits in either case, `_x005F_x0041_` the text `_x0041_`, and
three digits no escape. `_read.unescape` decodes every text value as `Cell.value` takes it.

Two deviations, reported by count in pytest's "captures" summary and pinned by a test, come from
what openpyxl has already done. Both are rare, and reading the shared strings without openpyxl
would fix both:
- **`x005F_` in shared strings.** Its shared-string reader removes `x005F_` wherever it occurs, so
  `_x005F_x0041_` arrives as `_x0041_` and decodes to `A` where Excel shows `_x0041_`, and
  `x005F_keep` shows as `keep`. Inline strings keep it (row 191).
- **Whitespace without `xml:space="preserve"`.** Excel drops leading and trailing whitespace there
  (xlrecalc's capture `whitespace`). openpyxl keeps it and does not say what the element said.

Rich text is read with openpyxl's `rich_text=True`, which keeps the runs of shared strings and
loses those of inline strings (§4.16).

## 6. The xlformat seam

### 6.1 The call

For every cell holding a number, text, boolean or error:

```python
xlformat.format_cell(value, number_format, column_width=w, font=cell_font,
                     normal_font=normal_font, date1904=sheet.date1904)  # excel mode
xlformat.format_cell(value, number_format, width=n, font=cell_font,
                     normal_font=normal_font, date1904=...)             # characters mode
xlformat.format_cell(value, number_format, font=cell_font,
                     normal_font=normal_font, date1904=...)             # none mode
```

- `value` is the raw serial for dates (§5.3), never openpyxl's `datetime`.
- `w` is the column width in dialog units, from the file width through xlformat's conversion
  (§6.4).
- For a merge anchor, `w` is the merged width (§4.4).
- `cell_font` and `normal_font` are `xlformat.Font(name, size, bold, italic)`, from the cell's
  font and the Normal style's, built once per distinct font. `characters` and `none` modes count
  characters, so there the fonts change nothing yet; they are passed so every mode makes the
  same call. `file_column_metrics` takes the Normal font as `normal_font=`.
- `NumberFormat.parse` is cached by xlformat, so a repeated format costs nothing extra.

### 6.2 Failures

- **`FormatError`**: the cell is formatted again with `"General"` at the same width, and the sheet
  records one warning naming the format. This is what Excel's repair does.
- **A font without a glyph table** (xlformat measures Calibri and Aptos Narrow; Arial and
  others would raise `LookupError`): **decided**, the layout measures it as Calibri at the same
  size and style, and records one warning per family, only where the font's widths decide the
  layout (the Normal font for column widths; cell fonts in `excel` mode). `xlformat.measured_fonts`
  tells which families are measured. Character mode was rejected because it counts a comma like a
  digit, which makes HTML disagree with Excel on more cells than a proportional approximation
  does.
- **`ValueError` / `TypeError`** (a value no format shows, such as the infinity openpyxl reads from
  `<v>1E400</v>`, or an unexpected type): shown as `str(value)`, left-aligned, with one warning
  per format string naming the first cell. When a refused format's `General` retry fails so, the
  warning names `General`.

### 6.3 Error values

The reader wraps an error value (`data_type == "e"`) as `xlformat.ExcelError`, which xlformat
formats like any value: its text under any number format, centered, with no section color (so
the font color applies), and a `#` run in a column too narrow for it (§4.2). An unlisted code
shaped like an error still renders; a text not shaped like one (a malformed file) is read as
text. sheetpeek does not re-export `ExcelError`.

### 6.4 What sheetpeek takes from xlformat

xlformat provides the measurements and formatting below; row heights (§4.7) are captured by
sheetpeek itself.

| from xlformat | used for |
|---|---|
| `format_cell(value, format, width= or column_width=, font=, normal_font=, date1904=)` → `FormattedCell` | every cell's text, overflow and section color in every mode (`_layout._Formatter`); `FormattedCell.align` is the default alignment (§4.3) |
| `Font(name, size, bold, italic)` and `measured_fonts()` | the cell font and the Normal font passed to xlformat, built once per font; a family xlformat does not measure becomes Calibri, with a warning (§6.2) |
| `file_column_metrics(file_width, base_col_width=, default_col_width=, normal_font=)` → `ColumnMetrics` | column widths: characters from `width` (§2.2), points from `points` (§3.6); `left_inset`, `right_inset` and `digit_width` place text and indent it |
| `column_metrics_from_points(points, normal_font=)` → `ColumnMetrics` | a merged cell's width: its `width` is the `column_width` of the merge's sum of points (§3.6) |
| `text_width(text, font=)` | `excel` mode: spill, placement, wrapping, `distributed` spacing, `fill` copies and `TextRun.width` (§3.6) |
| `glyph_widths(text, font=)` | where the glyphs of a `TextRun` sit, as the model documents it (§3.6); the layout itself measures whole texts |
| `ExcelError` | cached error values, formatted like any value (§6.3) |
| `FormatError` | a refused number format, shown as `General` (§6.2) |
| `DEFAULT_PALETTE` | the `RRGGBB` of a section color such as `[Red]` (§3.3) |

## 7. Backends

### 7.1 Precedent

- **pandas `Styler`** (pandas 3.1.0): `_translate()` builds one dictionary of cells with
  `display_value`, classes, spans and visibility. Four Jinja2 templates render it (`html.tpl`,
  `latex.tpl`, `typst.tpl`, `string.tpl`). It has no third-party output registration.
- **tabulate** 0.10.0: built-in formats are a module-level dictionary of `TableFormat` named
  tuples. A caller passes a `TableFormat` directly as `tablefmt` instead of a name. There is no
  registration and no entry points.
- **rich** 15.0.0: no registry; anything with `__rich_console__` (the `ConsoleRenderable`
  protocol) or `__rich__` renders.
- **Pygments** 2.21.0: built-in formatters are looked up by name, and third-party formatters
  register through the entry-point group `pygments.formatters`.
- **pandas plotting**: the `pandas_plotting_backends` entry-point group, selected by name.

### 7.2 Interface

```python
class Backend(Protocol):
    name: str
    mode: Literal["excel", "characters", "none"]
    def render(self, sheet: RenderedSheet, out: TextIO, /) -> None: ...
```

A backend writes to a text stream, so the server and large sheets stream instead of building one
string. Backends are classes whose constructors take their drawing options; layout options
(`show_hidden`, `scale`, `max_width`) belong to `Sheet.layout`. Built in:

- `TextBackend(separator=" ", column_headings=False, row_headings=False)`, name `text`;
- `RichBackend(separator=" ", column_headings=False, row_headings=False, color_system="auto",
  force_terminal=None)`, name `rich`;
- `HTMLBackend(fragment=False, column_headings=False, row_headings=False,
  color_scheme="light")`, name `html` (§7.3);
- `MarkdownBackend(header="first" | "letters", emphasis=False)`, name `markdown`;
- `DelimitedBackend(delimiter)`, names `csv` and `tsv`.

They receive a `RenderedSheet` already laid out in their mode and never call xlformat.
`render()` lays a `Sheet` out in the backend's mode, and lays a `RenderedSheet` out again when
its mode differs.

**Registration (decided): a plain registry, and backend objects passed directly.**
`render(..., to="html")` looks the name up in a module-level dict of the built-ins. `to=` also
accepts any object that satisfies `Backend`, as tabulate accepts a `TableFormat`, so a
third-party backend needs no registration. `register_backend(name, factory)` adds a name. Entry
points (a `sheetpeek.backends` group, as Pygments does) can be added later without breaking
anything.

`Sheet` and `RenderedSheet` implement `__rich_console__` and import `rich` only inside it, so
`rich.print(sheet)` works without sheetpeek importing rich at module import time.

### 7.3 HTML

The HTML backend draws `excel` mode's geometry (§3.6) as a **static** page, or as a fragment to
embed: no script, and nothing read from the viewer's system. Geometry is exact (column x, row y,
run starts, clip windows, fills, gridlines); glyphs inside a run may be off by a couple of
points.

- **A fixed-layout table** carries the cells: `table-layout: fixed`, a `<col>` per visible
  column, a height per row, each cell drawing its own left and top gridline as a 1 px border and
  the table its right and bottom edges. A merge is one cell with `colspan`/`rowspan` and the
  anchor's fill; its covered cells are left out.
- **A text layer:** each `TextRun` is an absolutely positioned span inside its origin cell,
  placed in sheet coordinates against the `.sheetpeek` wrapper, inside a span for its clip
  window (`overflow: hidden`). It is anchored by its `align`: left text by its start, right text
  by its end, centered text by its center, so a font with other widths keeps Excel's anchor and
  is still cut where Excel cuts. **The cells are not positioned**, so the browser paints every
  cell background before any run: spilled text lies over each neighbor's own fill, and the
  spilling cell's fill stays home. Two competing spills are both painted. Each run stays inside
  its origin cell in document order, so copying and screen readers see the cell's text there.
- **Whole CSS pixels.** Every position is `round(pt × 4/3)` px and every size the difference of
  two rounded positions. Measured in WebKit on a page of 300 columns and 310 rows: WebKit floors
  each row height to its 1/64 px unit, so rows of 19 pt (25.33 px) given in points drift 1.2 pt
  in 300 rows; in whole pixels cells stay within 0.25 pt of their exact place at any size, and
  every run keeps its exact offset from its own cell. Font sizes stay in points.
- **Fills** paint the cell and its own gridlines, and the right neighbor's left gridline and the
  top gridline of the cell below take the fill's color, as Excel's fill covers them. A crossed
  boundary (`RenderedRow.crossed`) gets the class `sp-crossed`, a transparent gridline.
- **Pattern and gradient fills, and borders** (§4.12) are two layers of absolutely positioned
  boxes (`sp-fills`, `sp-borders`) written before the table's text, so text lies over them and
  they over the cells' backgrounds and gridlines; their sizes are CSS px, not points. Each
  distinct fill or line is a generated class in the page's stylesheet (`sp-f-` and `sp-b-` and
  letters, not public interface).
- **Baseline:** with `line-height: 1`, a run's `top` is its baseline less 0.8418 em, in pixels;
  0.8418 em is where the baseline of Calibri and Carlito sits in a 1 em box (ascent 0.952,
  descent 0.269).
- **Clip spans only where needed:** a Calibri run of printable ASCII, spaced to Excel's width,
  whose box lies inside its clip window by Excel's measure is written without the clip span,
  positioned against the wrapper; every other run keeps its window.
- **Every run is anchored at its alignment edge**, as Excel anchors it: `left:` for left text,
  `right:` (from its window, or from the wrapper's right edge) for right text, its center for
  centered text. A font with other widths then moves where letters start, never where right or
  centered text ends.
- **Glyph spacing.** Excel for Mac draws glyphs on whole points; a browser draws a font's own
  advances, so Calibri runs about 7 % narrower. A Calibri run of printable ASCII gets
  `letter-spacing` = (`TextRun.width` − Calibri's own width) / characters, the own width from
  Carlito's advances (SIL OFL, metric-compatible with Calibri), bundled as a 2.9 KB table of
  numbers (`_carlito.py`, generated by `tools/carlito_widths.py`; no font file, nothing read from
  Calibri's file); bold, italic and bold italic within 2.5 pt in Calibri. No free font is known
  to be metric-compatible with Aptos Narrow, so its runs are not corrected.
- **Fonts.** Calibri text uses `'sheetpeek-calibri', 'sheetpeek-calibri-fallback',
  sans-serif`. `sheetpeek-calibri` is an `@font-face` whose only sources are a locally installed
  Calibri or Carlito, so the Calibri spacing lands on Calibri's glyphs. On most machines neither
  is installed (Excel for Mac keeps Calibri inside its app bundle). `sheetpeek-calibri-fallback`
  is then a local Arial, Helvetica or Liberation Sans (metric-compatible with each other) with
  `size-adjust` 91.6 % (bold 90.2, italic 91.4, bold italic 90.0: Carlito's widths over Arial's on
  spreadsheet text), and ascent and descent overrides that give it Calibri's 0.952 and 0.269 em,
  as Next.js builds fallback fonts. CSS cannot make the spacing depend on the font that loaded;
  the scaled fallback makes the same spacing fit it closely. Other families are named with a
  `sans-serif` fallback and are not spaced. A theme font is named by the theme's typeface,
  which the reader resolved (§4.14). Kerning and ligatures are off, as in Excel.
- **Measured in WebKit, all 205 runs of the `layout` page, in three cases** (Calibri or Carlito
  given to `sheetpeek-calibri` in a scratch harness, and the page as shipped on a Mac with
  neither):

  | font | left starts | right ends | centers | clips | glyphs |
  |---|---|---|---|---|---|
  | Calibri | 0.25 pt | 0.25 pt | 0.26 pt | 0.25 pt | 2.0 pt |
  | Carlito | 0.25 pt | 0.25 pt | 0.26 pt | 0.25 pt | 2.0 pt |
  | neither (scaled Arial) | 0.25 pt | 0.25 pt | 0.34 pt | 0.25 pt | 9.0 pt (median 0.6, 90 % within 2.8; the worst is thirty `i`s) |
- **Paper, headings and colors.** Excel 16 for Mac keeps its cells white in dark mode, so the
  paper is always white and automatic ink black. `column_headings` and `row_headings` add `<th>`
  cells (`sp-colhead` in a `<thead>`, `sp-rowhead` with `scope="row"`, `sp-corner`);
  `color_scheme="light" | "dark"` (default light, never read from the system) changes only the
  headings, as in Excel. A sheet without gridlines gets `sp-nogrid`. The heading sizes and
  colors and the grid color are §4.14's.
- **The outline gutter** (§4.11) is cells of the table, so it scrolls with it: a `<col>` and a
  `<td class="sp-ol">` before the row headings, and a row of them above the column headings.
  Each holds absolutely positioned segments (`<i>`), buttons (`sp-olbutton`, strokes as `<i>`, no
  glyph) and level buttons (`sp-ollevel`), sized in CSS px, not points; the column gutter is the
  row gutter's geometry transposed. The level buttons sit in the headings' bands, so they are
  drawn only with both headings. `buttons="none"` leaves the gutter out; `"static"`, the
  default, draws it, and leaves room for working buttons later. Its stylesheet is written only on
  a page that draws a gutter. The dark colors are captured; the light ones are placeholders.
- **The CSS surface is public interface**, changed only with a changelog entry: the custom
  properties `--sheetpeek-paper`, `--sheetpeek-ink`, `--sheetpeek-grid`,
  `--sheetpeek-heading-paper`, `--sheetpeek-heading-ink`, `--sheetpeek-heading-grid`,
  `--sheetpeek-outline-paper`, `--sheetpeek-outline-ink`, `--sheetpeek-outline-button`,
  `--sheetpeek-outline-button-paper`, the classes `sheetpeek`, `sp-crossed`, `sp-clip`,
  `sp-run`, `sp-center`, `sp-colhead`, `sp-rowhead`, `sp-corner`, `sp-nogrid`, `sp-ol`,
  `sp-olbutton`, `sp-ollevel`, and the ids `sp-<cell>` of cells a link goes to. Explicit
  colors from the file (fills, font and section
  colors) are inline styles, so restyled dark paper keeps a cell's explicit dark text (§9).
- **Links** (§4.10): a linked cell's runs sit inside one `<a>`, so spilled text stays
  clickable; the anchor inherits color and decoration (`.sheetpeek a`), since the cell's format
  draws a link. `href` only for an allowed URL (with `rel="noopener noreferrer"`, and
  `target="_blank"` with `new_tab=True`) or a cell on the page (`#sp-C2`, the id on the cell
  that draws it, a merge's first visible cell); `title` is the tooltip alone.
- **Text styles** (§4.13): underlines and strikethrough are `<i class="sp-rule">` elements, 1 px
  at any page scale, a double one `sp-double`; `sp-rule` and `sp-double` join the public classes.
  Superscript, subscript and shrunk runs take their drawn size; a right-to-left run
  `direction: rtl`.
- **Rich text** (§4.16): a nested span per run inside the positioned run, each with
  `line-height: 0`, its own font, color and Calibri spacing; superscript and subscript runs by a
  relative offset.
- **Page or fragment:** a page is a standalone document titled with the sheet's name; a
  fragment is the stylesheet and the `<div class="sheetpeek">`, for a host page.
- **The workbook page** (§4.15) holds every sheet's `<div class="sheetpeek">` in a
  `<div class="sp-book">` with a `<nav>` of tabs; its classes `sp-book`, `sp-first`,
  `sp-hidden` and `sp-notshown`, the section ids `sp-<n>` and the cell ids `sp-<n>-<cell>` join the public
  interface (`sp-t-` and letters, a tab color, is generated).
- **Size** (`bench/html_size.py`, 5,000 rows × 20 columns, 100,000 cells and 97,144 runs, Apple
  M4 Pro): 11.1 MB of HTML, 111 bytes per cell, 0.58 MB gzipped; read 0.60 s, layout 1.77 s,
  render 0.29 s, peak RSS 246 MB. The same grid with every cell using a newer feature (borders,
  pattern fills, runs, angled text, links, outlines) is 33.5 MB, 335 bytes per cell: layout
  2.74 s, render 0.65 s, 362 MB (`bench/README.md`). A sheet pays only for the features it
  uses: the passes and the HTML layers skip what is absent, and fonts, heights and line
  metrics are looked up by the identity of the style's font, which every cell of a style
  shares. The first version, with a clip span and a `calc()` baseline on
  every run, was 19.8 MB (198 bytes per cell); dropping clip spans a run does not need and
  writing the baseline as a number halved it. Layout time is mostly xlformat's formatting, at
  about 19 µs per cell.

**Ruled out:**
- **Plain `colspan` for spill** (Excel's own web export, route A): a spanned cell has one
  background, so the spilling cell's fill would cover its neighbors and a filled neighbor's fill
  would vanish; `colspan` cannot express left or centered spill or two spills sharing a cell; the
  neighbors' cells vanish from the markup; and the browser's font, not Excel's measure, would
  decide where text stops.
- **Per-cell slices of the text:** fragile across fonts, and they break selection.
- **One element per glyph:** exact glyphs, but it breaks selection and copying (a possible later
  option, §9).
- **A script fitting runs to the browser's font:** measured (every glyph within 1.8 pt in
  Calibri, 2.8 pt in Helvetica), but the output is static; whether to offer it is an open
  question (§9).
- **Canvas or SVG as the HTML output:** no table semantics; SVG belongs in a later geometry
  backend.

## 8. Package layout and API

### 8.1 Layout

```
src/sheetpeek/
  __init__.py        public names, __version__
  _model.py          Sheet, Cell, CellStyle, Column, Row, CellRange, Font
  _hyperlink.py      Hyperlink; the scheme check and place lookup every output shares (§4.10)
  _outline_model.py  Outline, OutlineMark, RenderedOutline (§4.11)
  _fill_border_model.py  Pattern, Gradient, Side, Borders, RenderedBorders; the conflict rule (§4.12)
  _fill_border_census.py fills and borders in the census; the style that draws a cell's fill
  _text_model.py     TextRule, the lines drawn with text (§4.13)
  _runs_model.py     TextSpan, a run of rich text (§4.16)
  _rendered.py       RenderedSheet, RenderedRow, RenderedCell, RenderedColumn, Gap
  _package.py        ReadLimits, the check before reading (§5.5), sheet name → worksheet part
  _structure.py      the structure scan (§5.4): one handler per element
  _scan_links.py     the scan's hyperlink handlers, and the reader's link notices (§4.10)
  _scan_outline.py   the outline's scan handlers, and its rows, columns and settings
  _read.py           the reader: openpyxl read-only, date conversion off; builds a Sheet
  _style/            a cell's style, one module per aspect: number format, border, fill, font,
                     alignment
  _detect.py         what a file uses that the model does not carry: one rule per sheet child
  _notices.py        Notice, the registry of codes, aggregation (§3.7)
  _census.py         what laid-out cells use, and each output's notices from it (§3.7)
  _colors.py         theme parsing, tint, indexed palette
  _layout.py         width modes, the xlformat seam, alignment, spill, merges, wrapping;
                     the ordered layout passes
  _pass_outline.py   the outline pass: groups, buttons and dots (§4.11)
  _pass_borders.py   the borders pass: one resolved side per edge, the diagonals (§4.12)
  _excel.py          excel mode: geometry in points, text runs, clips
  _rows.py           row heights (§4.7) and baselines
  _text_styles.py    superscript, subscript, underlines, shrink to fit, direction (§4.13)
  _runs.py           rich-text runs in the layout: measured, broken and cut by run (§4.16)
  _sheets.py         SheetInfo, NotAWorksheetError; sheet kinds, states, tab colors (§4.15)
  backends/
    __init__.py      Backend protocol, registry, render(), notices()
    _support.py      what each output shows of each feature (§3.7)
    _outline_text.py the outline gutter in characters, for rich and text
    text.py  rich.py  markdown.py  delimited.py
    html/            __init__ (backend, page), _css, _table (grid, fills, gridlines),
                     _text (runs), _headings, _links (anchors, ids), _outline (the gutter), _notices
                     _patterns (the fill layer), _borders (the border layer)
                     _workbook (render_workbook: the workbook page, §4.15)
  cli.py             `sheetpeek view`, `sheetpeek serve`
  _serve.py          the server behind `sheetpeek serve`
  _carlito.py        Carlito's advance widths, generated by tools/carlito_widths.py
  __main__.py
tests/               unit tests; captures/ (TSV) and their loader
capture/             capture tooling (not in the wheel)
tools/               the generator of _carlito.py (not in the wheel)
bench/               scripts that rerun §5 and §7.3's figures; README.md records them
docs/                Sphinx, as in xlformat
```

**The reader is isolated.** Only `_read.py` imports openpyxl (a test enforces it), and only the
package's entry points import `_read.py`. Everything else works on sheetpeek's own model, so a
later reader can replace the openpyxl one without touching layout or backends. `_structure.py`
reads the worksheet XML and knows nothing of layout.

**Where a feature plugs in.** Each stage has one place per feature, so that features can be
built side by side with little overlap:
- **reading the sheet XML:** a handler registered by tag in `_structure.HANDLERS`, keeping what
  it finds in `Scan.extra` under its own key (rows are cleared only after every row handler);
- **reading a style:** the aspect's module in `_style/`, whose `read` gives model fields and
  whose `unread` reports what is not drawn yet; a feature moves from one to the other;
- **reporting what is not read:** a rule per sheet child in `_detect.RULES`, replaced when the
  feature comes to be drawn; a code moves from the reader's `CODES` to the census;
- **laying out:** a pass in `_layout.PASSES`, working on the `Draft` (borders after the cells,
  since spill decides them; chrome after the geometry), with its own slot in `Draft.extra`;
- **drawing HTML:** a module per layer in `backends/html/`;
- **telling each output's support:** a line per output in `backends/_support.py`;
- **the model:** a feature's types in a module of their own, with one field on the shared
  dataclasses.

### 8.2 Public API

```python
import sheetpeek

sheet = sheetpeek.read("book.xlsx", sheet="Data", ref="A1:H40")   # Sheet, layout-free
rs = sheet.layout("characters")                                   # RenderedSheet
rs.cell(1, 1).text, rs.cell(1, 1).align, rs.columns[0].width

sheetpeek.render("book.xlsx", to="html", sheet="Data")            # str, from a path
sheetpeek.render(sheet, to="markdown")                            # or from a Sheet
sheetpeek.render(sheet, to="text", out=sys.stdout)                # or to a stream
rich.print(sheet)

with sheetpeek.open("book.xlsx") as book:                         # several sheets, one load
    for name in book.sheet_names: ...
    for info in book.sheets: info.kind, info.state, info.tab_color

sheetpeek.render_workbook("book.xlsx")                            # every sheet, one HTML page

limits = sheetpeek.ReadLimits(max_part_bytes=2**28, max_cells=10**6)
sheetpeek.read(upload, limits=limits)                             # a file from someone else (§5.5)
```

**Functions, not methods (decided).** Methods such as `rs.to_html()` can be added later; they
cannot be removed. **`ref=` names the range**, not `range=`, which would shadow the built-in
(ruff A002). `ref` is openpyxl's own name for a range in a user-facing constructor
(`openpyxl.worksheet.table.Table(ref="A1:E5")`, also `AutoFilter.ref`), matching the `ref`
attribute of `<dimension>`, `<mergeCell>` and `<autoFilter>` in the file format. Callers of
sheetpeek already use openpyxl.

**CLI.**
- `sheetpeek view book.xlsx [--sheet NAME] [--ref A1:H40]
  [--to rich|text|html|markdown|csv|tsv] [--column-headings] [--row-headings] [--hidden] [--list]
  [--all-sheets]
  [--max-parts N] [--max-part-bytes N] [--max-total-bytes N] [--max-cells N]`. The default is
  `rich` on a terminal and `text` when piped or when `rich` is not installed. `--all-sheets`
  writes the workbook page (§4.15), HTML only. Warnings go to
  standard error, and an unreadable or refused file, an unknown sheet or a bad range exits with
  status 1 and one line of explanation.
- `sheetpeek serve DIR [--host 127.0.0.1] [--port 8000]`, with the same `--max-*` options and its
  own defaults (§5.5) (`_serve.py`): a standard-library `ThreadingHTTPServer`. `/` lists the
  `.xlsx` files under DIR, under the folder's name and not its path, `/view/<path>` shows a
  workbook's default sheet through the HTML backend (a fragment with headings inside the server's
  page, with links to the other sheets), and `?sheet=` and `?ref=` choose a sheet and a range. It
  binds to 127.0.0.1 by default and is read-only with no state: GET and HEAD only.
  A request names a workbook only if, after resolving `..` and symbolic links, it is a regular
  `.xlsx` file inside DIR whose path has no part starting with a dot; anything else is a 404. A
  refused or unreadable file is a 422, a request with another host name a 403 and one that gets no
  read slot a 503 (§5.5). There is no authentication, and it is not for production: `http.server`
  is not recommended for it, and a production service calls the Python API with `limits=`.
  Pages carry `Content-Security-Policy: default-src 'none'; style-src 'unsafe-inline'`, since
  they have no script.

### 8.3 Dependencies

- `openpyxl>=3.1`, no upper bound (§5.3).
- `xlformat>=0.3,<1` (§6.4), imported normally: it is required.
- Extra `sheetpeek[rich]`: `rich>=13`.

## 9. Planned options

sheetpeek's defaults show what Excel shows. Where a rendering has no single obviously right form,
or Excel's form does not suit the output, a named option will choose another. Options are
named, not hard-coded, and each has a scope:
- a **document** option applies to the whole layout (an argument of `Sheet.layout` or a
  backend);
- a **column** option overrides it per column;
- a **cell** option overrides it per cell.

The options that exist are `spill_edge`, `show_hidden`, `scale` and `max_width` on
`Sheet.layout`, the backends' `column_headings`, `row_headings` and `color_scheme`, and
Markdown's `header` and `emphasis`; the rest are planned. Behaviors a capture has settled
with one right answer are not listed: Excel's behavior is the only rendering (hidden columns
crossed whatever they hold, one digit per indent level, a merge's text at the bottom by default).

**Spill at the edge of the range** (`spill_edge`; §4.1). Default `"extend"`: empty
columns at the default width past the used range, marked `added`.
- *Alternatives:* `"margin"` (a borderless overflow margin past the last column, not a column);
  `"widen"` (the last column widened to hold its spill); `"clip"`, with an `ellipsis` of `"…"`,
  `"..."` or a custom string.
- *Edge colors:* whether added columns and the margin carry the row's fill or stay unfilled.
- *Scope leaning:* document. A per-column form has no use, since only the outer columns have an
  edge.
- *Related:* when an explicit `ref` cuts through the data, `"extend"` adds nothing on that side,
  because the cells there were not read. Reading one more column per side to learn whether it is
  empty would let `"extend"` work there too.

**Two spills competing for one empty cell** (§4.1). Excel lets both take the cell and draws the
texts over each other, unreadably (captured). Characters cannot be drawn over each other, so
`characters` mode cannot copy Excel; its default is that the first spill placed, left to right,
takes the cell. HTML could overlay the two texts as Excel does.
- *Alternatives:* the right cell's spill wins; the cell is split between them; neither spills
  into it; in HTML, Excel's overlap.
- *Scope leaning:* document.

**Clipped text** (§4.1). Excel cuts blocked text at the cell's edge, through a glyph, with no
marker (captured); `characters` mode cuts at a character.
- *Alternatives:* an ellipsis marker where text is cut, with the same `ellipsis` strings as
  `spill_edge="clip"`.
- *Scope leaning:* document, with a column override for one narrow column of long labels.

**Values that do not fit** (§4.2). Default: `#` filling the column in the section's color, as Excel
shows numbers, dates, booleans and errors (captured).
- *Alternatives:* `"full"` (the full text, spilling or clipped like text); `"truncate"`; an
  ellipsis.
- *Scope leaning:* document, with a column override. A terminal user may want one column's numbers
  never hidden.
- *Related:* a `#` run without the section color (`"plain"`).

**Column widths in characters** (§2.2). Default: Excel's width rounded, with `scale` and
`max_width`.
- *Alternatives:* floor or ceiling instead of rounding; a `min_width`; widths from the content
  (autofit) instead of the file; explicit widths per column.
- *Scope leaning:* document for the rule, column for explicit widths.

**Odd remainders in `characters` mode** (§4.1). Centered text with an odd remainder puts the extra
character on the right, as Excel puts its extra point on the right (captured).
- *Alternatives:* the extra character on the left.
- *Scope leaning:* document.

**Hidden rows and columns** (§3.4). Default: omitted; `show_hidden` keeps them.
- *Alternatives:* kept and marked (dimmed in `rich`).
- *Scope leaning:* document.

**Wrapped text and row heights** (§4.7). Excel: a row with a custom height shows only the lines
that fit, and a row without one is fitted to its wrapped lines (both captured). `characters`
default: the row grows to fit.
- *Alternatives:* respect a custom row height and cut the lines that do not fit, as Excel does when
  the height is fixed; never wrap (one line, clipped or spilled).
- *Scope leaning:* document, with a row override for fixed heights.

**Line breaks in a cell without wrapping** (§4.1). Default: not drawn, the lines joined without a
space, as Excel shows them (captured).
- *Alternatives:* drawn as a space; drawn as a visible marker such as `↵`.
- *Scope leaning:* document.

**Merged ranges in formats without merges** (Markdown, CSV; §3.4). Default: the anchor's text in
its cell, the covered cells empty.
- *Alternatives:* the anchor's text repeated in every covered cell, which suits data processing (as
  a forward fill does).
- *Scope leaning:* document.

**Colors in the terminal** (§3.3). Default: automatic colors and the untinted "Text 1" are the
terminal's own foreground; fills and font colors are drawn.
- *Alternatives:* Excel's look (black on white throughout); no colors; fills only; fonts only.
- *Scope leaning:* document (a `rich` backend option).

**Separators and gridlines** (§3.4). Default: one separator character between columns, which
spilled and merged text covers, as text covers Excel's gridlines.
- *Alternatives:* honor the sheet's `show_gridlines` (a space when the sheet hides them); hide the
  separators beside a filled cell, as Excel hides gridlines there (captured in print); keep
  separators visible under spilled text.
- *Scope leaning:* document (backend).

**Markdown header and alignment** (backend options). Default: the first row is the
header, and each column's alignment comes from the first body row.
- *Alternatives:* column letters as the header (`header="letters"`); alignment from the majority of a column.
- *Scope leaning:* document.

**Column and row headings** (backend options). `column_headings` (letters) and `row_headings`
(numbers) are separate options, both off by default. They exist in every backend that can show them
(HTML, rich, text).
- *Scope:* document (backend).

**Light or dark** (`color_scheme`, in the HTML backend). An explicit argument, default light, never read from
the OS. In `excel` mode it changes only the headings, because Excel 16 for Mac keeps its cells
white in dark mode (checked by eye).
- *Alternatives:* a later HTML `auto` through CSS `prefers-color-scheme`.
- *Scope:* document.

**Exact glyph placement** (possible later option, off by default). HTML places each run with
one `letter-spacing`, which leaves glyphs inside a run up to about 2 pt from Excel's whole-point
positions. One element per glyph would place every glyph exactly, but it breaks text selection
and copying.
- *Scope leaning:* document.

**A script fitting each run to Excel's width** (measured, §7.3; not in the default
output, which is static). It corrects glyph spacing in any font. Whether to offer it
is an open question.
- *Scope leaning:* document (HTML backend).

**An on-screen scale for HTML** (planned). The page uses CSS points and pixels at 4/3 px per
point, so it is 4/3 the size Excel for Mac shows at 100 % if Excel draws one point as one screen
point (being measured). The default stays as it is. The option is a CSS custom property, such as
`--sheetpeek-scale`, applied to the wrapper, so a caller can match Excel's on-screen size without
the layout changing.
- *Scope:* document (HTML backend, or the caller's stylesheet).

**Carlito embedded in the page** (planned). A page that carries Carlito's four styles as
`@font-face` data (about 100 KB, SIL OFL, no dependency) renders Calibri text exactly anywhere,
without a local Calibri, Carlito or the scaled fallback.
- *Scope:* document (HTML backend), off by default.

**A style override for every backend** (a later design question). HTML exposes the paper, ink,
gridline and heading colors as CSS custom properties with stable class names, so a caller can
restyle a static page. Whether the model should carry such an override for the terminal and
text backends too is open.

**Colors outside `excel` mode** (to be designed later). The hazard to settle: an explicit fill
or explicit dark text shown on a dark background (a dark host page, a dark terminal). Automatic
ink already follows the terminal; an explicit `000000` font color does not.

**Excluding rows and columns** (planned). Besides `ref`, a list of rows and columns to leave out,
anywhere in the range. *Open:* whether an excluded column behaves like a hidden one, crossed by
spilled text and by merges as Excel crosses hidden columns, or is absent, so that spill stops at
it as at the edge of the range. The first keeps what Excel shows around it; the second is what
leaving data out of a report usually means.
- *Scope:* document.

**Hover text for hidden values** (planned). A `title` attribute with `full_text` on a `#` run or
on a clipped text, so the full value shows on hover. No script: the output stays static.
- *Scope:* document (HTML backend), off by default.

**Opt-in HTML scripts** (planned). The default output stays static; each script is an option:
- a copy helper that copies a cell's full value, also behind a `#` run, rather than the text
  drawn;
- selecting rows and columns to copy, as a range is copied from Excel;
- custom scripts the caller supplies.
- *Scope:* document (HTML backend).

**Frozen headings and panes** (planned). The column headings, the row headings, or the first N
rows and columns stay in place while the page scrolls, through CSS `position: sticky`. By default
the sheet's own frozen panes (`<pane>` in its `<sheetView>`) decide.
- *Scope:* document (HTML backend).

**A terminal pager** (planned). `sheetpeek view --pager` pages the output with rich's pager,
through `less -RS` so that long rows scroll sideways. Output piped elsewhere is unchanged.
- *Scope:* command line.

**A Textual viewer** (planned). An optional interactive viewer in the terminal, on its own entry
point (`sheetpeek browse`) and its own extra, with a row-count option for large sheets. sheetpeek
itself stays free of any UI; the viewer draws from the same model as the other backends.
- *Scope:* a separate entry point.

**`serve`: stale pages and live reload** (planned). A page that shows when its workbook has
changed on disk, or reloads itself: the server polls the file's modification time and pushes a
notice over Server-Sent Events, through a script that `serve` alone adds to its pages.
- *Scope:* `serve`.

**`serve` as a WSGI or ASGI app** (planned). The server's pages as an application that other
frameworks can mount under a path, with the same rule that only `.xlsx` files inside the folder
are read.
- *Scope:* `serve`.

## 10. Tests

- **Unit tests** (default `pytest`, no Excel needed): workbooks built in the test with openpyxl,
  and written as XML parts directly where cached values, missing attributes or exact XML matter.
  They cover:
  - the package lookup and the structure scan, compared with openpyxl's normal mode on the same
    files;
  - the read-only reader's fast path and its date fallback;
  - wrong or missing `<dimension>`;
  - colors;
  - each layout rule, on the characters each column draws (`tests/test_layout.py`), with the
    verified rules citing §4 and characters mode's own choices named as such;
  - each backend's output for one workbook that exercises the layout, compared with golden
    files in `tests/golden/` (regenerated on purpose with `SHEETPEEK_UPDATE_GOLDEN=1`, then read
    before committing);
  - the command line, and that only the reader imports openpyxl;
  - crafted packages (`tests/test_limits.py`), built in the test and never committed, and the
    forms a real package may take: zip64, data descriptors, a prefixed archive, stored parts;
  - that a guard's regression costs a failed test and not the machine's memory
    (`tests/watchdog.py`): a crafted case that could fail into a huge allocation runs in a child
    under a limit of 1 GiB, and the whole run under 4 GiB;
  - that no string a file controls becomes markup (`tests/test_markup.py`): one payload in each
    string, through every HTML and Markdown configuration in every layout mode and through `serve`,
    with the output parsed, Markdown through a CommonMark renderer that allows HTML.
- **Captures** (`tests/captures/*.tsv`, marker `excel`, in the default run): `layout.tsv` from
  route B, against which `tests/test_captures.py` checks the reach of every captured spill of
  digits in `characters` mode, with the rows where characters and points must differ listed as
  known deviations and their reasons, and every row's boundaries, runs and baselines in `excel`
  mode; `theme_colors`, `merge_width` and `error_width` cell by cell (§4.9). Rows for xlformat's
  own behavior are not captured again; xlformat's suites own them.
  A `--tier=fast` option can be added when a suite is large enough to need it.
- **Extras:** CI runs the suite with `rich` installed and without it. Tests needing `rich` skip via
  `pytest.importorskip`. Doc examples are doctests.
- **Latest openpyxl:** one CI job upgrades openpyxl to the newest release before testing.
- **Real-world files:** workbooks saved by Excel for Mac (later Windows and LibreOffice) under
  `tests/files/`, a few KB each. The first, `excel_formulas.xlsx`, holds Excel's own cached
  formula results (§4.8).

## 11. Roadmap

1. **Cell styles and sheet features, on openpyxl:** what openpyxl's read-only mode and the structure
   scan reach: cell styles and runs, hyperlinks, outlines, theme fonts and the palette, grid
   colors and headings, sheets and the workbook page, then rotated and stacked text.
2. **Screen captures:** the capture tool for the screen (route C, beside the PDF route of §4),
   with every existing suite rerun on Microsoft 365 as the check on the rules built from earlier
   captures.
3. **The middle tier:** filter and sort buttons, tables with the differential-format layer,
   conditional formatting, pictures and their placement, notes and comments, sheet-view parts
   (panes, zero display), print settings; then pivot tables. Most of it needs parts that
   openpyxl's read-only mode does not expose.
4. **The shared xlsx reader:** sheetpeek moves from openpyxl onto the reader now inside
   `xlrecalc`, once it is a package of its own; `_read.py` is the only module that changes
   (§8.1).
5. **Recalculation** through `xlrecalc`, an optional sibling. Charts, possibly as a package of
   their own. Planned options (§9) as they are asked for.
6. **After 1.0:** a PDF backend drawing from the `excel` mode geometry (§2.1), possibly as an
   optional extra.
7. **Last: a cross-platform phase**, which waits on having those machines and browsers to measure
   on. Refinements that stop paying on the Mac are logged here:
   - Windows Excel: column widths, row heights and glyph tables (xlformat measures the Mac only);
   - Linux viewers, and Chrome and Firefox for the HTML technique (only WebKit is measured);
   - a glyph-spacing correction for Aptos Narrow, for which no free metric-compatible font is
     known;
   - whether to offer the HTML fitting script (§9).

## 12. Prior art

Checked 2026-10-27; versions and dates from PyPI.

- **LibreOffice headless** (`soffice --headless --convert-to html`, 26.8.1, downloads of 220 to
  375 MB). Its HTML export (`sc/source/filter/html/htmlexp.cxx`):
  - writes all sheets into one file;
  - applies number formats with LibreOffice's own formatter;
  - writes merges as `colspan`/`rowspan` and widths in pixels;
  - skips hidden rows and columns.

  Concurrent runs need a separate profile per process (`-env:UserInstallation`). `unoconv` was
  archived in 2025; `unoserver` 3.7 keeps a listener running. It is a heavy external process, and
  its formatting is LibreOffice's, not Excel's.
- **xlsx2html** 0.6.4 (2026-02, MIT): openpyxl `data_only`, with its own Babel-based number
  formatter that applies only the first section of date formats.
  - Merges become spans; hidden rows and columns are dropped.
  - Widths are heuristic pixels.
  - Fonts, fills and borders are copied.
  - No `####` and no spill; HTML only.
- **xlsx2csv** 0.8.6: classifies formats as float, date or percentage and applies the user's own
  formats; CSV only.
- **xlsx2md** 0.1.1: read-only mode, raw values, no merges, widths from content.
- **markitdown** 0.1.9: `pandas.read_excel(...).to_html()`.
- **pandas** 3.1.0: `read_excel` parses values "regardless of display format", merged regions
  become NaN, and no styles, widths or hidden state are read. `Styler` styles what the user
  rebuilds by hand.
- **Terminal viewers:** VisiData 3.4 (via openpyxl), tabiew 0.16 (Rust), sc-im. None documents
  Excel number formats or merges.
- **Others:** excel2img drives a real Excel over COM (Windows only). Aspose.Cells and Spire.XLS are
  commercial engines. SheetJS Community Edition's `sheet_to_html` writes formatted text and merges
  but no styles, widths or hidden state.

What sheetpeek does that these do not:
- Excel-exact number formats through xlformat, including width-dependent `####`, `General` digit
  reduction and `*` fills.
- Text spill.
- One model shown in HTML, Markdown, terminal and plain text.
- The rendered cells as public data.
- Pure Python, with no office suite.
