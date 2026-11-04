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
- **Degrade, never fail.** A backend that cannot show a feature ignores it, as Markdown does
  with merges and delimited text with widths. A number format Excel would refuse falls back to
  `General`. A Normal font without glyph tables is approximated, with a warning.
  The exception is a package refused as unsafe or a request over a limit (§5.5): those raise
  `PackageError` or `LimitError`, since nothing can be shown of a file that may not be read.
- **Excel is the reference.** Correct means what Excel 16 for Mac shows. ECMA-376 is secondary;
  LibreOffice is a lower-trust secondary reference.
- Out of scope: conditional formatting, charts, images, comments, editing, any interactive UI.

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
Calibri 11), not 8.43. A column gets `round(units × scale)` characters, minimum 1, capped by
`max_width` (`Sheet.layout` takes both). At Calibri 11:

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
| `h_align` | `"general" \| "left" \| "center" \| "right" \| "fill" \| "justify" \| "distributed" \| "center_across"` | `cell.alignment.horizontal` | as stored; `centerContinuous` is `center_across` |
| `v_align` | `"top" \| "center" \| "bottom" \| "justify" \| "distributed"` | `cell.alignment.vertical`, default `bottom` | |
| `wrap` | `bool` | `cell.alignment.wrap_text` | |
| `indent` | `int` | `cell.alignment.indent` | one indent level is one digit's width, 6 pt at Calibri 11 (§4.6) |
| `hyperlink` | `str \| None` | the structure scan's `<hyperlinks>` (§5.4) | target or location |
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

**Borders are not read.** They are cheap to read, but drawing them well differs in every
backend, and gridlines already give a table its structure. A `borders` field with four
`(style, color)` sides is the natural shape when they are added. Rich-text runs inside one cell
are flattened to the cell's font.

### 3.2 What a sheet carries

| field | source |
|---|---|
| `name`, `state` (`visible`, `hidden`, `very_hidden`) | `ws.title`, `ws.sheet_state` |
| `used_range` | the structure scan: bounding box of every `<c>` element, from `<row r>` and `<c r>` (positions inferred when `r` is omitted, as the schema allows), widened to cover every merged range, since a merge occupies its cells even where the file has no `<c>` for them; `<dimension>` is never trusted |
| `ref` | the requested range, or `used_range` |
| `merges` | scan: `<mergeCells>` |
| `columns`: per column `width`, `hidden` | scan: `<cols>` (each `<col>` covers `min` to `max`) |
| `rows`: per row `height` in points, `hidden` | scan: `ht`, `hidden` of `<row>` |
| `default_column_width`, `default_row_height` | scan: `<sheetFormatPr>` (`defaultColWidth`, `baseColWidth`, `defaultRowHeight`) |
| `show_gridlines` | scan: `<sheetView showGridLines>` |
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
  palette (`openpyxl.styles.colors.COLOR_INDEX`). Index 64 is the system foreground color, which
  means automatic.
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
palette recolors those in Excel. sheetpeek ignores that and documents it.

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
- `scale` and `max_width` raise `ValueError` in `excel` mode.

`tests/test_captures.py` lays the `layout` workbook out in `excel` mode too: every row's crossed
and fill-covered boundaries against the missing gridlines, every captured run's start (exact) and
end (within a point), color and baseline.

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

### 4.8 How Excel saves formula results

**Captured** (`capture/excel_saved.py`; the saved file is `tests/files/excel_formulas.xlsx`):
Excel writes a formula returning `""` as `<c t="str"><f>""</f><v/></c>`, `t="str"` with an empty
value element. The structure scan treats that form as the empty string, which stops a spill;
openpyxl alone reads it as an empty cell. A formula without a cached value (as openpyxl writes it)
has no `t`. The same file confirms Excel stores `=DATE(1900,2,29)` as serial 60, which the reader
keeps (§5.3).

### 4.9 Captures

Captured so far: `theme_colors` (§3.3), `merge_width` and `error_width` (§4.2, §4.4, no deviation), merged
text wider than its merge (§4.1), and row autofit on open (§4.7, suite `row_height`). What
waits for the screen capture: the gridline and heading colors, the heading sizes, the text insets
and baselines, which come from print, and the 15 pt rows that print 14 pt apart.

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
- `<mergeCells>` and `<hyperlinks>`.

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
  `sans-serif` fallback and are not spaced. Kerning and ligatures are off, as in Excel.
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
  colors and the gridline color are placeholders until the screen capture.
- **The CSS surface is public interface**, changed only with a changelog entry: the custom
  properties `--sheetpeek-paper`, `--sheetpeek-ink`, `--sheetpeek-grid`,
  `--sheetpeek-heading-paper`, `--sheetpeek-heading-ink`, `--sheetpeek-heading-grid`, and the
  classes `sheetpeek`, `sp-crossed`, `sp-clip`, `sp-run`, `sp-center`, `sp-colhead`,
  `sp-rowhead`, `sp-corner`, `sp-nogrid`. Explicit colors from the file (fills, font and section
  colors) are inline styles, so restyled dark paper keeps a cell's explicit dark text (§9).
- **Page or fragment:** a page is a standalone document titled with the sheet's name; a
  fragment is the stylesheet and the `<div class="sheetpeek">`, for a host page.
- **Size** (`bench/html_size.py`, 5,000 rows × 20 columns, 100,000 cells and 97,144 runs, Apple
  M4 Pro): 11.1 MB of HTML, 111 bytes per cell, 0.58 MB gzipped; read 0.55 s, layout 1.90 s,
  render 0.29 s, peak RSS 245 MB. The first version, with a clip span and a `calc()` baseline on
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
  _rendered.py       RenderedSheet, RenderedRow, RenderedCell, RenderedColumn, Gap
  _package.py        ReadLimits, the check before reading (§5.5), sheet name → worksheet part
  _structure.py      the structure scan (§5.4)
  _read.py           the reader: openpyxl read-only, date conversion off; builds a Sheet
  _colors.py         theme parsing, tint, indexed palette
  _layout.py         width modes, the xlformat seam, alignment, spill, merges, wrapping
  _excel.py          excel mode: geometry in points, text runs, clips
  _rows.py           row heights (§4.7) and baselines
  backends/
    __init__.py      Backend protocol, registry, render()
    text.py  rich.py  markdown.py  delimited.py  html.py
  cli.py             `sheetpeek view`, `sheetpeek serve`
  _serve.py          the server behind `sheetpeek serve`
  _carlito.py        Carlito's advance widths, generated by tools/carlito_widths.py
  __main__.py
tests/               unit tests; captures/ (TSV) and their loader
capture/             capture tooling (not in the wheel)
tools/               the generator of _carlito.py (not in the wheel)
bench/               scripts that rerun §5
docs/                Sphinx, as in xlformat
```

**The reader is isolated.** Only `_read.py` imports openpyxl (a test enforces it), and only the
package's entry points import `_read.py`. Everything else works on sheetpeek's own model, so a
later reader can replace the openpyxl one without touching layout or backends. `_structure.py`
reads the worksheet XML and knows nothing of layout.

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
  [--max-parts N] [--max-part-bytes N] [--max-total-bytes N] [--max-cells N]`. The default is
  `rich` on a terminal and `text` when piped or when `rich` is not installed. Warnings go to
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

**Column and row headings**.
`column_headings` (letters) and `row_headings` (numbers) are separate options, both off by
default. They exist in every backend that can show them (HTML, rich, text).
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

1. **By 1.0:** automated screen capture of the test sheets, checked against the PDF route on a
   sample (§4). Planned options (§9) as they are asked for.
2. **Later: a cross-platform phase.** Refinements that stop paying on the Mac are logged here:
   - Windows Excel: column widths, row heights and glyph tables (xlformat measures the Mac only);
   - Linux viewers, and Chrome and Firefox for the HTML technique (only WebKit is measured);
   - a glyph-spacing correction for Aptos Narrow, for which no free metric-compatible font is
     known;
   - whether to offer the HTML fitting script (§9).
3. **A later 0.x minor:** live recalculation through `xlrecalc`, an optional sibling.
4. **After 1.0:** a PDF backend drawing from the `excel` mode geometry (§2.1), possibly as an
   optional extra.

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
