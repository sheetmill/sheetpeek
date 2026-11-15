# HTML

The `html` backend ({class}`sheetpeek.backends.HTMLBackend`) draws a sheet laid out in `excel`
mode as static HTML: no script, and nothing read from the viewer's system. Column widths, row
heights, where each text starts and where it is cut, fills and gridlines come from Excel's own
layout; spilled text lies over its neighbors' fills, as Excel draws it.

```python
>>> import tempfile, pathlib, openpyxl, sheetpeek
>>> from sheetpeek.backends import HTMLBackend
>>> path = pathlib.Path(tempfile.mkdtemp()) / "notes.xlsx"
>>> wb = openpyxl.Workbook()
>>> wb.active["A1"] = "A note long enough to spill"
>>> wb.active["B2"] = 42
>>> wb.save(path)
>>> sheet = sheetpeek.read(path)
>>> page = sheetpeek.render(sheet, "html")
>>> page.splitlines()[0]
'<!doctype html>'
>>> part = sheetpeek.render(sheet, HTMLBackend(fragment=True, row_headings=True))
>>> part.startswith("<style>")
True
>>> '<th class="sp-rowhead" scope="row">2</th>' in part
True
```

Angled text is turned with CSS's `transform: rotate()` about the start of its baseline, as
Excel turns it, and spills by the box it turns into; stacked text draws a character a line. The
render model carries the angle on each run ({attr}`sheetpeek.TextRun.angle`).

```python
>>> from openpyxl.styles import Alignment
>>> wb.active["A1"].alignment = Alignment(textRotation=45)
>>> wb.save(path)
>>> rendered = sheetpeek.read(path).layout("excel")
>>> rendered.rows[0].runs[0].angle
45.0
>>> "transform:rotate(-45deg)" in sheetpeek.render(rendered, "html")
True
```

## Options

- `fragment=True` writes the stylesheet and a `<div class="sheetpeek">` to place in a page of
  your own, instead of a whole document.
- `column_headings=True` and `row_headings=True` add Excel's column letters and row numbers.
- `color_scheme="light"` (the default) or `"dark"` colors the headings. Excel for Mac keeps its
  cells white in dark mode and changes only the headings, so the cells stay white either way.
  The scheme is never read from the viewer's system.
- `new_tab=True` opens links to web pages and e-mail addresses in a new tab
  (`target="_blank"`). It is off by default.
- `local_links=True` also links local targets (`file:` URLs, relative, UNC and drive-letter
  paths), which Excel follows; off by default (see "Links" below).
- `notices=True` adds, after the sheet, a collapsed list of what the page does not show; off by
  default.
- `buttons="static"` (the default) or `"none"`: whether to draw Excel's buttons, which do nothing
  on a static page. For now this is the outline of grouped rows and columns: a gutter before the
  row headings and above the column headings, with its bars, dots and − and + buttons, drawn at
  Excel's screen sizes in CSS pixels, and the level buttons when both headings are drawn. A sheet's
  outline is drawn only when the file shows outline symbols.
- `zoom=None` (the default) draws each sheet at the zoom saved with it, as Excel opens it; a
  whole percent from 10 to 400 overrides it. Cells, text and the headings scale; gridlines,
  border lines, pattern tiles, underlines and the outline's symbols keep their size in CSS
  pixels, as Excel keeps them in screen points. The page scales the layout Excel makes at 100 %:
  Excel lays out each zoom in whole screen pixels, so where it wraps or places text differently
  at another zoom, the page does not follow. The other outputs have no zoom.
- `freeze`: how a sheet's frozen pane is drawn. `"panes"` keeps the frozen rows and columns in
  place while the page scrolls, through CSS `position: sticky`, with no script; `"static"`
  draws only the freeze lines. By default, panes on a whole page and static in a fragment, whose
  host page decides how it scrolls. Every row and column of the range is drawn either way: where
  the file was scrolled to is the viewer's state, not content, and is not followed.

```python
>>> wb = openpyxl.Workbook()
>>> for row in range(1, 5):
...     wb.active.cell(row, 1).value = row
>>> for row in (2, 3):
...     wb.active.row_dimensions[row].outlineLevel = 1
>>> wb.save(path)
>>> '<td class="sp-ol">' in sheetpeek.render(path, "html", on_notice="ignore")
True
>>> "sp-ol" in sheetpeek.render(path, HTMLBackend(buttons="none"), on_notice="ignore")
False
>>> "--sheetpeek-zoom:1.5" in sheetpeek.render(path, HTMLBackend(zoom=150), on_notice="ignore")
True
```

## Restyling

The colors are CSS custom properties on `.sheetpeek`, and every element the backend writes has
a stable class name, so a stylesheet of your own can restyle a page and the output stays
static. **These names are public interface:** a change to them is listed in the changelog like
an API change.

| custom property | what it colors | default |
|---|---|---|
| `--sheetpeek-paper` | the cells | `#ffffff` |
| `--sheetpeek-ink` | text in the automatic color | `#000000` |
| `--sheetpeek-grid` | gridlines; a sheet with a grid color of its own sets it from the file | `#d4d4d4` |
| `--sheetpeek-heading-paper` | the headings' background | `#f7f7f7`, dark `#161616` |
| `--sheetpeek-heading-ink` | the headings' text | `#5e5e5e`, dark `#ffffff` |
| `--sheetpeek-heading-grid` | the headings' lines | `#e0dfdd`, dark `#565656` |
| `--sheetpeek-zoom` | set by the page when its zoom is not 100: the zoom as a factor, which scales the headings' font | `1` |
| `--sheetpeek-freeze` | a frozen pane's freeze lines | `#9b9b9b` |
| `--sheetpeek-outline-paper` | the outline gutter | `#eeeeee`, dark `#555555` |
| `--sheetpeek-outline-ink` | the outline's bars and dots | `#a6a6a6`, dark `#999999` |
| `--sheetpeek-outline-button` | the outline buttons' borders and strokes | `#808080`, dark `#333333` |
| `--sheetpeek-outline-button-paper` | the outline buttons' background | `#ffffff` |

| class | element |
|---|---|
| `sheetpeek` | the wrapper `<div>`, with `data-color-scheme` |
| `sp-nogrid` | on the wrapper of a sheet that hides its gridlines |
| `sp-crossed` | a cell whose left gridline text or a merge crosses |
| `sp-clip` | a text's clip window |
| `sp-run` | a text |
| `sp-center` | a centered text |
| `sp-rule`, `sp-double` | an underline or strikethrough, 1 px thick; a double underline |
| `sp-colhead`, `sp-rowhead`, `sp-corner` | the headings |
| `sp-ol`, `sp-olbutton`, `sp-ollevel` | the outline gutter's cells, its − and + buttons, its level buttons |
| `sp-fills`, `sp-borders` | the layers of pattern and gradient fills and of borders, each a box per fill or line; their own classes (`sp-f-…`, `sp-b-…`) are generated per page, as is a grid color's (`sp-g-…`, on the wrapper) |
| `sp-book` | the workbook page's wrapper `<div>`, around its `<nav>` of tabs and a `<section>` per sheet |
| `sp-first`, `sp-hidden` | on the workbook page, the tab and section shown first; a hidden sheet's tab. A tab color's class (`sp-t-…`) is generated |
| `sp-notshown` | on the workbook page, the line that says a chartsheet or dialog sheet is not shown |
| `sp-freeze` | a freeze line |
| `sp-panes`, `sp-pane`, `sp-pane-top`, `sp-pane-left`, `sp-pane-corner` | a sheet with frozen panes, and the copies of its frozen rows, columns and corner |
| `sp-notices` | the `<details>` list of notices, when the backend's `notices` option asks for it |

A cell that a link on the page goes to has the id `sp-` and its coordinate, such as `sp-C2`; on
the workbook page, `sp-`, the sheet's position and its coordinate, such as `sp-2-C2`.

```css
/* Dark cells: the file's own fills and font colors are inline styles and stay as they are,
   so text the file colors black stays black. */
.sheetpeek { --sheetpeek-paper: #1e1e1e; --sheetpeek-ink: #e6e6e6; --sheetpeek-grid: #3a3a3a; }
```

## Links

A linked cell's text is wrapped in an `<a>`, which holds the text whole, also where it spills
over its neighbors. The link adds no color or underline of its own: as in Excel, the cell's
format draws it (the Hyperlink cell style is blue and underlined, a link in the Normal style
looks like any text). Its `title` is the file's tooltip, if it has one. The text shown is the
cell's value; the hyperlink's `display` text, which Excel does not show either, is not drawn.

By default only `http`, `https` and `mailto` targets become links, with
`rel="noopener noreferrer"`. A `javascript:` or `data:` target, or any other scheme, is never
linked: it is drawn as plain text, never echoed into the page, and reported (`link.scheme`).
Local targets, which Excel follows (`file:` URLs, relative paths, UNC names and drive-letter
paths), are drawn as text and reported (`link.local`) unless `HTMLBackend(local_links=True)`,
which links them as the file writes them. Such links point into the viewer's own files and reveal
the paths the workbook names, and a target that starts with `//` (a UNC name written with forward
slashes) resolves to a network host when the page is served over HTTP, as an `http` link to that
host would. A link to a cell on the page goes to that cell; a link to another sheet or a defined
name stays an `<a>` without a target (`link.elsewhere`).

```python
>>> from openpyxl.worksheet.hyperlink import Hyperlink
>>> wb = openpyxl.Workbook()
>>> wb.active["A1"] = "Example"
>>> wb.active["A1"].hyperlink = "https://example.com/"
>>> wb.active["A2"] = "Unsafe"
>>> wb.active["A2"].hyperlink = "javascript:alert(1)"
>>> wb.active["A3"] = "Back to the top"
>>> wb.active["A3"].hyperlink = Hyperlink(ref="A3", location="A1", tooltip="First row")
>>> wb.save(path)
>>> page = sheetpeek.render(path, "html", on_notice="ignore")
>>> '<a href="https://example.com/" rel="noopener noreferrer">' in page
True
>>> "javascript" in page
False
>>> '<td id="sp-A1">' in page, '<a href="#sp-A1" title="First row">' in page
(True, True)
>>> [n.code for n in sheetpeek.notices(path, "html")]
['link.scheme']
```

## The whole workbook

{func}`sheetpeek.render_workbook` draws every sheet on one page, an option beside the one sheet a
page that {func}`sheetpeek.render` draws. Each sheet is drawn as above in a section of its own,
and a strip of tabs, as Excel's, follows them (`tabs="top"` puts it first): the visible sheets in
workbook order, each tab in its tab color (filled with white text; the active one a gradient to a
light tint of it). A chartsheet or a dialog sheet gets a tab and a line saying it is not shown.
Hidden sheets have no tab, as in Excel, unless `hidden_sheets=True`; hidden rows and columns are a
separate option, `show_hidden=True`, as for {meth}`sheetpeek.Sheet.layout`. The page has no script: a tab is a link to its sheet's section, and
in a browser with CSS's `:has()` only the chosen sheet shows. A link to a cell of another sheet on
the page goes there. Every option of the `html` backend applies.

```python
>>> wb = openpyxl.Workbook()
>>> wb.active.title = "Summary"
>>> wb.active["A1"] = "See the detail"
>>> wb.active["A1"].hyperlink = Hyperlink(ref="A1", location="Detail!B2")
>>> wb.active.sheet_properties.tabColor = "4472C4"
>>> wb.create_sheet("Detail")["B2"] = 42
>>> wb.create_sheet("Notes").sheet_state = "hidden"
>>> wb.save(path)
>>> with sheetpeek.open(path) as book:
...     [(s.name, s.state, s.tab_color) for s in book.sheets]
[('Summary', 'visible', '4472C4'), ('Detail', 'visible', None), ('Notes', 'hidden', None)]
>>> page = sheetpeek.render_workbook(path, on_notice="ignore")
>>> '<a href="#sp-1" class="sp-first sp-t-a">Summary</a><a href="#sp-2">Detail</a>' in page
True
>>> '<a href="#sp-2-B2">' in page, '<td id="sp-2-B2">' in page
(True, True)
```

Each section's id is `sp-` and the sheet's position in the workbook, from 1, and a cell a link goes
to is `sp-<n>-<cell>`. `sheetpeek view book.xlsx --all-sheets` writes the same page (`--hidden-sheets` for hidden
sheets, `--hidden` for hidden rows and columns). The text
outputs have no workbook page.

## Fonts

Excel for Mac draws each glyph on whole points, so its text is wider than a browser draws the
same font. Text in Calibri is spaced to Excel's widths with a table of Carlito's advance widths
(Carlito is metric-compatible with Calibri, under the SIL Open Font License); the package holds
the numbers only, no font. That text uses a font face whose only sources are a Calibri or
Carlito installed on the viewer's system, then a local Arial or Helvetica scaled to Calibri's
widths. Every text is anchored where Excel anchors it, by its start, its end or its center, so a
different font changes only the far end of a text, and the clip windows cut where Excel cuts.

The outline gutter's colors in the light scheme are placeholders; its dark colors and its sizes
come from Excel's screen. The heading sizes and colors and the grid color come from Excel's
screen too; the light headings' colors are provisional. The heading bands are in points, so they
scale with the page like the cells, and grow with the Normal font's size. A sheet's own grid
color draws its gridlines half a CSS pixel wide, one device pixel on a high-density display, as
Excel draws it. The text's insets and baseline come from Excel's print layout and are provisional
until they are measured on screen.
