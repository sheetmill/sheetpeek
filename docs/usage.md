# Usage

## Reading a sheet

{func}`sheetpeek.read` returns a {class}`sheetpeek.Sheet`: what the file says about one worksheet,
before any layout. The examples on this page use a small workbook written with openpyxl.

```python
>>> import tempfile, pathlib, openpyxl, sheetpeek
>>> path = pathlib.Path(tempfile.mkdtemp()) / "sales.xlsx"
>>> wb = openpyxl.Workbook()
>>> ws = wb.active
>>> ws.append(["Region", "Sales"])
>>> ws.append(["North", 1234567.891])
>>> ws.append(["A note long enough to spill"])
>>> ws["B2"].number_format = "#,##0.00"
>>> wb.save(path)
>>> sheet = sheetpeek.read(path)
>>> sheet.cell(2, 2).value, sheet.cell(2, 2).number_format
(1234567.891, '#,##0.00')
```

## Showing it

{func}`sheetpeek.render` shows a sheet with a backend: `text`, `rich`, `html`, `markdown`,
`csv` or `tsv`. Fixed-width output gives every column its width in characters, as Excel lays it out: a
number that does not fit shows `#`, and text spills into empty neighbors.

```python
>>> print(sheetpeek.render(sheet, "text"))
Region   Sales
North    ########
A note long enough to spill
<BLANKLINE>
>>> print(sheetpeek.render(sheet, "markdown"))
| Region | Sales |
| --- | ---: |
| North | 1,234,567.89 |
| A note long enough to spill |  |
<BLANKLINE>
```

Markdown and CSV have no column widths, so every cell shows its full text. Text that spills past
the last used column runs on into empty columns added for it, as it does in Excel; the layout
marks them as added.

## The rendered cells

{meth}`sheetpeek.Sheet.layout` gives the layout itself, the same data every backend draws from.
In `characters` mode each cell records the characters drawn in its column and the cell they come
from.

```python
>>> rendered = sheet.layout("characters")
>>> cell = rendered.cell(2, 2)
>>> cell.text, cell.full_text, cell.overflow, cell.align
('########', '1,234,567.89', True, 'right')
>>> spilled = rendered.cell(3, 2)
>>> spilled.lines, spilled.source
(('ng enoug',), (3, 1))
```

## Excel's geometry

In `excel` mode the layout is in points, as Excel for Mac lays the sheet out: each column's x
and width, each row's y and height, and each text as a {class}`sheetpeek.TextRun` with its
start, baseline, width and clip window. A number shows as many `#` as Excel shows.

```python
>>> excel = sheet.layout("excel")
>>> [(c.letter, c.x_pt, c.width_pt) for c in excel.columns]
[('A', 0.0, 53.0), ('B', 53.0, 53.0), ('C', 106.0, 53.0)]
>>> run = excel.rows[2].runs[0]
>>> run.text, run.x, run.y, run.clip
('A note long enough to spill', 3.0, 42.0, (1.0, 31.0, 159.0, 45.0))
>>> excel.rows[2].crossed
(True, True)
>>> excel.cell(2, 2).text
'#########'
```

Column C was added for the spill: it lies past the last used column, where the sheet is empty.

## HTML

The `html` backend draws that geometry as a static page, or as a fragment to embed; the
{doc}`html` page describes its options and the CSS names a stylesheet can restyle.

```python
>>> page = sheetpeek.render(sheet, "html")
>>> page.splitlines()[0]
'<!doctype html>'
```

## What is not shown

sheetpeek never drops silently what a workbook uses. Whatever a rendering leaves out or
approximates is reported as a {class}`~sheetpeek.Notice`, with a code, an outcome (`omitted` or
`approximated`), a reason, the cells, and the output that decided it. The reasons are
`not_implemented` (sheetpeek does not draw it yet), `not_representable` (this output cannot),
`needs_dependency`, `option` (an option you gave), `unsafe` (refused on purpose, such as a linked
picture, which Excel fetches and sheetpeek never does) and `file`. {func}`sheetpeek.notices` lists
them without rendering:

```python
>>> from openpyxl.styles import Border, Font, Side
>>> wb = openpyxl.Workbook()
>>> wb.active["A1"] = "Total"
>>> wb.active["A1"].font = Font(bold=True)
>>> wb.active["A1"].border = Border(bottom=Side("thin"))
>>> wb.save(path)
>>> for notice in sheetpeek.notices(path, "csv"):
...     print(notice.code, "|", notice.message)
border | borders not shown (1 cell): not implemented yet
font.bold | bold text not shown in csv (1 cell): csv cannot show it
>>> [notice.code for notice in sheetpeek.notices(path, "html")]
['border']
```

A file is reported for what it **uses**, not what it contains: page margins, page setup,
protection, validation and the like draw nothing on the sheet and are never reported, and a plain
workbook gives no notice at all. Anything a sheet holds that sheetpeek does not recognize is
reported as `sheet.unknown_element` or `sheet.unknown_extension`.

{func}`sheetpeek.render` issues one {class}`~sheetpeek.NotShownWarning` per call when there is
something to report. Filter it as any warning (`warnings.simplefilter("error",
sheetpeek.NotShownWarning)` makes it an error), or pass `on_notice`: a callable that receives each
notice instead, or `"ignore"`.

```python
>>> found = []
>>> output = sheetpeek.render(path, "csv", on_notice=found.append)
>>> [notice.matches("font") for notice in found]
[False, True]
```

Codes are dotted, and {meth}`~sheetpeek.Notice.matches` takes a code or the area above it
(`"font"` matches `"font.bold"`). {data}`sheetpeek.FEATURES` lists every code. A backend of your
own declares what it shows in a `support` mapping from code to {class}`~sheetpeek.Support`; one
without it has every feature reported as not shown.

## From the command line

```bash
sheetpeek view sales.xlsx                 # rich on a terminal, plain text when piped
sheetpeek view sales.xlsx --column-headings --row-headings   # letters and numbers
sheetpeek view sales.xlsx -t markdown     # or html, csv, tsv, text, rich
sheetpeek view sales.xlsx --list          # the sheet names
sheetpeek view sales.xlsx -t csv --notices full     # what is not shown, with its cells
sheetpeek view sales.xlsx --strict --suppress font  # exit status 3 if anything else is not shown
sheetpeek serve reports/                  # every workbook under reports/ as HTML pages
```

`sheetpeek view` writes what the output does not show to standard error, one line per kind, the
code first: `--notices summary` (the default), `full` (with the cells) or `none`. `--suppress
CODE` leaves out a code and the codes under it, and `--strict` exits with status 3 when anything
is left to report. `sheetpeek serve` lists it at the foot of each page.

`sheetpeek serve` is a small read-only web server, from the standard library: an index of the
`.xlsx` files under the folder, and one page per sheet. It binds to 127.0.0.1 unless `--host`
says otherwise, and it reads only `.xlsx` files inside the folder.

**`serve` has no authentication.** Bound to 127.0.0.1, it answers only to `localhost`, `127.0.0.1`
and `[::1]`, which keeps web pages in your browser from reading the folder's sheets. With `--host`,
anyone who can reach the address can read every workbook in the folder. It reads two workbooks at
a time and answers 503 to more.

**`serve` is not for production.** It is built on Python's `http.server`, which Python's
documentation does not recommend for production, and its safeguards do not make it a production
server. For a production service, call sheetpeek's Python API from a web framework and pass
`limits=`: the library's defaults are not memory-safe for untrusted files (see below).

## Files from people you do not trust

A workbook is a zip of XML, and a small one can ask for a great deal. Before anything parses a
file, sheetpeek refuses a part that declares a DTD, sizes in the archive that are not the parts'
sizes, compression methods and encryption an xlsx package does not use, and a part named twice, all
with {class}`sheetpeek.PackageError`. {class}`sheetpeek.ReadLimits` sets
what a read may cost, and going over it raises {class}`sheetpeek.LimitError`, a kind of
`PackageError`, which is a `ValueError`.

```python
>>> limits = sheetpeek.ReadLimits(max_cells=4)
>>> sheetpeek.read(path, ref="A1:B3", limits=limits)
Traceback (most recent call last):
sheetpeek...LimitError: A1:B3 is 6 cells, more than the limit of 4: read a smaller range
```

`max_cells` is the area of the grid read, rows times columns, whether the cells hold anything or
not. It applies to `ref` when one is given and otherwise to the sheet's used range, and again to
the layout, which adds columns for spilled text. The defaults refuse only what could hardly be read
anyway: they do not make reading safe on a small machine, because some parts are held in memory
(a part of markup read whole can take about 25 times its size). A service should lower them, as
`sheetpeek serve` does:

| limit | `read`, `open`, `view` | `serve` |
|---|---|---|
| `max_parts` | 100,000 | 10,000 |
| `max_part_bytes` | 4 GiB | 256 MiB |
| `max_total_bytes` | 16 GiB | 1 GiB |
| `max_cells` | 100,000,000 | 1,048,576 |

Both commands take `--max-parts`, `--max-part-bytes`, `--max-total-bytes` and `--max-cells`.

CSV and TSV output keeps a text such as `=1+1` as it is, because it is the cell's text. A
spreadsheet opening that file may treat it as a formula, so be careful with output made from a file
you do not trust.
