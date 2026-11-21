# sheetpeek

sheetpeek shows an Excel workbook the way Excel shows it: as HTML, Markdown, terminal output
or plain text. It is pure Python, needs no office suite, and gives programs access to the
rendered cells. Number formats are applied by
[xlformat](https://github.com/sheetmill/xlformat), whose output is tested against text captured
from Excel 16 for Mac.

Documentation: <https://sheetpeek.readthedocs.io>.

## Reading a sheet

`read` returns a `Sheet`: the values Excel shows on opening, number formats, styles, column
widths, row heights and merged ranges of one worksheet. Dates stay the serial numbers Excel stores.

```python
>>> import tempfile, pathlib, openpyxl, sheetpeek
>>> path = pathlib.Path(tempfile.mkdtemp()) / "example.xlsx"
>>> wb = openpyxl.Workbook()
>>> wb.active["A1"] = 1234.5
>>> wb.active["A1"].number_format = "#,##0.00"
>>> wb.active.merge_cells("B1:C1")
>>> wb.save(path)
>>> sheet = sheetpeek.read(path)
>>> cell = sheet.cell(1, 1)
>>> cell.value, cell.kind, cell.number_format
(1234.5, 'number', '#,##0.00')
>>> [str(m) for m in sheet.merges]
['B1:C1']
```

## Showing a sheet

`sheetpeek.render` draws a sheet with a backend. Each shows what Excel shows, as far as its
format can: number formats, `#` for a number too wide for its column, text spilling into empty
neighbors.

```python
>>> path = pathlib.Path(tempfile.mkdtemp()) / "sales.xlsx"
>>> wb = openpyxl.Workbook()
>>> ws = wb.active
>>> ws.append(["Region", "Sales", "Change"])
>>> ws.append(["North", 1234.5, 0.125])
>>> ws.append(["South", -98765.4, -0.0425])
>>> ws.append(["A note long enough to spill"])
>>> for row in (2, 3):
...     ws.cell(row, 2).number_format = "#,##0.00;[Red]-#,##0.00"
...     ws.cell(row, 3).number_format = "0.0%"
>>> wb.save(path)
>>> sales = sheetpeek.read(path)
>>> print(sheetpeek.render(sales, "text"), end="")
Region   Sales    Change
North    1,234.50    12.5%
South    ########    -4.3%
A note long enough to spill
>>> print(sheetpeek.render(sales, "markdown"), end="")
| Region | Sales | Change |
| --- | ---: | ---: |
| North | 1,234.50 | 12.5% |
| South | -98,765.40 | -4.3% |
| A note long enough to spill |  |  |
>>> print(sheetpeek.render(sales, "csv"), end="")
Region,Sales,Change
North,"1,234.50",12.5%
South,"-98,765.40",-4.3%
A note long enough to spill,,
>>> page = sheetpeek.render(sales, "html")  # a static page in Excel's own geometry
>>> page.startswith("<!doctype html>")
True
```

Fixed-width text gives every column its width in characters; Markdown and CSV have no widths,
so they show every value in full. HTML is drawn in points from Excel's layout: column widths,
row heights, each text's position and clip window, the `#` runs in the section's color (the
negative is red), and fills. `sheetpeek.render(sales, "rich")`, or `rich.print(sales)`, draws
the fixed-width layout in a terminal with fonts, colors and fills (the optional `rich` extra).

From the command line:

```bash
sheetpeek view sales.xlsx                 # rich on a terminal, plain text when piped
sheetpeek view sales.xlsx -t html > sales.html
sheetpeek serve reports/                  # every workbook under reports/ as HTML pages
```

`Sheet.layout` gives the layout itself, the data every backend draws from: in `excel` mode each
column's position and width and each text's position in points, in `characters` mode the
characters each column draws. The documentation's usage page has examples.

Workbooks are read with [xlio](https://github.com/sheetmill/xlio), in one pass over each sheet's
cells, with values as Excel reads them. A formula shows what Excel shows when it opens the file:
the cached value, or where Excel recalculates (a file another program wrote, for one), the value
[xlrecalc](https://github.com/sheetmill/xlrecalc) computes; `recalc_on_load="never"` keeps the
cached values. `ref="A1:H40"` keeps part of a sheet (the whole sheet is
still read, for its used range), and `sheetpeek.open` reads several sheets with one load.

A file from someone you do not trust is checked before anything parses it, and `limits=` sets what
a read may cost; the documentation's usage page has the details. `sheetpeek serve` has no
authentication: with `--host`, anyone who can reach the address can read the folder. It is not for
production: it is built on Python's `http.server`, which Python's documentation does not recommend
for production. For a production service, call the Python API from a web framework and pass
`limits=`, since the library's defaults are not memory-safe for untrusted files.

## Installation

```bash
pip install sheetpeek
pip install "sheetpeek[rich]"   # with styled terminal output
```

Python 3.11 or later. Depends on `xlio`, `xlformat` and `xlrecalc`. The examples here build their workbooks
with openpyxl (`pip install openpyxl`), which sheetpeek itself does not need.

## Related packages

sheetpeek is part of the [sheetmill](https://github.com/sheetmill) project, with
[xlio](https://github.com/sheetmill/xlio), which reads and writes the workbooks,
[xlformat](https://github.com/sheetmill/xlformat), which applies Excel number formats, and
[xlrecalc](https://github.com/sheetmill/xlrecalc), a formula engine.

## Development

`uv sync --group docs` then `uv run pytest`;
[CONTRIBUTING.md](https://github.com/sheetmill/sheetpeek/blob/main/CONTRIBUTING.md) has the pip route, the full
set of checks and the commit conventions.
[ARCHITECTURE.md](https://github.com/sheetmill/sheetpeek/blob/main/ARCHITECTURE.md) describes the design and
the measurements behind it.

Most of the code is written with Claude Code, reviewed before each commit and tested against
output captured from Excel; see [CONTRIBUTING.md](https://github.com/sheetmill/sheetpeek/blob/main/CONTRIBUTING.md).

## License

MIT.
