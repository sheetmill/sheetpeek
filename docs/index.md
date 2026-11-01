# sheetpeek

sheetpeek shows an Excel workbook the way Excel shows it: as HTML, Markdown, terminal output
or plain text. It is pure Python, needs no office suite, and gives programs access to the
rendered cells. Number formats are applied by `xlformat`.

**Status: in development, not released; 0.1.0 is being prepared.** Reading, layout in
Excel's own geometry and in fixed-width characters, the text, `rich`, HTML, Markdown and CSV
backends, `sheetpeek view` and `sheetpeek serve` work.

Python 3.11 or later. The source is at <https://github.com/sheetmill/sheetpeek>.

sheetpeek is part of the [sheetmill](https://github.com/sheetmill) project, with
[xlformat](https://github.com/sheetmill/xlformat), which applies Excel number formats. A third
package, a formula engine, is in development.

```{toctree}
:maxdepth: 2

usage
html
api
changelog
```
