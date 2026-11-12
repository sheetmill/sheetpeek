### Added
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

### Changed
- HTML headings take Excel's sizes: the column band 21 pt and the row band 16 pt plus 3 pt a digit
  of the last row, at Calibri 11, growing with the Normal font's size. Their colors are Excel's:
  dark `#161616`, `#ffffff` and `#565656`; light `#f7f7f7`, `#5e5e5e` and `#e0dfdd`.
- A workbook palette replaces the default entry by entry; a `<colors>` element without
  `indexedColors` keeps the default palette instead of drawing indexed colors black.
