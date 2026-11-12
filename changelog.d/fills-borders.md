### Added

- Pattern fills: the 17 patterns are drawn in HTML as Excel draws them, from tiles captured on
  the screen, anchored at the sheet's origin and fixed in screen pixels.
- Gradient fills: linear gradients at Excel's angle, which is taken in the unit cell, and path
  gradients in rectangular bands.
- Borders: the four sides in all 13 styles and their colors, and diagonals, drawn in HTML as a
  layer of lines in CSS px. Where two cells' sides meet, the layout keeps the one Excel draws
  (the stronger style, then the darker color), keeps each cell's own sides around a merge, and
  drops the vertical sides that spilled text crosses. `slantDashDot` is drawn as
  `mediumDashDot`, and reported.
- `CellStyle.fill_spec` (a `Pattern` or a `Gradient`) and `CellStyle.borders` (a `Borders` of
  `Side`s), and `RenderedSheet.extra["borders"]`, a `RenderedBorders` of the resolved lines in
  `excel` mode. New public types: `Pattern`, `Gradient`, `Fill`, `PatternKind`, `Borders`, `Side`,
  `BorderStyle`, `RenderedBorders`, `Diagonal`.
- Census codes `fill.pattern`, `fill.gradient`, `border`, `border.diagonal` and
  `border.slant_dash_dot`, declared by every output: `rich` shows a pattern as one blended color
  and a gradient as its midpoint color; borders are drawn only in HTML so far.
- The HTML classes `sp-fills` and `sp-borders`.

### Changed

- `CellStyle.fill` of a pattern fill is the blend of its two colors by the pattern's share,
  not its foreground: a `gray0625` fill, 6.25 % black dots, was shown as solid black. A
  gradient's is its midpoint color, where it was no fill.
- A solid fill in indexed color 64 is no fill, and a fill index past the palette is black, as
  Excel draws them; an index past the palette was no fill.
- `border`, `border.diagonal`, `fill.pattern` and `fill.gradient` are no longer reported by the
  reader for every output; each output reports them from the census.
