# tables/simple-table

A Table Grid with a shaded header row, per-cell borders and column widths.

`mso-yfti-irow:-1` on the first row is Word's marker for a repeating header
row: it becomes a real `<thead>` with `<th scope="col">`, which is what makes
the table navigable rather than a grid of anonymous cells.

Note also that Word states widths twice — a pixel `width` attribute and a point
`width` in the style — and that the borders are split across the table rule,
each cell, and `mso-border-*-alt`.
