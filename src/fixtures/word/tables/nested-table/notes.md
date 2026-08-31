# tables/nested-table

A table inside a cell, with paragraphs either side of it.

The requirement is that the inner table is a real `Table` in the model at
`depth: 1` — not markup passed through — so a consumer can walk into it, and so
its own grid, borders and widths are resolved independently of the outer one.
