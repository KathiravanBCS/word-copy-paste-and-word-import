# tables/list-in-cell

Lists inside table cells — one bulleted, one numbered, in adjacent cells.

Each cell is its own list scope. A numbered list in the second cell must not
continue the numbering of anything in the first, and both must come out as real
`<ul>`/`<ol>` structures inside their `<td>`, not as flattened paragraphs.
