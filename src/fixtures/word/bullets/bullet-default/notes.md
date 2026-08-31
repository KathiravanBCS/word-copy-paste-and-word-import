# bullets/bullet-default

The bullet golden test from the spec, and the most important fixture in the
library.

Word's three default bullet levels are not Unicode bullets. They are bytes in
legacy symbol fonts, and they reach the clipboard twice over:

| level | `mso-level-text` | rendered span | font        | means |
|-------|------------------|---------------|-------------|-------|
| 1     | `\F0B7`          | `&middot;`    | Symbol      | `•`   |
| 2     | `o`              | `o`           | Courier New | `o`   |
| 3     | `\F0A7`          | `&sect;`      | Wingdings   | `▪`   |

Taking either representation at face value is how a paste ends up showing
`·` and `§`. The expected model must show `•`, `o`, `▪` as `marker.glyph`,
with the raw byte and the font preserved beside it — and the paragraph text
must be exactly `Parent`, `Child`, `Grandchild`, with no glyph anywhere in a
text node.
