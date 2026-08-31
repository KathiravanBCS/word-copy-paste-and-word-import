# formatting/character-formatting

The formatting golden test from the spec. Every run boundary must survive:
`Hello <b><span color>world</span></b> and <i>goodbye</i>.` is five runs, not
one paragraph with flags.

Also covers the two spellings Word uses for the same thing — presentational
elements (`<b>`, `<i>`, `<u>`, `<s>`, `<sup>`, `<sub>`) and CSS declarations —
and the fact that `background:yellow` on a span is the highlighter pen.
