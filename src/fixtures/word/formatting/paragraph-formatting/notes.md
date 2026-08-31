# formatting/paragraph-formatting

Alignment, the four margins, hanging indent, line spacing, paragraph borders
and shading, tab stops, and an explicit page break.

The `mso-tab-count:1` span is the case worth watching: Word writes a tab as a
run of non-breaking spaces inside a span that says how many tabs it stands for.
It must become a tab run, not six spaces of text.
