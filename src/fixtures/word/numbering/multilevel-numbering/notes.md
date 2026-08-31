# numbering/multilevel-numbering

The golden test the spec calls out by name:

```
I. Background
   1.1 Orbis India
   1.2 We understand
   1.3 VSTN
   1.4 Given
```

must come out as exactly that. Never `Article I`, never `Section 1.01` — those
are an application numbering scheme overriding Word's, which is the specific
failure this whole project exists to avoid.

Word's definition is the authority: level 1 is `roman-upper` with level text
`%1.`, level 2 is decimal with level text `%1.%2`. `%1.%2` has no CSS
counter-style equivalent, so the renderer keeps the literal marker Word drew
and puts it in a real `::marker`. It never re-derives it from a scheme of its
own.
