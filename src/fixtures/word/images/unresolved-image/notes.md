# images/unresolved-image

The common case, and the one most implementations get quietly wrong.

Word emits the picture twice: as a VML shape inside a downlevel-*hidden*
conditional comment (which no browser parses into elements), and as an `<img>`
inside a downlevel-*revealed* one. Both point at a temp file that exists only on
the machine the content was copied from.

Expected behaviour: the VML twin is mined for its title and size, the `<img>`
is recognised as unresolvable, a `WORD_LOCAL_FILE_IMAGE` and a
`WORD_UNRESOLVED_IMAGE` diagnostic are raised, and the renderer emits a
labelled placeholder — never a broken `<img src="file:///…">`.
