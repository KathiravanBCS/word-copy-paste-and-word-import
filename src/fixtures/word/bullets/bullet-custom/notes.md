# bullets/bullet-custom

Custom glyphs from the Wingdings and Symbol picker, which are the ones a real
document actually uses once someone has been through Word's bullet library.

`\F0A8` is Wingdings 0xA8 (`▫`), `\F0D8` is Wingdings 0xD8 (`➔`), `\F0E0` is
Symbol 0xE0 (`◊`). None of them is the character the ANSI code page maps the
byte to, which is why the rendered span shows `¨`, `Ø` and `à`.

The requirement is that these are *preserved*, not normalised to a standard
bullet: the model keeps the mapped Unicode glyph, the raw byte and the font
name, and the renderer draws the glyph Word chose.
