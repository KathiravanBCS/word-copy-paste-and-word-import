# basic/headings

Heading detection from three independent signals: the Word style name
(`mso-style-name:"Heading 1"` on the `MsoHeading1` class), the class name
itself, and a genuine `<h3>` element.

`Title` is deliberately *not* a heading: Word's Title style is a display style,
not an outline level, and promoting it to `<h1>` would invent structure the
document does not have. It stays a paragraph carrying its style name.
