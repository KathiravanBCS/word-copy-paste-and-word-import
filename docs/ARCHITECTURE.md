# Architecture

## The pipeline

```
raw clipboard HTML                          never mutated; kept on the model
      │
      ├─▶ stylesheet mined from raw text    @list, @font-face, @page, mso-*
      ├─▶ metadata mined from raw text      namespaces, Office XML, <meta>
      ├─▶ VML mined from raw text           lives inside conditional comments
      │
      ▼
   DOM parsed as a working clone
      │
      ├─▶ CF_HTML fragment boundary honoured
      ├─▶ security scrub
      ├─▶ comments and footnotes lifted out of the flow
      │
      ▼
   content walked into the canonical model  markers lifted before runs
      │
      ▼
   normalisation                            list identity, tables, styles, images
      │
      ▼
   renderer                                 model → HTML + CSS
```

**Cleanup is last.** Every stage before it is extraction. Nothing is discarded
before something has had the chance to understand it — that single ordering
constraint is what separates this design from a cleaner, and most of the
non-obvious code exists to maintain it.

### Why the stylesheet is mined from raw text

Word's numbering definitions look like this:

```css
@list l0:level1 {mso-level-number-format:roman-upper; mso-level-text:"%1\.";}
```

`@list` is not a CSS at-rule. A conforming CSS parser — including the browser's,
including the one behind `CSSStyleSheet` — discards every unknown at-rule at
parse time. Touch the stylesheet with a real CSS parser and the numbering
definitions are gone before you can read them.

So `WordCssTokenizer` is a deliberately tolerant hand-written tokenizer that
keeps at-rules, tolerates Word's `<!-- -->` comment wrapper, tolerates truncated
input, and does not split on a `;` inside a string. It is not trying to be a CSS
parser; it is trying to be a *lossless* one.

The same reasoning applies to VML: Word wraps it in a downlevel-hidden
conditional comment, where a browser sees a comment node and nothing else. It is
scanned out of the raw text before the DOM is consulted.

### Why the fragment boundary is trimmed on the DOM

The clipboard's `text/html` flavour is a whole document; only the part between
`<!--StartFragment-->` and `<!--EndFragment-->` is what the user selected.

Slicing the HTML string between those markers is the obvious approach and it
breaks the moment the selection starts inside a table: `<tr>…</tr>` reparsed on
its own has its rows stripped by the HTML parser. So the trim is done by
removing siblings outward from each marker in the DOM, which preserves the
enclosing table.

### Why executing elements are removed from the text

`DOMParser` documents are inert in a browser: no scripts run, no subresources
load. But the engine also runs under Node DOM shims, and not all of them honour
that — happy-dom, for one, eagerly loads `<iframe src>`.

Relying on the host parser being inert makes a security property depend on
somebody else's correctness. So `<script>`, `<iframe>`, `<embed>`, `<link>`,
`<base>` and friends are removed **from the source text**, before any DOM
implementation acts on them. The removal applies to the working copy only;
`document.rawHtml` keeps the original bytes.

They are *removed*, never renamed to an inert placeholder element. That
distinction cost a real bug and is worth stating plainly:

> `<link>` is a void element. A custom element by that name is not. Word puts
> `<link rel=File-List ...>` in the head of every clipboard payload, so a
> renamed `<link>` is hoisted into the body by a real browser, stays open,
> swallows the entire document as its children, and is then removed along with
> all of it. The symptom is a paste that yields zero paragraphs — silently,
> because the payload "parsed successfully".

Three rules follow, and each of them is a test:

- **Void elements** (`link`, `base`, `embed`, `frame`) — only the tag is
  removed. There are no children to lose.
- **Raw-text elements** (`script`, `noscript`, `noframes`, `applet`) — tag and
  contents both go; their contents are code, not document content. An unclosed
  one runs to the end of the document, which is how the HTML tokenizer reads it
  too.
- **Containers** (`iframe`, `frameset`) — the element goes, its children stay.

`<object>` is deliberately left alone: it loads nothing in an inert document,
and it is how Word represents an embedded OLE object, so removing it would
destroy the very thing the OLE diagnostic exists to report.

The same reasoning governs the DOM scrub that runs afterwards: when it removes
a forbidden element that slipped through, it unwraps the children first unless
they are raw text. Deleting a subtree on the strength of an element name is how
a whole paste disappears.

One more detail worth recording, because it is easy to reintroduce. The
attribute pattern is written so its alternatives cannot match the same text two
ways:

```
[^>"']*(?:(?:"[^"]*"|'[^']*')[^>"']*)*
```

The obvious spelling, `(?:"[^"]*"|'[^']*'|[^>])*`, is ambiguous — `[^>]` also
matches the characters inside a quoted value — and the engine explores
exponentially many partitions of the same input. That turned a 1 ms parse into
771 ms on a small fixture and hung the test suite outright on a large one.

### Why paragraph spacing is padding, not margin

Word states `spaceBefore`/`spaceAfter` as absolute values, and the gap between
two paragraphs is the *sum* of the first one's space-after and the second
one's space-before. CSS block-level margins do not work that way: two
touching margins **collapse** to the larger of the two, not their sum. A 12pt
space-after meeting a 12pt space-before renders as a 12pt gap instead of 24pt
— half of what Word declared, on every paragraph boundary in the document.

```
Word:      12pt spaceAfter + 12pt spaceBefore = 24pt gap
CSS margin: max(12pt, 12pt)                   = 12pt gap   (wrong)
CSS padding: 12pt + 12pt                      = 24pt gap   (correct)
```

Padding never collapses, so `spaceBefore`/`spaceAfter` are rendered as
`padding-top`/`padding-bottom`, and every container class that could carry a
browser's own default margin (`<blockquote>`, list items, bordered/shaded
containers) is explicitly zeroed in the base stylesheet — a UA default is just
as capable of adding an unrequested gap as a collapse is of removing one.

The same reasoning is why the standalone document shell reads its page width
and padding from the payload's own `@page` rule (`size`/`margin`) when Word
declared one, rather than assuming Letter with 1in margins. A payload with A4
paper or non-uniform margins gets its own geometry, not a guess that happens
to be right for the common case.

## The model

`WordDocument` is the source of truth. Every consumer — this renderer, an editor
adapter, a DOCX writer — reads the model and never re-interprets Word's
semantics.

```ts
interface WordDocument {
  blocks: WordBlock[];                       // the content
  styles: WordStyleSheet;                    // parsed @list / @font-face / styles
  lists: WordListDefinition[];               // numbering definitions
  images: Record<string, WordImage>;         // out of line, referenced by id
  hyperlinks: Record<string, WordHyperlink>; // out of line, referenced by id
  bookmarks: Record<string, WordBookmark>;
  diagnostics: WordDiagnostic[];             // everything not represented exactly
  metadata: WordDocumentMetadata;
  detection: WordDetectionResult;
  rawHtml: string;                           // untouched, for debugging
}
```

Three design decisions are worth stating explicitly.

**Runs are never flattened.** `Hello <b>world</b>` is two runs, not a paragraph
with a bold flag. Word nests spans several deep for a single word and each level
may change one property, so formatting is threaded down as an inherited context
and each text node captures the exact state in effect where it sits. Adjacent
runs are merged only on *structural equality of every formatting property*.

**Images, links and bookmarks live out of line.** A run references them by id.
This is what lets a bold run inside a hyperlink keep its own boundary — if links
were a property of runs, a link spanning three differently-formatted runs would
have to be either three links or one flattened run.

**Nothing in the model names an editor.** No `OUTLINE_SCHEME`, no
`RoosterListType`, no editor-specific node names. There is a test asserting
this, because it is the property that makes the model reusable and the one most
likely to erode.

## Module layout

```
src/
  model/            the canonical model — types only, no logic
  diagnostics/      diagnostic codes, collector, fidelity report
  clipboard/        DataTransfer capture, image blobs, payload shape
  detection/        weighted Word-source detection
  word/             the parsers: CSS, lists, formatting, symbols, content
  normalization/    list identity, tables, styles, images, units
  rendering/        model → HTML + CSS
  fixtures/word/    the fixture library
  tests/            detection, parsing, normalization, rendering, regression
  demo/             the clipboard lab
```

The dependency direction is strictly downward: `rendering` reads `model`,
`normalization` reads `model`, `word` writes `model`. Nothing in `model` imports
anything else.

## Detection

Word detection is weighted rather than boolean, because a single `MsoNormal`
class surviving a trip through another editor is not evidence that a payload
came from Word.

Each of the 25+ signals carries a strength — decisive, strong, moderate, weak —
and confidence saturates: each signal closes a fraction of the remaining gap to
1.0, so many weak signals can corroborate without any single one lying. A
payload is classified as Word when confidence clears the threshold **and** at
least two distinct signals fired, unless one of them is decisive (`@list`, the
`ProgId` meta, `<w:WordDocument>`), which is conclusive on its own.

Detection also distinguishes Word desktop from Word Online, Outlook, Excel and
PowerPoint, because their payloads differ in ways the parser cares about — Word
Online, for instance, emits real `<ol>`/`<ul>` and omits the rendered marker
spans entirely.

## Budgets

Clipboard content is untrusted input of unbounded size. Every walker charges
against a budget (`maxNodes`, `maxBlocks`, `maxDepth`, `maxTableDepth`,
`maxHtmlLength`) and stops descending when it is spent, raising a diagnostic
rather than throwing. A 400-page paste is legitimate; an adversarially nested
one is not, and neither should be able to wedge the tab.

## Extending it

The model is the extension point. A new back end reads `WordDocument` and emits
whatever it emits:

```ts
import { parseWordHtmlString, walkBlocks } from 'word-clipboard-engine';

const document = parseWordHtmlString(html);
for (const block of walkBlocks(document.blocks)) {
  if (block.type === 'paragraph' && block.listItem) {
    // block.listItem.marker holds the glyph or number text, its format,
    // its level text, and where it came from.
  }
}
```

Do not add editor concepts to the model to make an adapter simpler. Put them in
the adapter — that is what keeps the next adapter possible.
