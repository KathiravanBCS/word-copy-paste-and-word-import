# word-clipboard-engine

Microsoft Word clipboard HTML → **canonical Word Document Model** → clean HTML.

Not a "Word HTML cleaner". The distinction matters, and it is the reason this
project exists.

```
Word clipboard  →  Parser  →  Word Document Model  →  Renderer  →  HTML
                              (the source of truth)
```

A cleaner takes dirty HTML and strips things out of it, which means every
decision is made in terms of markup that has already lost the information the
decision needs. This engine does the opposite: it **reads Word's own
declarations first**, builds a model that states what the document *means*, and
only then produces output. Cleanup happens last, on a model that already
understands what it is looking at.

Zero runtime dependencies. No editor library — not RoosterJS, CKEditor, Quill,
Froala or TinyMCE, and no commercial paste-from-Office product. The model knows
nothing about any editor, which is what makes it reusable by all of them.

## The problem

Paste this from Word:

```
• Parent
   o Child
      § Grandchild
```

and a naive implementation gives you three paragraphs whose text is
`"·      Parent"`, `"o      Child"`, `"§      Grandchild"`. The bullets have
become text: unselectable as a list, wrong when re-indented, wrong when
re-numbered, and wrong glyphs into the bargain.

That happens because Word does not send `<ul><li>`. It sends this:

```html
<p class=MsoListParagraphCxSpFirst style='text-indent:-.25in;mso-list:l0 level1 lfo1'>
  <![if !supportLists]>
    <span style='font-family:Symbol'>
      <span style='mso-list:Ignore'>&middot;<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp; </span></span>
    </span>
  <![endif]>
  Parent<o:p></o:p>
</p>
```

Three separate facts are hiding in there, and all three are needed:

| where | what it says |
|---|---|
| `mso-list:l0 level1 lfo1` | this paragraph is item at level 1 of list `l0` |
| `<![if !supportLists]>…` | the marker Word *drew*, as literal text |
| `@list l0:level1 {mso-level-text:\F0B7; font-family:Symbol}` | what the marker *means* |

And `·` is not a bullet. It is byte `0xB7` of the **Symbol** font, which is
`•`. `§` is byte `0xA7` of **Wingdings**, which is `▪`. The engine reads the
`@list` rule, decodes the font's code page, lifts the rendered marker out of
the content *before any text run is created*, and emits:

```html
<!-- data-word-* / style attributes elided for readability -->
<ul class="wce-list wce-1">
  <li><span class="wce-marker">•</span>Parent
    <ul class="wce-list wce-2">
      <li><span class="wce-marker">o</span>Child
        <ul class="wce-list wce-1"><li><span class="wce-marker">▪</span>Grandchild</li></ul>
      </li>
    </ul>
  </li>
</ul>
```

```css
.wce-1 > li { padding-left: 24px; text-indent: -24px; }
.wce-1 > li > .wce-marker { display: inline-block; min-width: 24px; }
```

A real list, at Word's own indentation, with the marker in its own element —
never text — and positioned the way Word actually positions it: a fixed-width
gutter with the marker flush left and a real gap before the paragraph text,
not a browser counter-style marker crowded up against it. That distinction is
verified, not assumed: see
[LIST-PARSING.md § 6](docs/LIST-PARSING.md#6-rendering-the-marker).

## Install

```bash
npm install word-clipboard-engine
```

## Use

```ts
import {
  captureFromPasteEvent,
  parseWordClipboard,
  renderWordDocument,
} from 'word-clipboard-engine';

editor.addEventListener('paste', (event) => {
  event.preventDefault();

  const payload  = captureFromPasteEvent(event);   // every clipboard flavour
  const document = parseWordClipboard(payload);    // the canonical model
  const { html, css } = renderWordDocument(document);

  applyStylesheetOnce(css);
  insertHtml(html);

  for (const diagnostic of document.diagnostics) {
    console.info(diagnostic.fidelity, diagnostic.code, diagnostic.message);
  }
});
```

Parse a string instead (a server, a test, a saved `.htm` file):

```ts
import { parseWordHtmlString } from 'word-clipboard-engine';
const document = parseWordHtmlString(html);
```

Produce a complete, self-contained `.html` file — the "download as HTML" case,
with the generated counter styles in the head so the native markers survive:

```ts
import { renderStandaloneHtml, suggestFileName } from 'word-clipboard-engine';

const { document: file } = renderStandaloneHtml(document);
download(new Blob([file], { type: 'text/html;charset=utf-8' }), suggestFileName(document));
```

## What it preserves

| Word construct | How |
|---|---|
| Paragraphs, headings | `<p>`, `<h1>`–`<h6>` from style name / class / element, never from font size |
| Character formatting | run by run — bold, italic, underline styles, strike, colour, highlight, font, size, super/sub, spacing, small caps, language |
| Bullets | positioned marker element at Word's own gutter width from Word's `@list` rule (a real `@counter-style` is available as an opt-in); symbol fonts decoded to Unicode with the raw byte and font kept |
| Numbering | Word's own format and level text; `%1.%2` stays `1.1` |
| Multilevel lists | arbitrary depth, per-level format, mixed number/bullet levels |
| List identity | same / nested / continued / restarted, from `mso-list` and `lfo` |
| Start values | `mso-level-start-at` → `<li value>` and `<ol start>` |
| Indentation | Word's hanging indent converted to CSS list geometry, per level, relative to the parent |
| Tables | resolved grid, merged cells, nested tables, borders, shading, `<thead>`, percentage column widths |
| Images | data URIs and clipboard blobs resolved; local `file:///` references become labelled placeholders |
| Hyperlinks & bookmarks | out of line in the model, so a bold run inside a link survives |
| Page & section breaks | `page-break` blocks, with a diagnostic for what page setup could not carry over |
| Paragraph spacing | additive, matching Word — rendered as padding, never as CSS margin, which would collapse two touching gaps into one |
| Page geometry | the downloaded standalone file uses the payload's own `@page` size/margins, not a Letter-with-1in guess |

## What it does not do silently

Everything that cannot be represented exactly produces a diagnostic with a
fidelity class:

- `EXACT` — represented without loss
- `EQUIVALENT` — a different mechanism, same result
- `APPROXIMATED` — visibly close, not identical
- `UNSUPPORTED` — no HTML equivalent; preserved on the model, not rendered

```ts
import { getFidelityReport, formatFidelityReport } from 'word-clipboard-engine';
console.log(formatFidelityReport(getFidelityReport(document)));
```

```
Word detected: true (confidence 1.00)
Blocks 20  paragraphs 19  headings 2  runs 23
Lists 2 definitions / 2 instances / 6 items (2 bullet, 4 numbered, max depth 2)
Tables 1 (0 nested, 0 merged cells)
Images 0 (0 resolved, 0 unresolved)
Hyperlinks 1  bookmarks 0  page breaks 0
Fidelity: EXACT 2, EQUIVALENT 3, APPROXIMATED 0, UNSUPPORTED 0
```

## Clipboard Lab

```bash
npm run dev
```

Three panes: the untouched clipboard payload, the model the parser built from
it, and the HTML the renderer produced from that model. When a paste comes out
wrong, the question is always *which stage* got it wrong, and a preview alone
cannot answer that.

It also loads any of the 21 fixtures, opens a saved `.htm` file, downloads the
standalone HTML, and exports whatever you pasted as a new fixture.

## Development

```bash
npm test              # 225 tests
npm run typecheck
npm run lint
npm run dev           # clipboard lab at :5180
npm run build         # library → dist/
UPDATE_FIXTURES=1 npm test   # re-bless golden fixtures (read the diff first)
```

## Documentation

| | |
|---|---|
| [ARCHITECTURE.md](docs/ARCHITECTURE.md) | the pipeline, the model, and why the order is what it is |
| [WORD-CLIPBOARD.md](docs/WORD-CLIPBOARD.md) | what Word actually puts on the clipboard, in detail |
| [LIST-PARSING.md](docs/LIST-PARSING.md) | the hardest part: markers, numbering, identity, indentation |
| [FIDELITY.md](docs/FIDELITY.md) | every diagnostic code and what it means |
| [TESTING.md](docs/TESTING.md) | fixtures, golden tests, and capturing a real Word payload |

## Licence

MIT
