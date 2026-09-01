# List parsing

The hardest part of the problem, and the part everything else is arranged
around. Four separate questions have to be answered, and they have four separate
sources.

| question | source |
|---|---|
| Is this paragraph a list item, and at what level? | `mso-list:l0 level1 lfo1` |
| What did Word *draw* as the marker? | `<![if !supportLists]>…<![endif]>` |
| What does that marker *mean*? | `@list l0:level1 { … }` |
| How does this item relate to its neighbours? | adjacency + `(listId, lfo)` |

## 1. Lifting the marker out — before anything else

The order in `parseParagraph` is fixed and is the whole ballgame:

1. Resolve the paragraph's formatting.
2. Decide whether it is a list item (`mso-list`).
3. **Lift the rendered marker out of the content.**
4. *Then* parse the remaining inline content into runs.

Reversing 3 and 4 is exactly the bug that produces `"• Item"` as paragraph text.
By the time a run exists, the glyph is inside a text node and every downstream
consumer inherits the mistake.

Two markup shapes are handled, because Word emits both:

```html
<!-- bracketed -->
<![if !supportLists]><span style='font-family:Symbol'><span style='mso-list:Ignore'>·<span
style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp; </span></span></span><![endif]>

<!-- bare -->
<span style='mso-list:Ignore'>1.</span>
```

For the bare form the parser walks up from the `mso-list:Ignore` span, removing
the outermost wrapper that exists only to hold the marker, so no empty span is
left to produce a phantom run. The font is read from those wrappers, skipping
the innermost small-font spacer span — that span is padding to the tab stop, and
mistaking its `font:7.0pt "Times New Roman"` for the marker's font would break
glyph decoding.

## 2. Decoding symbol-font bullets

Word's three default bullet levels are not Unicode bullets. They are bytes in
legacy symbol fonts, and they arrive twice over:

| level | `mso-level-text` | rendered span | font | means |
|---|---|---|---|---|
| 1 | `\F0B7` | `&middot;` (`·`) | Symbol | `•` |
| 2 | `o` | `o` | Courier New | `o` |
| 3 | `\F0A7` | `&sect;` (`§`) | Wingdings | `▪` |

Neither representation is the character. `\F0B7` is the byte `0xB7` lifted into
the Unicode private use area at `U+F0B7`; `·` is what the ANSI code page maps
that byte to. The actual glyph is Symbol's bullet, `U+2022`.

`resolveSymbolGlyph(glyph, font)` handles both: it strips the private-use offset
when present, then looks the byte up in the font's code page table (Symbol,
Wingdings, Wingdings 2 and 3, Webdings, Dingbats). What comes back is:

```ts
{
  glyph:    '•',        // the Unicode equivalent
  rawGlyph: '',   // exactly what Word sent
  font:     'Symbol',   // the font it was drawn in
  mapped:   true,       // a mapping was found
  codePoint: 0xb7,
}
```

All of it is kept on the marker. A byte with no known mapping sets
`unmapped: true`, keeps the raw glyph, and is rendered *in its original font*
rather than being replaced by a guess — with a `WORD_SYMBOL_FONT_UNMAPPED`
diagnostic saying so.

Custom bullets from Word's picker work the same way: `\F0A8` is Wingdings `▫`,
`\F0D8` is Wingdings `➔`, `\F0E0` is Symbol `◊`. They are preserved, not
normalised to a standard bullet.

## 3. Numbering

The `@list` level definition is authoritative about **format**; the rendered
marker is authoritative about **appearance**. Where both exist they corroborate,
and any disagreement is a diagnostic rather than a silent choice.

### Formats

`mso-level-number-format` maps to 25+ formats, of which the common ones are
`decimal`, `decimal-leading-zero`, `alpha-lower`/`alpha-upper`,
`roman-lower`/`roman-upper`, `ordinal` (1st, 2nd), `ordinal-text` (first,
second), `cardinal-text` (one, two), `chicago` (*, †, ‡), `hebrew`.

**An absent `mso-level-number-format` means arabic**, not unknown. Word omits it
for the default. A parser that treats absence as unknown loses the format on the
most common list in existence.

### Level text

`mso-level-text` is a pattern, where `%N` is the counter at level N:

| pattern | at level | renders |
|---|---|---|
| `%1.` | 1 | `I.` (with `roman-upper`) |
| `%1.%2` | 2 | `1.1` |
| `(%1)` | 1 | `(c)` (with `alpha-lower`) |
| `%1.%2.%3` | 3 | `2.1.5` |

`expandLevelText` substitutes counter values through their per-level formats.
**It never invents a scheme.** `%1.%2` renders `1.1` — never `Section 1.01`,
never `Article I`. Those are an application numbering scheme overriding Word's,
which is the specific failure this project exists to avoid, and there is an
acceptance test asserting the output contains neither string.

### Counter state

`normalizeLists` walks the blocks maintaining a counter array per list instance:

- entering a level for the first time, or re-entering it after coming back up,
  sets it to the definition's `mso-level-start-at` (default 1) and marks the
  item `restart`
- otherwise the counter increments
- every deeper level is reset, so `1 / 2 / a / b / 3 / a` restarts the alpha
  level at the second `a` — which is what Word does
- a bullet level does not count, but entering one still resets the numbered
  levels below it

Where Word's rendered marker disagrees with the computed one, **Word's text
wins** and the counter is re-seated on it. That happens legitimately when a copy
starts part-way through a list: the counters cannot know the true start value,
but Word drew the right number, so the rest of the list continues correctly from
there. The disagreement is reported as
`WORD_LIST_NUMBER_FORMAT_APPROXIMATED` with both values.

## 4. List identity

Four cases have to be distinguished, and they cannot be told apart from the
visible numbers:

| case | how it is recognised |
|---|---|
| **same list** | consecutive items sharing `(listId, lfo)` |
| **nested** | a deeper `level` within the current run |
| **continuation** | the same `(listId, lfo)` resuming after an interruption — numbering carries forward |
| **restart** | a *different* `(listId, lfo)`, or a level re-entered from above |

`1. A / 2. B / 3. C`, a paragraph, then `1. D / 2. E` — both runs start at 1, so
the numbers say nothing. The `mso-list` declaration says everything: the second
run is list `l1` with override `lfo2`, a different list. Whereas
`1. / 2. / paragraph / 3. / 4.` is one list interrupted, and the model records
both runs under one `listInstanceId`.

Each table cell is its own list scope. A numbered list in the second cell of a
row must not continue the numbering of one in the first.

## 5. Indentation

Word expresses a list item's indentation as a hanging indent:

```css
margin-left:  .5in    /* where the text sits   */
text-indent: -.25in   /* how far left the marker starts */
```

So the marker begins at 0.25in and the text at 0.5in. CSS lists work
differently — the marker sits outside the content box — so the two numbers are
converted rather than copied:

```
textOffsetPx   = marginLeft                     // 48
markerOffsetPx = marginLeft + textIndent        // 24  (textIndent is negative)
hangingPx      = textOffset - markerOffset      // 24
```

A nested list inherits its parent's indentation, so the child emits only the
*difference*. Copying the absolute value doubles the indent at every level —
which is why `bullets/bullet-default` asserts `padding-left: 48px` at all three
levels rather than 48 / 96 / 144.

When nothing is declared anywhere, the fallback is Word's own default ladder of
half an inch per level, and `explicit: false` records that it was a default.

## 6. Rendering the marker

The requirement: **a list marker is never text.** Two modes keep that
separation, and they place the marker differently enough that only one of them
matches Word's actual on-page position.

### `element` (default)

`list-style-type: none` plus an explicit `<span class="wce-marker">`:

```html
<li style="padding-left:48px;text-indent:-48px">
  <span class="wce-marker" style="display:inline-block;min-width:48px;text-indent:0">1.1.1</span>Orbis India
</li>
```

The `<li>` starts its first line at `-hangingPx` (so the marker begins at the
block's own left edge, matching Word's hanging-indent start), and the marker
span *reserves* `min-width: hangingPx` — an `inline-block` narrower than its
reserved width leaves the unused space to its own right, between the marker
and the paragraph text. That is Word's actual layout: the number is followed by
a literal tab character to a fixed column, so the blank space sits **after**
the number and **before** the text.

This is why `element` is the default rather than `native`, and it was not a
guess — it was checked directly in Chromium, rendering the identical CSS both
modes generate at the same gutter width:

```
native:  1. Saji George Yohannan            (marker flush against the text)
element: 1.1.1    Saji George Yohannan      (Word's actual gap)
```

A native `::marker`/`list-style-type` marker is **right-aligned** within its
gutter — any blank space in the gutter collects to the marker's *left*, flush
against the text on the *right*. Widening the gutter does not create a wider
gap between the marker and the text; it only pushes the whole "marker + text"
unit further right, since the marker stays glued to the text either way. That
is a real, verified browser behaviour, not a Word semantic being reinterpreted
— it is simply the wrong CSS mechanism for reproducing a tab stop.

### `native`

Real `<ul>`/`<ol>` with real browser-drawn markers. Word's numbering definition
is compiled into a generated `@counter-style`:

```css
@counter-style wce-1 { system: cyclic; symbols: "•"; suffix: " "; }
.wce-2 { margin: 0; padding-left: 48px; list-style-type: wce-1; }
```

```css
@counter-style wce-3 { system: extends upper-roman; suffix: ". "; }
```

Numbered items also carry `<li value="N">`, taken from the marker Word actually
drew, so a list copied from the middle of a document numbers from where it
really started rather than from 1.

A level text CSS counter styles cannot express — `%1.%2`, or a Word format with
no CSS equivalent such as `ordinal-text` — falls back to `element` rendering
for that level instead of `native`'s usual counter-style:

```html
<li style="padding-left:24px;text-indent:-24px">
  <span class="wce-marker" contenteditable="false" aria-hidden="true">1.1</span>Orbis India
</li>
```

An earlier version of this fallback used a real `::marker { content:
attr(data-marker) " " }`, reading the literal marker text off a `data-marker`
attribute — technically still native. It does not survive a real editor:
verified directly against RoosterJS, the `data-marker` attribute is dropped
when pasted content is converted into the editor's own model, the same way an
external stylesheet's classes are (§ this document, `element` vs `native`
above) — so the marker silently vanished after paste, with nothing left in
its place. `element` rendering needs only the inline `style` this engine
already sets on the span, which is what actually survives that round trip.
There was no real loss in trading this: a composite counter like "1.1" was
never expressible as a single CSS counter-style, so `native` mode was never
going to auto-continue it either way — the fallback gives up nothing `native`
could actually deliver for this case.

Where the glyph could not be mapped to Unicode, the marker keeps its original
font:

```css
.wce-4 > li::marker { font-family: "Courier New"; }
```

Use `native` when the target matters more than the pixel position — a real
`<li value>` lets a paste land inside another editor's own numbered list and
auto-continue it, and an external stylesheet can restyle `::marker` in a way it
cannot restyle a plain `<span>`. Pass it explicitly:

```ts
renderWordDocument(document, { markerMode: 'native' });
```

## 7. Nesting

Word's output is flat: every item is a sibling paragraph carrying its own level
number. `buildListTree` rebuilds the hierarchy with a stack.

Levels can jump by more than one — a document can go from level 1 straight to
level 3, and that is legal. The tree synthesises the intermediate level with a
marker-less spacer item rather than flattening the jump, because the nesting
depth is part of what the document says:

```
ol > li "Top" > ol(level 1) > li(spacer) > ol(level 2) > li "Deep"
```

## 8. When Word's structure is missing entirely

For payloads that have been through another application, `recoverMarkerFromText`
will strip a leading `• ` or `1. ` from paragraph text and reconstruct a marker
from it. Unlike everything else in this module, that **is** a guess: it is
diagnosed as `WORD_LIST_MARKER_HEURISTIC` every time, and can be turned off:

```ts
parseWordClipboard(payload, { recoverMarkersFromText: false });
```

Similarly, when the `@list` rule is missing, `inferNumberFormat` reads the format
back out of the rendered marker (`IV.` → upper-roman, `c)` → lower-alpha,
`1.1` → a two-level decimal pattern). Also diagnosed, also a guess, and the
level text it produces keeps the separators Word actually drew rather than
inventing a new scheme.
