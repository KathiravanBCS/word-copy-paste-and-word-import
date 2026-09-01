# Rich text editor integration

`word-clipboard-engine` stays a plain library — zero runtime dependencies,
no editor imported anywhere in `src/index.ts` or anything it pulls in. This
page is about the other half of the picture: using it as the paste handler
*inside* a live, editable rich-text editor, where the output isn't just
displayed once but has to keep working while someone edits it.

That is a materially different job from the static HTML this engine produces
for display or download, and it surfaces two problems neither of those cases
has to deal with:

1. **A list marker that used to be a read-only `<span>` is now sitting inside
   editable content.** Nothing stops a cursor from landing inside it, or a
   Backspace from eating it one character at a time — silently reintroducing
   exactly the "bullet became text" failure this whole engine exists to
   prevent, just one editing session after the paste instead of at parse
   time.
2. **A rich-text editor doesn't render your HTML — it re-derives its own.**
   Most non-trivial editors (RoosterJS included) convert whatever you paste
   into their own internal document model, then render *that* back to DOM.
   Anything your HTML implied but their model doesn't represent is gone the
   moment that conversion happens, whether or not it "looked fine" the
   instant it landed.

A third requirement followed once someone actually used an editor built this
way: real, auto-continuing markers — pressing Enter at the end of a list item
should continue the numbering the way Word, and every other editor, does
natively, not leave a blank unmarked line.

Two integrations demonstrate, fix, and verify all three, against two
editors chosen for how differently they're built:

- **[RoosterJS](#roosterjs-integration)** — Microsoft's own
  content-model-based editor. It converts whatever you give it into a fixed
  internal model with no field for Word-specific numbering data, so getting
  it right takes real work: a rendering-mode choice, one library-level fix,
  and a JS layer this document explains in full — because that is what
  integrating with a *closed* editor model actually costs, honestly shown
  rather than glossed over.
- **[TipTap](#tiptap-integration)** (ProseMirror) — an editor where *you*
  define the schema. Given a real node type for a Word list item, none of
  RoosterJS's workarounds are needed at all: the schema itself has the field
  RoosterJS's didn't, and every Word number format auto-continues correctly,
  not just decimal chains.

Read the RoosterJS section for what a closed content model costs and how far
workarounds can close the gap; read the TipTap section for what removes the
gap entirely. If you're integrating with your own editor, jump to
[Building your own adapter](#building-your-own-adapter) — it names which
path fits which kind of editor.

## RoosterJS integration

That closed content model is also what makes RoosterJS a deliberately
unforgiving choice to verify against — it does the most re-deriving of any
mainstream editor, so whatever survives it is a real result, not a lucky
default.

### Try it

```bash
npm run dev
```

Open `/rooster-editor/`. It's a real `createEditor()` from the `roosterjs`
package, completely unmodified, with one additional plugin:
`WordClipboardEnginePlugin`. Paste real Word content, or use **Paste a
fixture** to simulate one. Try all three of the things this page exists to
demonstrate:

- Click at the end of a numbered or bulleted item and press Enter — the next
  item gets the next number or bullet automatically, drawn by the browser,
  the same as typing a new item in Word itself.
- Click at the end of a **composite**-numbered item (a "1.1"/"1.1.1" style
  level, see § 3) and press Enter — the new item gets numbered too, just not
  instantly: it appears once you move on to something else (typically the
  next Enter, or clicking away), because that number comes from this plugin's
  own JS, not the browser — § 4 explains why, and why the short delay is
  deliberate rather than a bug.
- Click right before any marker and press Backspace — the whole marker still
  comes out as one unit, never one corrupted character at a time.

### 1. Protecting the marker: `contenteditable="false"`

`element`-mode markers carry it:

```html
<span class="wce-marker" contenteditable="false" aria-hidden="true">1.1.1</span>
```

This is a core-library change (`HtmlListRenderer.ts`), not something specific
to RoosterJS — it benefits any editable surface the output lands in. It marks
the span as an atomic, non-editable island inside editable content, the same
technique every mainstream rich-text editor uses for generated "chip" content
(mentions, emoji) embedded in editable text.

Verified directly in Chromium, not assumed — a synthetic paragraph, one
Backspace at the start of the text, with and without the attribute:

```
without contenteditable="false":  "1.1.1" -> "1.1."   (corrupted — last
                                                          character silently
                                                          deleted)
with contenteditable="false":     "1.1.1" -> gone       (removed as a whole
                                                          unit, cleanly)
```

The second is correct: it matches Word's own UX for backspacing at the very
start of a list item (the marker/list formatting is removed in one keystroke,
not eaten digit by digit), and it holds up **inside the live RoosterJS
editor** after a real paste, not just in isolation — same test, run against
actual pasted content: one Backspace removes the marker as a whole, typing
immediately afterward continues to work normally, nothing is left corrupted.

It costs nothing when the output isn't inside an editable ancestor —
`contenteditable="false"` is inert on a page that isn't itself editable. And
since § 3 below moved this editor to `native` markers, most items in a real
paste don't have a marker span at all any more — a genuine `::marker` is
drawn by the browser outside the editable DOM entirely, so there is nothing
for a stray keystroke to land in. This attribute still matters for the levels
`native` mode itself falls back to `element` for (composite formats like
"1.1" — § 3), which is the case actually pictured above.

### 2. Making the geometry survive RoosterJS's content model: inline everything

This one took an actual empirical isolation to find, because the first
attempt looked like it worked and didn't.

The engine's generated CSS (native counter styles, or — in `element` mode,
the default — the marker's `min-width`/`padding-left`/`text-indent` geometry
from [LIST-PARSING.md § 6](LIST-PARSING.md#6-rendering-the-marker)) is
class-based, meant to be installed once per document rather than repeated on
every element. The first version of `WordClipboardEnginePlugin` did exactly
that — installed the CSS in `<head>`, handed RoosterJS a fragment whose
`<ol>`/`<li>`/marker carried only the class names.

Pasting a multilevel numbered list came out **double-numbered**: RoosterJS's
own default `<ol>` numbering (`1. 2. 3.`) rendered right alongside this
engine's own marker text (`I. / 1.1 / 1.2`), because RoosterJS's DOM→Content
Model conversion — which runs on the detached fragment, before it's ever part
of the live document — reads `element.style`, not CSS classes resolved
against a stylesheet the fragment isn't yet attached to. `list-style-type:
none` from the class was never in scope, so `<ol>` fell back to the browser's
own default numbering.

Confirmed directly, both directions, in Chromium:

```
padding-left / list-style:none set via a CSS class:    dropped after paste
padding-left / list-style:none set as inline style:     preserved after paste
```

The fix (`WordClipboardEnginePlugin.inlineGeneratedStyles`) reads the already-
parsed `CSSStyleSheet` this same plugin installs, walks its rules, and
applies each one's declarations as inline `style` on every matching element in
the fragment before handing it to RoosterJS — reusing the browser's own
selector engine (`querySelectorAll` on the detached fragment) rather than a
hand-rolled CSS matcher. Existing inline declarations always win, so nothing
this engine already set explicitly can be overwritten.

With that in place, the same multilevel list pastes with the correct,
Word-matching structure and no duplicate numbering — verified end to end
against the real integration, not a synthetic case.

### 3. Real `<ol>`/`<ul>` markers, so the numbering auto-continues

This engine's own default output uses `element` markers (§ 1) — a span with
the literal marker text, positioned to match Word's exact gutter. That is the
right choice for a static render, but it is not a real list as far as any
editor's own editing logic is concerned: nothing about a `<span>1.</span>`
tells RoosterJS "this is item 1 of an ordered list, and the next one should
say 2." Press Enter at the end of it and RoosterJS just starts a new
paragraph with no marker at all.

`WordClipboardEnginePlugin` renders with `markerMode: 'native'` instead —
real `<ol>`/`<ul>` with real browser `::marker`s, driven by a generated
`@counter-style` (see
[LIST-PARSING.md § 6](LIST-PARSING.md#6-rendering-the-marker)). That makes it
an actual list as far as RoosterJS's content model is concerned, so its own
list-continuation logic applies: pressing Enter inside an item inserts the
next `<li>` in the same `<ol>`, and the *browser* draws the next number —
verified directly, pasting a Word list with roman-numeral top-level items
(I., II., …) and pressing Enter at the last one produces III., then IV.,
purely from the browser's own counter, nothing this engine or the plugin
computed.

This traded away Word's exact gutter spacing for that (§ "native" in
LIST-PARSING.md: a native `::marker` sits flush against the text, not at
Word's tab-stop distance) — a deliberate choice for a live editor, where
continuing to type correctly matters more than a few pixels of gutter.

**It also surfaced a real bug**, not merely a tradeoff. `native` mode's
handling of a level text CSS counter styles cannot express — a composite like
Word's `%1.%2` ("1.1", "1.2", …) — used to fall back to a real `::marker {
content: attr(data-marker) }`, reading the marker text off a `data-marker`
attribute on the `<li>`. Pasting a list with that kind of level (common for
legal/technical section numbering) made those markers **vanish entirely**
after going through RoosterJS — not corrupted, not misplaced, just gone, with
the item's text left unmarked.

Isolated directly: RoosterJS's DOM→Content Model conversion keeps inline
`style` and a curated set of attributes it recognises (like `<li value>`), but
drops a `data-marker` attribute it has no reason to know about — the same
mechanism § 2 already found for CSS classes, just hitting an attribute this
time instead of a stylesheet rule. Confirmed both directions in Chromium: the
engine's own rendered HTML has `data-marker="1.1"` on the `<li>` right up
until it's handed to RoosterJS; the DOM RoosterJS renders back out has no
`data-marker` anywhere, and no visible marker where "1.1" should be.

The fix is at the library level, not the plugin: `HtmlListRenderer.ts` no
longer generates the `data-marker`/`::marker` pairing at all. A level whose
text can't become a single CSS counter-style now falls back to the same
`element` rendering non-native mode uses — a real span, protected by
`contenteditable="false"` (§ 1), needing nothing but the inline `style` this
engine already sets, which is what actually survives an editor's content
model. A composite counter like "1.1" was never going to auto-continue from a
single CSS counter-style, so this fallback gives up nothing the browser
itself could have delivered — what it does give up, the plugin gets back a
different way, in § 4.

Verified end to end: pasting a multilevel Word list where the top level is
simple roman numbering (I., II., …) and the second level is composite ("1.1",
"1.2", …) now renders both correctly — the top level as real, auto-continuing
`::marker`s, the second level as durable `element` spans that render every
time and never vanish, exactly matching what the engine's own renderer
produced before RoosterJS ever saw it.

### 4. Auto-continuing composite markers too, in JS

§ 3 leaves one thing genuinely undone: a composite marker's *text* survives
now, but it does not renumber itself the way a real `::marker` does. Press
Enter at the end of "1.4 Given" and the new item still reads nothing — a
plain `contenteditable="false"` span has no counter to advance, because
nothing (there is no such CSS mechanism) is watching it.

`WordClipboardEnginePlugin` closes that gap itself, in
`renumberCompositeMarkers`: it listens for RoosterJS's `contentChanged` event
(skipping the one that *is* the paste itself — those numbers already came
from Word) and, on every other edit, recomputes every composite marker in the
editor from its current position in the list. Concretely: split a marker's
own text on its digit runs to recover its shape ("1.1." → literal parts
`["", ".", "."]` around two placeholders), then refill those placeholders
from the item's real position at each nesting level — the 2nd sub-item of the
1st item is `[1, 2]`, giving "1.2". An item a fresh Enter created with no
marker of its own gets one synthesised from a sibling's, not left blank.

That position lookup needed its own fix along the way: `HtmlRenderer.ts`
nests a sub-list *inside* its owning `<li>` (`<li>Background<ol>…</ol></li>`),
but verified directly, RoosterJS does not keep that shape — a real paste
comes back with the sub-list as Background's *sibling* instead
(`<li>Background</li><ol>…</ol>`), because RoosterJS's list representation
tracks nesting as a per-item level stack, not DOM containment. (It is not
even consistent about it — a single-item sub-list came back properly nested
in the same document. `listContainer()` in the plugin handles both shapes,
rather than assuming either.)

Two things stop this from racing the person actually typing:

- **A brand-new item never gets a marker while the caret is still in it.**
  Verified directly: synthesising one immediately (the moment Enter creates
  the empty item) works until the very next keystroke, at which point
  RoosterJS's own DOM regeneration — which has no idea our marker is
  supposed to stay glued to the front of the line, since we put it there
  outside its content model entirely — was seen to relocate it mid-word.
  Skipping the item the `Selection` is currently anchored in avoids the race
  by never touching a line that is still being typed into; it gets its
  marker as soon as *something else* changes (typically the next Enter).
- **The very last item still gets numbered eventually.** Skipping the active
  item forever would leave whatever was typed last permanently blank once
  nothing else changes it. A `focusout` listener on the document catches
  that: once the editor has genuinely lost focus (checked a tick later, since
  `hasFocus()` still reports the outgoing state on the event itself), nothing
  is mid-keystroke anywhere, so the same renumbering pass runs once more with
  no item excluded.

Scoped honestly, not silently: this recovers **decimal** composite numbering
only — "1.1", "1.1.1", the overwhelmingly common case (business, legal,
technical section numbering, this file's own worked examples). A composite
that mixes formats (Word's rarer "I.1.a", roman then decimal then alpha) has
no digit run for its roman or alpha segments and is left exactly as § 3 leaves
it: a real, protected, static marker, just not one this function knows how to
advance. And because recomputation always starts from the document's current
structural position, an edit *anywhere* resequences every composite marker
in the document from 1 at each level — a paste that started mid-document at
"3.4" keeps reading "3.4" until the next edit, then renumbers like everything
else. Preserving that offset indefinitely would need the same Word-level
metadata § 3 already established does not survive RoosterJS's round trip.

### 5. A limitation this integration has, honestly

Word content the engine cannot resolve to a real image (a `file:///` path
with no clipboard bytes to recover) renders as a labelled placeholder — a
dashed-border box, per [FIDELITY.md](FIDELITY.md). Inside RoosterJS, the box
styling does not survive: the placeholder's *text* label comes through fine,
and — this is the part that actually matters — **no broken `file:///` image
reference ever reaches the editor either way**, but the dashed border,
background and padding are gone.

Isolated directly: even set as inline `style` from the very first frame (no
class involved at all), RoosterJS's Content Model preserves only a small,
curated set of inline-segment properties (`background-color` among them) and
drops `border`, `padding` and `display` outright. This is not a bug in the
inlining pass above — border/padding/display on an inline text segment simply
has no field in RoosterJS's model to be preserved into. Where the previous
problem was "the right CSS, delivered the wrong way," this one is "RoosterJS's
document model has no concept for this kind of content at all."

The safety property that actually matters is intact regardless: nothing in
this engine's output ever emits a live `src="file:///…"`, on any surface, so
there is nothing here for a broken image reference to leak through. The open
item is cosmetic — restoring the placeholder's visible box — and the
principled fix is RoosterJS's **Entity** API (a way to mark a subtree as
opaque, editor-managed content RoosterJS won't try to re-derive formatting
for), which is a real, separate integration this file does not yet cover.

## TipTap integration

RoosterJS's every workaround above exists because its Content Model is
*fixed*: a curated set of properties the editor already knows about, with no
field for "this marker is `%1.%2`, decimal, decimal" — nowhere to put it, no
matter how the plugin tries to hand it over. TipTap ([tiptap.dev][tiptap]) is
built on ProseMirror, where the schema is *yours*. Give it a real node type
for a Word list item, and the field simply exists — nothing needs recovering
after the fact, because nothing was ever going to be thrown away.

[tiptap]: https://tiptap.dev

### Try it

```bash
npm run dev
```

Open `/tiptap-editor/`. Paste real Word content, or use **Paste a fixture**.
Press Enter at the end of *any* numbered or bulleted item — including a
composite level like "1.1" or "1.1.1" — and the next one is numbered
correctly **immediately**. No delay, no caret-tracking workaround: unlike the
RoosterJS integration's § 4, there is nothing here that needs one.

### The schema: `wordList` and `wordListItem`

Two node types, defined in
[`src/demo/tiptap-editor/WordListNodes.ts`](../src/demo/tiptap-editor/WordListNodes.ts),
mirror `HtmlRenderer.ts`'s own `<ol>`/`<li>` nesting shape. `wordListItem`
carries Word's numbering *declaration* as node attributes — `levelText`
(`"%1.%2"`), `numberFormat` (`decimal`, `upper-roman`, …), the list's
`startAt` — captured straight off the `data-word-*` attributes this engine's
renderer already emits with `includeWordMetadata: true`. No new rendering
mode was needed in the core library for this: `element` marker mode already
carries everything, because a plain HTML attribute is something ProseMirror's
own `DOMParser` reads directly into node attrs, by a rule this file writes
itself — nothing is filtered through an editor's own idea of which properties
matter, the way RoosterJS's Content Model conversion filters everything.

A marker is never stored as text. A ProseMirror plugin
(`wordListMarkerPlugin`) draws it as a `Decoration.widget`, recomputed fresh
from live document position on every read, using this engine's own
`expandLevelText`/`formatNumber` (`WordListStyleParser.ts`, already public
API) — the exact functions the static renderer uses to cross-check Word's own
marker text. The 2nd sub-item of the 1st item is structurally `[1, 2]`; feed
that through the item's declared format and level text, and the marker is
correct for *any* Word number format — roman, alpha, ordinal, any mix at any
depth — not just decimal composites, because nothing here treats decimal as
special the way the RoosterJS integration's JS-side regex had to.

A widget decoration also means the marker was never part of the document's
editable text in the first place — not protected by an attribute the way
RoosterJS's markers need `contenteditable="false"` (§ 1 above), just
structurally impossible to land a cursor inside or eat with Backspace one
character at a time. Backspace at the very start of an item's text merges it
into the previous item — ProseMirror's ordinary block-join behaviour,
requiring no special handling — which is closer to Word's own UX than either
of RoosterJS's marker-removal behaviours turned out to need to be.

### Three bugs this surfaced, one of them in the core library

**Pasting dropped the outermost list.** The first working version pasted a
multilevel list with every item at the top level rendered as a bare
paragraph — no marker, no `<ol>` — while a *nested* list further inside the
same paste came through fine. Isolated directly by dumping
`editor.getJSON()`: `DOMParser.parseSlice()`, the usual choice for pasted
content, computes `openStart`/`openEnd` from how deep the first and last leaf
sit and trims wrapping nodes at those edges — correct, wanted behaviour for
merging a paste into existing content, wrong for this handler, which wants
the structure inserted exactly as rendered. Fixed by parsing as a complete,
self-contained document (`parser.parse(root)`) and wrapping its content in a
`Slice` built with `openStart`/`openEnd` both `0`, in
[`WordClipboardExtension.ts`](../src/demo/tiptap-editor/WordClipboardExtension.ts) —
no trimming assumed.

**Pressing Enter added a paragraph, not a new item.** `wordListItem.content`
allows more than one paragraph (`paragraph+`, for the rare genuinely
multi-paragraph list item), so ProseMirror's default Enter handling
(`splitBlock`) just split the *paragraph*, leaving the typed text inside the
same item — no new marker anywhere, because no new item was ever created.
Fixed with `prosemirror-schema-list`'s own `splitListItem` command, bound to
Enter in `WordListNodes.ts` — the standard mechanism every list
implementation on ProseMirror uses for this, not something specific to this
schema.

**A composite marker's ancestor digit used the wrong format — a real,
previously-hidden bug, not a TipTap-specific one.** Given a roman level 0
("I.") and a decimal-declared level 1 whose level text is `"%1.%2"`, the
first version computed `"I.1"`. The correct, Word-verified answer is
`"1.1"`: a composite level text's *own* format governs **every** placeholder
it contains, including ones naming an ancestor's counter — `%1` here does not
mean "show the ancestor the way it shows itself," it means "show the
ancestor's current count, in the format *this* level declared." This
integration's live marker computation has no Word-rendered text to fall back
on for a freshly typed item, so getting the format actually right (not just
plausible) mattered immediately — and tracing it back found the exact same
bug already latent in `NormalizeLists.ts`'s own fallback computation (used
when Word omits the marker entirely, e.g. content pasted from Word Online).
It was invisible until now only because that fallback is a diagnostic
cross-check, not the displayed text, whenever Word's own literal marker is
present — which every existing fixture's payload does provide. Fixed in both
places; see the `NormalizeLists.ts` comment at the same spot for the core-library
side, and [LIST-PARSING.md](LIST-PARSING.md) for how a composite level text is
otherwise handled.

### What's not (yet) ported

Lists get this integration's full, custom treatment; paragraphs, tables,
images, and character formatting (bold/italic/underline/strikethrough) go
through TipTap's own default schema and `DOMParser`, unmodified — verified
directly, a pasted table and character formatting both survive intact. Two
things a from-scratch TipTap setup would also need, not wired into this demo
specifically: colour/highlight/font-family marks (`TextStyle`/`Color`/
`Highlight`/`FontFamily` extensions this demo doesn't install), and the same
placeholder-styling gap § 5 above describes for RoosterJS — an unresolved
image still never becomes a broken `file:///` reference, it just isn't
wrapped as richly as the static renderer's dashed-border box. Neither is a
limitation of the *approach* — both are additional TipTap extensions or
schema work this demo scoped out, not a wall the way RoosterJS's Content
Model was.

## Building your own adapter

Which pattern to follow depends on what your editor lets you define:

- **Your editor lets you register a custom node/schema type** (ProseMirror —
  TipTap, Lexical's node system, or similar) — follow the **TipTap**
  integration's shape: a real node for a Word list item, attributes for its
  numbering declaration, and a decoration/plugin that draws the marker fresh
  from live position. This is the one that actually removes the problem,
  not just works around it, and it isn't ProseMirror-specific in spirit —
  only in which APIs (`Node.create`, `Decoration.widget`) it uses.
- **Your editor has a fixed content model you cannot extend** (RoosterJS, or
  anything similarly closed) — follow the **RoosterJS** integration's shape
  below: render with this engine's own marker geometry, inline the generated
  CSS, and accept that composite (non-decimal-counter) markers need a JS
  layer to auto-continue, because there is nowhere in the editor's model to
  put a declaration it could read back later.

The RoosterJS integration is one file,
[`src/demo/rooster-editor/WordClipboardEnginePlugin.ts`](../src/demo/rooster-editor/WordClipboardEnginePlugin.ts),
and none of it is RoosterJS-specific in spirit — only in which hook it uses.
The shape any *closed-model* editor's adapter needs is the same:

1. **Find the editor's own pre-insertion hook.** RoosterJS calls it
   `BeforePasteEvent`; look for the equivalent — a point where the editor
   hands a plugin the about-to-be-inserted content before its own paste
   normalisation runs.
2. **Read the raw clipboard HTML from that hook**, not from the editor's own
   (already Word-mangled) interpretation of it. RoosterJS exposes it as
   `event.clipboardData.rawHtml`.
3. **Detect, parse and render:** `detectWordHtml` to skip non-Word paste
   (let the editor's own general-purpose handling run for that — this engine
   has no reason to intercept a paste from a webpage), then
   `parseWordClipboard` + `renderWordDocument(...)`. Pick the marker mode
   deliberately, not by default: `native` (§ 3) if the target editor's own
   list-continuation should take over after the paste — the common case for a
   live editor, and what this integration uses; `element` if it shouldn't, or
   if exact Word-matching gutter spacing matters more than auto-continuation
   (a native `::marker` also has no `contenteditable` to protect it with, but
   in `native` mode that is a non-issue rather than a gap — see § 1).
4. **Check whether the target keeps CSS classes across its own content
   model, or only inline style**, before assuming either survives — this
   is the one step worth actually testing empirically per editor rather than
   assuming, given what section 2 above found.
5. **Replace the target's insertion point with the rendered content.**

For an extensible-schema editor, the equivalent shape is
[`src/demo/tiptap-editor/WordClipboardExtension.ts`](../src/demo/tiptap-editor/WordClipboardExtension.ts)
(the paste handler) and
[`WordListNodes.ts`](../src/demo/tiptap-editor/WordListNodes.ts) (the schema
and live marker computation) — read the raw clipboard HTML from the editor's
own pre-insertion hook exactly as above, render with `element` marker mode
and `includeWordMetadata: true` so the HTML carries a numbering declaration
your schema's `parseHTML` can read, then parse straight into your schema
instead of handing the editor a DOM fragment to interpret on its own.

None of this requires `word-clipboard-engine` to know anything about
RoosterJS, TipTap, or whatever editor you're integrating with — the model and
renderer stay exactly as editor-agnostic as
[ARCHITECTURE.md](ARCHITECTURE.md) describes. The adapter is where the
editor-specific knowledge belongs, and it can live entirely outside this
package.
