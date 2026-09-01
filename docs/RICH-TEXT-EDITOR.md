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

Both are demonstrated, fixed, and verified below using
[RoosterJS](https://microsoft.github.io/roosterjs) — Microsoft's own
content-model-based editor, and a deliberately unforgiving choice for this,
since it does the most re-deriving of any mainstream option.

A third requirement followed once someone actually used the editor: real
`<ol>`/`<ul>` markers, not this engine's usual `element` spans, so that
pressing Enter at the end of a list item continues the numbering the way
Word — and every other editor — does natively. That is a different rendering
mode entirely (`markerMode: 'native'`, see
[LIST-PARSING.md § 6](LIST-PARSING.md#6-rendering-the-marker)), and switching
to it surfaced its own RoosterJS-specific bug, covered in § 3 below.

## Try it

```bash
npm run dev
```

Open `/rooster-editor/`. It's a real `createEditor()` from the `roosterjs`
package, completely unmodified, with one additional plugin:
`WordClipboardEnginePlugin`. Paste real Word content, or use **Paste a
fixture** to simulate one. Try both of the things this page exists to
demonstrate:

- Click at the end of a numbered or bulleted item and press Enter — the next
  item gets the next number or bullet automatically, drawn by the browser,
  the same as typing a new item in Word itself.
- Click right before a marker that *isn't* auto-numbering (a composite level
  like "1.1", see § 3) and press Backspace — the whole marker still comes out
  as one unit, never one corrupted character at a time.

## 1. Protecting the marker: `contenteditable="false"`

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

## 2. Making the geometry survive RoosterJS's content model: inline everything

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

## 3. Real `<ol>`/`<ul>` markers, so the numbering auto-continues

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
model. Nothing was really lost: a composite counter like "1.1" was never
going to auto-continue from a single CSS counter-style anyway, so the
fallback gives up an auto-continuation `native` mode could never have
delivered for that case regardless of RoosterJS.

Verified end to end: pasting a multilevel Word list where the top level is
simple roman numbering (I., II., …) and the second level is composite ("1.1",
"1.2", …) now renders both correctly — the top level as real, auto-continuing
`::marker`s, the second level as durable `element` spans that render every
time and never vanish, exactly matching what the engine's own renderer
produced before RoosterJS ever saw it.

## 4. A limitation this integration has, honestly

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

## Building your own adapter

The whole integration is one file,
[`src/demo/rooster-editor/WordClipboardEnginePlugin.ts`](../src/demo/rooster-editor/WordClipboardEnginePlugin.ts),
and none of it is RoosterJS-specific in spirit — only in which hook it uses.
The shape any editor's adapter needs is the same:

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

None of this requires `word-clipboard-engine` to know anything about
RoosterJS, or about whatever editor you're integrating with — the model and
renderer stay exactly as editor-agnostic as
[ARCHITECTURE.md](ARCHITECTURE.md) describes. The adapter is where the
editor-specific knowledge belongs, and it can live entirely outside this
package.
