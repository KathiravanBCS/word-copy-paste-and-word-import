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

## Try it

```bash
npm run dev
```

Open `/rooster-editor/`. It's a real `createEditor()` from the `roosterjs`
package, completely unmodified, with one additional plugin:
`WordClipboardEnginePlugin`. Paste real Word content, or use **Paste a
fixture** to simulate one. Then click right before a marker and press
Backspace — the whole marker comes out as one unit, never one corrupted
character at a time.

## 1. Protecting the marker: `contenteditable="false"`

Every marker `element` mode renders now carries it:

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
`contenteditable="false"` is inert on a page that isn't itself editable.

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

## 3. A limitation this integration has, honestly

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
   `parseWordClipboard` + `renderWordDocument({ markerMode: 'element' })`.
   `element` mode specifically, not `native` — see LIST-PARSING.md for why,
   and note that a native `::marker` has no `contenteditable` to protect it
   with in the first place.
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
