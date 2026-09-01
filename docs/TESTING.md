# Testing

```bash
npm test                      # everything — 217 tests
npm run test:word             # parsing + regression only
npm run test:golden           # the fixture suite
npm run typecheck
UPDATE_FIXTURES=1 npm test    # re-bless the golden files
```

## The suites

| directory | what it holds |
|---|---|
| `src/tests/detection/` | signal weighting, thresholds, source identification |
| `src/tests/parsing/` | the CSS tokenizer, stylesheet, lengths, colours, symbol fonts, number formats |
| `src/tests/normalization/` | list identity and counters, indentation, table grids, style folding |
| `src/tests/rendering/` | markers, runs, tables, security, the standalone document |
| `src/tests/regression/` | the fixture golden tests, and the acceptance criteria |

`acceptance.test.ts` is the one to read first. It encodes the specification's
non-negotiable criteria as assertions that are independent of any fixture's
blessed output — because a golden file records whatever the engine *currently*
does, which is not the same as what it is *required* to do.

## Fixtures

A fixture is a directory under `src/fixtures/word/`:

```
src/fixtures/word/bullets/bullet-default/
  input.html            the Word clipboard payload, byte for byte
  expected-model.json   the projected canonical model
  expected.html         the rendered output plus the generated stylesheet
  notes.md              what this fixture is testing and why
```

There are 21, covering: plain paragraphs and headings; character and paragraph
formatting; default and custom bullets; simple, roman, multilevel, restarted and
start-at numbering; mixed number/bullet hierarchies; list continuation; simple,
merged, nested and list-containing tables; resolved and unresolved images;
hyperlinks and bookmarks; and one complex combined report.

Each one runs five checks:

1. the projected model matches `expected-model.json`
2. the rendered HTML matches `expected.html`
3. **no list marker appears inside a text node** — the engine's central
   invariant, checked on every fixture rather than only the list ones
4. no Word-only markup survives (`mso-*`, `<o:p>`, conditional comments, `Mso`
   classes)
5. no script, event handler or `javascript:` URL is emitted

### Why `expected-model.json` is a projection

The full `WordDocument` embeds the raw payload and the entire parsed stylesheet.
A golden file would be mostly Word's CSS, every irrelevant change to it would
break every test, and no human could read the diff.

`src/tests/support/model-projection.ts` projects the facts the engine actually
promises: block structure, list identity and markers, run boundaries and their
formatting, table geometry, image resolution, links, and a sorted diagnostic
summary. Anything not projected is still tested — by unit tests against the
module that produces it — but is not part of the golden contract.

### Blessing

```bash
UPDATE_FIXTURES=1 npm test
git diff src/fixtures/
```

Read the diff before committing it. For a new fixture the expected files are
written automatically on first run. For an existing one, a change is either a
regression or an intentional behaviour change, and there is no third
possibility — that is the entire value of the mechanism.

Three of the engine's real bugs were found exactly this way. The
`mso-bidi-font-weight` bug showed up as `"Bold" | [null]` in a re-blessed model
diff: Word writes every bold run as `<b style='mso-bidi-font-weight:normal'>`,
and reading that `normal` as the font weight had been silently un-bolding
everything.

The third was found by a user pasting a real document rather than by the suite,
which is the honest version of the story: the fixtures were all *tidier* than a
genuine Word payload, and the bug lived in exactly the gap between them. See
**Head hoisting** below.

## Capturing a real Word payload

The fixtures in this repository are hand-authored to reproduce the markup shapes
Word documented and emits — the `ProgId`/`Generator` meta pair, the Office
namespaces, the comment-wrapped `<style>` with `@font-face`/`@list`/`Mso*`
rules, the CF_HTML markers, the `WordSection1` wrapper, the
`<![if !supportLists]>` marker blocks, the `<o:p>` marks, the `mso-*`
declarations. They are not captures from a running copy of Word.

That is a real limitation, and the fix is to add captures. The rule from the
specification is that **every bug found in real Word content becomes a
fixture**, and the lab makes that a two-step job:

1. `npm run dev`, paste the content that misbehaves.
2. **Export fixture** in the first pane downloads the exact `input.html`.
3. Put it at `src/fixtures/word/<category>/<name>/input.html`, write a
   `notes.md` saying what it is testing, and run `UPDATE_FIXTURES=1 npm test`.
4. **Read the blessed output before committing it.** If it is wrong, that is the
   bug — fix the engine, not the fixture.

Fixtures containing anything confidential should not be committed. Replace the
text with something neutral first; the markup structure is what matters, and it
survives a search-and-replace of the content.

## Test environment caveats

The suite runs under happy-dom, which diverges from browsers in two ways worth
knowing:

**Conditional comments.** A browser keeps `<!--[if gte vml 1]> … <![endif]-->`
as a single comment node — that is the entire point of a downlevel-hidden
comment. happy-dom splits it into a comment plus live elements. The engine
handles both, because the VML is mined from the raw text and the live elements
defer to the `<img>` twin, and the test suite therefore exercises the harder of
the two paths.

**Inert documents.** `DOMParser` documents do not load subresources in a
browser. happy-dom's do. This is why executing elements are removed from the
source text before parsing rather than after — see
[ARCHITECTURE.md](ARCHITECTURE.md#why-executing-elements-are-removed-from-the-text).

**Head hoisting.** happy-dom and Chromium disagree about where an unknown
element in `<head>` ends up, and that disagreement hid a bug that made real
pastes come out empty. The fixtures now all carry the `<link rel=File-List>`
and `<link rel=Edit-Time-Data>` elements Word puts in every payload, precisely
so the suite exercises the case that broke. It is the clearest evidence there
is that a fixture tidier than the real thing is not really a fixture.

Because of the first point, native marker rendering cannot be verified under
happy-dom — `::marker` is not implemented there. It was verified manually in
Chromium against the built demo: the generated `@counter-style` is accepted
(`getComputedStyle(li).listStyleType === 'wce-1'`), and
`getComputedStyle(li, '::marker').content` resolves to `"1.1 "` for the
multilevel case. A browser-based test would be a worthwhile addition.

## Adding a test

Unit-test the module, and add a fixture if the behaviour is visible end to end.
The fixture proves the pipeline holds together; the unit test says which part
broke when it does not.
