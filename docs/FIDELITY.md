# Fidelity and diagnostics

The engine never silently discards a Word construct. Anything it cannot
represent exactly produces a diagnostic on `document.diagnostics`, and the
`FidelityReport` aggregates them.

This document is the contract for what "preserved" means, construct by
construct.

## Fidelity classes

| class | meaning |
|---|---|
| `EXACT` | represented without loss; the output means what the source meant |
| `EQUIVALENT` | a different mechanism, the same result — a Symbol byte rendered as its Unicode glyph, a VML picture rendered as an `<img>` |
| `APPROXIMATED` | visibly close, not identical — a floating image become a CSS float, a cropped image become a CSS clip |
| `UNSUPPORTED` | no HTML equivalent; preserved on the model, not rendered |

The class describes the *outcome*, not the severity. An `EQUIVALENT` mapping is
not a problem; it is a note that the mechanism changed. A pasted document with
50 `EQUIVALENT` diagnostics and no `APPROXIMATED` ones came through perfectly.

## Reading them

```ts
const document = parseWordClipboard(payload);

for (const d of document.diagnostics) {
  console.log(d.severity, d.fidelity, d.code, d.message, d.count ?? 1);
}
```

Identical `code` + `message` pairs are folded into one entry with a `count`, so
a 400-page paste with 900 unresolvable images reports one line, not 900.

```ts
import { getFidelityReport, formatFidelityReport } from 'word-clipboard-engine';

const report = getFidelityReport(document);
report.unresolvedImages;        // 3
report.fidelityBreakdown;       // { EXACT: 2, EQUIVALENT: 4, APPROXIMATED: 1, UNSUPPORTED: 0 }
report.unsupportedFeatures;     // ['WORD_OLE_OBJECT', …]
console.log(formatFidelityReport(report));
```

`renderStandaloneHtml` appends the same summary to the downloaded file as a
trailing HTML comment, so an exported document is self-documenting about what
survived.

## Codes

### Lists

| code | fidelity | when |
|---|---|---|
| `WORD_LIST_DEFINITION_MISSING` | APPROXIMATED | a paragraph references a list the stylesheet does not define — the rendered marker is used as the authority instead |
| `WORD_LIST_LEVEL_MISSING` | APPROXIMATED | the `@list` rule exists but not for that level |
| `WORD_LIST_MARKER_HEURISTIC` | APPROXIMATED | the marker was guessed — from leading text, or by inferring the format from what Word drew. The only place the engine guesses, and it always says so |
| `WORD_LIST_NUMBER_FORMAT_APPROXIMATED` | EXACT | Word's rendered marker disagrees with the computed one; Word's is kept and the counter re-seated. Expected when a copy starts part-way through a list |

### Symbols

| code | fidelity | when |
|---|---|---|
| `WORD_SYMBOL_FONT_MAPPED` | EQUIVALENT | a symbol-font byte was mapped to its Unicode equivalent (`\F0B7` in Symbol → `•`). The raw byte and font are kept on the model |
| `WORD_SYMBOL_FONT_UNMAPPED` | APPROXIMATED | a symbol-font byte with no known mapping. Rendered as-is **in its original font** rather than replaced by a guess |

### Images

| code | fidelity | when |
|---|---|---|
| `WORD_LOCAL_FILE_IMAGE` | APPROXIMATED | `file:///` reference — exists only on the machine the content was copied from |
| `WORD_CID_IMAGE` | APPROXIMATED | `cid:` reference; the bytes live in a MIME part the HTML flavour does not carry |
| `WORD_UNRESOLVED_IMAGE` | APPROXIMATED / EQUIVALENT | no bytes available → labelled placeholder. `EQUIVALENT` when bytes *were* recovered from a clipboard image item, in which case the message names the pairing index, because the pairing is positional rather than certain |
| `WORD_IMAGE_CROP_APPROXIMATED` | APPROXIMATED | Word cropped the picture; HTML has no crop, so a CSS clip is applied and the percentages kept |
| `WORD_FLOATING_IMAGE_APPROXIMATED` | APPROXIMATED | Word anchors floating pictures to the page or margin; HTML can only float within the text column |

### Drawings and objects

| code | fidelity | when |
|---|---|---|
| `WORD_VML_OBJECT` | EQUIVALENT | a VML shape wrapping a picture — the picture is extracted, the VML kept on the model |
| `WORD_VML_SHAPE_APPROXIMATED` | UNSUPPORTED | a VML drawing with no picture inside. Vector drawings have no HTML equivalent; the markup is preserved, nothing is rendered |
| `WORD_OLE_OBJECT` | UNSUPPORTED | an embedded object. Metadata and fallback text kept, inert placeholder rendered |
| `WORD_ACTIVEX_OBJECT` | UNSUPPORTED | an ActiveX control |
| `WORD_SMARTART_UNSUPPORTED` | UNSUPPORTED | SmartArt |
| `WORD_CHART_UNSUPPORTED` | UNSUPPORTED | an embedded chart |
| `WORD_EQUATION_UNSUPPORTED` | UNSUPPORTED | an OMML equation |
| `WORD_UNSUPPORTED_SHAPE` | UNSUPPORTED | any other drawing object |
| `WORD_TEXT_BOX_APPROXIMATED` | APPROXIMATED | a text box or frame flattened into a bordered container; its floating position is lost |

### Structure

| code | fidelity | when |
|---|---|---|
| `WORD_SECTION_BREAK_APPROXIMATED` | APPROXIMATED | a section break rendered as a page break. Per-section page setup — size, margins, headers, columns — has no HTML equivalent |
| `WORD_NESTED_TABLE` | EXACT | a nested table was parsed as a structured table rather than passed through as markup. Informational |
| `WORD_TABLE_GRID_REPAIRED` | EQUIVALENT | a row did not fill the resolved grid; empty cells were added so the table stays rectangular. No content removed |
| `WORD_FOOTNOTE_APPROXIMATED` | APPROXIMATED | footnote/endnote containers lifted out of the flow and reported separately, rather than appended to the body as text |
| `WORD_REVISION_MARK_FLATTENED` | APPROXIMATED | tracked changes flattened into ordinary formatting |
| `WORD_CONTENT_CONTROL_APPROXIMATED` | APPROXIMATED | a content control rendered as its current value |
| `WORD_FORM_FIELD_UNSUPPORTED` | UNSUPPORTED | a legacy form field |
| `WORD_UNSUPPORTED_FIELD` | APPROXIMATED | a field whose result is kept but whose live behaviour is not — including review comments, whose anchors are removed and text reported separately |
| `WORD_FRAGMENT_BOUNDARY_MISSING` | EQUIVALENT | only one of `StartFragment`/`EndFragment` was present; the whole body was used |
| `WORD_NAMESPACE_ELEMENT_DROPPED` | EQUIVALENT | an Office-namespaced element with no content contribution |
| `WORD_UNKNOWN_ELEMENT` | EQUIVALENT | an element the parser has no rule for; its children were still walked |
| `WORD_CSS_PARSE_WARNING` | EQUIVALENT | a malformed CSS construct was skipped |

### Security

All security diagnostics report something that was **removed**. The raw payload
is untouched on `document.rawHtml` if you need to inspect what it was.

| code | fidelity | when |
|---|---|---|
| `SECURITY_SCRIPT_REMOVED` | EQUIVALENT | `<script>` or another executing element. Clipboard HTML is never executed |
| `SECURITY_EVENT_HANDLER_REMOVED` | EQUIVALENT | an `on*` attribute |
| `SECURITY_URL_BLOCKED` | EQUIVALENT | a `javascript:`, `vbscript:` or `data:` URL in a position where it is unsafe. The original target is kept on the model as `rawHref` and rendered as an inert `data-word-unresolved-href` |
| `SECURITY_EXTERNAL_RESOURCE` | EQUIVALENT | content pointing at a third-party host. Not blocked — reported, so a host application can decide |

### Limits

| code | fidelity | when |
|---|---|---|
| `LIMIT_NODE_BUDGET_EXCEEDED` | UNSUPPORTED | the node budget was spent; parsing stopped early and the result is truncated |
| `LIMIT_DOCUMENT_TRUNCATED` | UNSUPPORTED | the payload or block budget was exceeded |
| `LIMIT_DEPTH_EXCEEDED` | UNSUPPORTED | nesting exceeded `maxDepth` / `maxTableDepth` |

### Detection

| code | fidelity | when |
|---|---|---|
| `NOT_WORD_CONTENT` | EQUIVALENT | the payload did not look like Word. Parsed with the generic path; Word-specific rules still applied where signals were present |

## The honest limits

Things this engine deliberately does not claim to do:

- **Page layout.** Columns, headers, footers, page size and margins are read into
  the model and reported, not rendered. HTML has no page.
- **Vector drawings.** VML shapes that are not pictures are preserved as markup
  and diagnosed, not converted to SVG.
- **Equations.** OMML is not translated to MathML.
- **Fonts.** The output names the fonts Word named. Whether the reader has
  Wingdings installed is not something the engine can control — which is exactly
  why symbol bullets are mapped to Unicode rather than left as font-dependent
  bytes.
- **Tracked changes and comments.** Flattened and lifted out respectively, both
  diagnosed. Representing them would need a model of revision state that the
  clipboard payload does not fully carry.

Where a limit is reached, the model still holds everything the payload said. A
future back end can do better with the same model without re-parsing anything.
