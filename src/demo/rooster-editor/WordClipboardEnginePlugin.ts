import type { BeforePasteEvent, EditorPlugin, PluginEvent } from 'roosterjs-content-model-types';
import {
  clipboardPayloadFromHtml,
  detectWordHtml,
  parseWordClipboard,
  renderWordDocument,
} from '../../index.js';

/**
 * A RoosterJS paste plugin backed by word-clipboard-engine.
 *
 * This is the integration point the spec always intended: the engine itself
 * stays a plain library with zero runtime dependencies and no idea any editor
 * exists (see `docs/ARCHITECTURE.md` — "nothing in the model names an
 * editor"). This file is the only place in the whole codebase that imports
 * `roosterjs-*`, and it lives in `src/demo/`, outside the published package —
 * `word-clipboard-engine`'s own `dist/index.js` never contains a line of
 * RoosterJS. What ships here is an *adapter*: a plugin, in RoosterJS's own
 * plugin shape, that hands RoosterJS's default Word-paste handling off to
 * this engine instead.
 *
 * RoosterJS gives a plugin exactly one hook for this — `BeforePasteEvent`,
 * fired after the browser's native paste but before RoosterJS converts the
 * pasted fragment into its own Content Model. `event.clipboardData.rawHtml`
 * is the untouched `text/html` clipboard flavour (RoosterJS's own name for
 * what this engine calls `payload.html`); `event.fragment` is the
 * `DocumentFragment` RoosterJS is about to read from. Replacing its contents
 * before that read is the whole mechanism — no monkey-patching, no
 * intercepting the native `paste` event ourselves, just the seam RoosterJS
 * already ships for exactly this purpose.
 */
export class WordClipboardEnginePlugin implements EditorPlugin {
  getName(): string {
    return 'WordClipboardEngine';
  }

  // This plugin only reads event.fragment / event.clipboardData; it never
  // calls back into the editor, so unlike most EditorPlugin implementations
  // it has no editor reference to keep.
  initialize(): void {}

  dispose(): void {}

  onPluginEvent(event: PluginEvent): void {
    if (event.eventType !== 'beforePaste') return;
    this.handleBeforePaste(event);
  }

  private handleBeforePaste(event: BeforePasteEvent): void {
    // 'mergeFormat' and 'asPlainText' are the user explicitly asking for
    // something other than "paste it as Word had it" — deferring to
    // RoosterJS's own handling for those is the correct behaviour, not a
    // gap. 'asImage' and 'asMarkdown' are unrelated paste shapes entirely.
    if (event.pasteType !== 'normal') return;

    const rawHtml = event.clipboardData.rawHtml;
    if (!rawHtml) return;

    // Screen out non-Word HTML (a webpage, a screenshot's alt text, plain
    // copy from a terminal) so RoosterJS's own general-purpose paste
    // handling — which is good at things this engine deliberately isn't,
    // like normalising arbitrary web markup — still runs for everything
    // that isn't actually a Word payload.
    if (!detectWordHtml(rawHtml).isWord) return;

    let html: string;
    try {
      // Named to avoid shadowing the global `document` — this is a parsed
      // WordDocument model, not a DOM Document.
      const wordDocument = parseWordClipboard(clipboardPayloadFromHtml(rawHtml));
      // 'element' markers (the engine's default) are what make this safe to
      // drop into a *live* editable surface: the marker is a real element
      // with contenteditable="false" on it, so typing or backspacing next to
      // a bullet can delete the marker as a whole but can never partially
      // corrupt it into "· becomes -" or "1.1.1 becomes 1.1." the way a bare
      // text-node glyph could. Verified directly against Chromium's actual
      // editing behaviour — see docs/RICH-TEXT-EDITOR.md.
      const rendered = renderWordDocument(wordDocument, { markerMode: 'element', includeWordMetadata: false });
      html = rendered.html;
      // The generated marker/list geometry CSS is per-document, not per-node,
      // so it is installed once here rather than threaded through the
      // fragment — see `ensureWordClipboardStylesheet` below.
      ensureWordClipboardStylesheet(rendered.css);
    } catch {
      // A parse failure here must fall through to RoosterJS's own paste
      // handling, not swallow the paste. `parseWordClipboard` runs the
      // security scrub before this plugin ever sees a node, so a caught
      // failure means something about the payload was unparseable, not that
      // this plugin is choosing to render something unsafe.
      return;
    }

    const parsed = new DOMParser().parseFromString(html, 'text/html');
    const root = parsed.body.firstElementChild;
    if (!root) return;

    // RoosterJS converts the pasted fragment into its own Content Model
    // before ever rendering it back to DOM, and that conversion keeps the
    // inline `style` it finds on each element but drops arbitrary CSS
    // classes — a class carrying `list-style:none` or `padding-left` from an
    // external stylesheet is gone by the time anything is on screen, and
    // <ol>/<li> then falls back to the browser's own default numbering
    // stacked right on top of this engine's own marker text. Verified
    // directly: the same list-geometry CSS set as an inline `style`
    // attribute round-trips through RoosterJS's paste pipeline intact; set
    // as a class on an external stylesheet, it does not. So every element in
    // this fragment gets its matching declarations inlined before RoosterJS
    // ever sees it.
    inlineGeneratedStyles(root);

    while (event.fragment.firstChild) event.fragment.removeChild(event.fragment.firstChild);
    // `root` (the engine's own <div class="wce-document"> wrapper) becomes
    // the fragment's only child; RoosterJS reads block content out of it the
    // same way it would read any pasted <div>.
    event.fragment.appendChild(root);
  }
}

/**
 * Apply every declaration the generated stylesheet has for elements under
 * `root`, as inline `style`, without disturbing whatever inline style an
 * element already carries.
 *
 * Reuses `installedStylesheet` (already parsed once by the browser as a real
 * `CSSStyleSheet` via `ensureWordClipboardStylesheet`) rather than
 * hand-rolling a second CSS parser — the browser's own selector matching via
 * `querySelectorAll` is exactly as correct on a detached `DocumentFragment`
 * subtree as it is on a live one.
 */
function inlineGeneratedStyles(root: Element): void {
  const sheet = installedStylesheet?.sheet;
  if (!sheet) return;

  for (const rule of Array.from(sheet.cssRules)) {
    // Only plain style rules (`.foo { … }`) apply here. `@counter-style` and
    // other at-rules (only emitted in 'native' marker mode, which this
    // plugin never requests) have no selector to match against.
    if (!(rule instanceof CSSStyleRule)) continue;

    let matches: Element[];
    try {
      matches = root.matches(rule.selectorText)
        ? [root, ...Array.from(root.querySelectorAll(rule.selectorText))]
        : Array.from(root.querySelectorAll(rule.selectorText));
    } catch {
      // A selector this engine never generates but a future version might —
      // skip it rather than let one bad selector abort the whole paste.
      continue;
    }

    for (const element of matches) {
      const htmlElement = element as HTMLElement;
      for (const property of Array.from(rule.style)) {
        if (htmlElement.style.getPropertyValue(property)) continue; // inline wins
        htmlElement.style.setProperty(property, rule.style.getPropertyValue(property));
      }
    }
  }
}

/**
 * The CSS the engine generates per document (native counter styles in
 * `native` marker mode, or the marker-geometry rules `element` mode always
 * needs) is not part of `event.fragment` — `renderWordDocument` returns it
 * separately by design, so a caller decides where it lives rather than the
 * engine injecting a `<style>` tag into content no one asked for one in.
 *
 * Call this once per editor (or once per app) rather than once per paste: the
 * class names it generates are content-addressed, so calling it again after a
 * second paste only ever *adds* rules for classes the sheet doesn't already
 * have — it never needs to clear and rebuild.
 */
let installedStylesheet: HTMLStyleElement | null = null;

export function ensureWordClipboardStylesheet(css: string): void {
  if (!installedStylesheet) {
    installedStylesheet = document.createElement('style');
    installedStylesheet.setAttribute('data-word-clipboard-engine', '');
    document.head.appendChild(installedStylesheet);
  }
  // Append only rules the sheet doesn't already have, keyed by the whole rule
  // text — cheap and correct, since the generated class names are already
  // content-addressed and a rule never needs to change once emitted.
  const existing = new Set(
    installedStylesheet.textContent!.split('\n').filter((line) => line.trim().length > 0),
  );
  const additions = css.split('\n').filter((line) => line.trim().length > 0 && !existing.has(line));
  if (additions.length > 0) {
    installedStylesheet.textContent += (installedStylesheet.textContent ? '\n' : '') + additions.join('\n');
  }
}
