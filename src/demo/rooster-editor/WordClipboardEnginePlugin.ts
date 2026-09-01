import type {
  BeforePasteEvent,
  ContentChangedEvent,
  EditorPlugin,
  IEditor,
  PluginEvent,
} from 'roosterjs-content-model-types';
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
  // Needed only for renumberCompositeMarkers (§ below) to reach the editor's
  // live DOM after a paste — handleBeforePaste itself still touches nothing
  // but event.fragment / event.clipboardData.
  private editor: IEditor | null = null;

  getName(): string {
    return 'WordClipboardEngine';
  }

  initialize(editor: IEditor): void {
    this.editor = editor;
    // A "final settle" for renumberCompositeMarkers's own caret-skipping
    // (see `activeItem` there): the very last item someone types stays
    // unnumbered until *some* later edit runs the renumbering pass again —
    // fine while they keep adding items, not fine if that was the last one
    // and they just click away. `focusout` catches that: once focus has
    // genuinely left the editor (checked a tick later, since the outgoing
    // target is still "focused" for this same event), nothing is being
    // actively typed into any more, so every item can be numbered.
    document.addEventListener('focusout', this.handleFocusOut, true);
  }

  dispose(): void {
    document.removeEventListener('focusout', this.handleFocusOut, true);
    this.editor = null;
  }

  private handleFocusOut = (): void => {
    queueMicrotask(() => {
      if (this.editor && !this.editor.getDOMHelper().hasFocus()) {
        renumberCompositeMarkers(this.editor);
      }
    });
  };

  onPluginEvent(event: PluginEvent): void {
    if (event.eventType === 'beforePaste') {
      this.handleBeforePaste(event);
    } else if (event.eventType === 'contentChanged') {
      this.handleContentChanged(event);
    }
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
      // 'native' markers here, not the library's own default ('element'):
      // this editor needs a real <ol>/<li> with a real browser ::marker, not
      // a text span glued to the paragraph, because only a real list is
      // something RoosterJS's own list-continuation logic understands — press
      // Enter at the end of an item and RoosterJS inserts the next <li>,
      // which the browser numbers on its own. An 'element' marker is baked-in
      // text: RoosterJS has no way to know "insert the next number here",
      // because there is no next number in its model, just a span someone
      // wrote once and never renders again.
      //
      // This also sidesteps the corruption risk 'element' mode needed
      // contenteditable="false" for in the first place: a native ::marker is
      // drawn by the browser outside the editable content entirely — there is
      // no DOM node inside the editable flow for a stray keystroke to land in
      // or a Backspace to eat character-by-character.
      //
      // The tradeoff (see HtmlRenderer.ts) is Word's exact gutter spacing:
      // 'native' markers are right-aligned in their gutter with no visible
      // gap ("1. Saji George"), where 'element' reproduces Word's actual
      // tab-stop gap ("1.1.1    Saji George"). Auto-continuing numbers in a
      // live editor is the feature being traded for here.
      const rendered = renderWordDocument(wordDocument, { markerMode: 'native', includeWordMetadata: false });
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

  /**
   * Keep composite markers ("1.1", "1.1.1" — § renumberCompositeMarkers)
   * numbered correctly as the list they belong to is edited.
   *
   * Skipped for the change that IS the paste itself ('Paste' — the literal
   * value of RoosterJS's own `ChangeSource.Paste`, not re-imported here to
   * avoid a dependency on `roosterjs-content-model-dom` for one string
   * constant): those markers came straight from Word's own clipboard HTML,
   * already numbered exactly as Word had them (including a snippet pasted
   * from the middle of a longer document, starting at "3.4" rather than
   * "1.1") — recomputing from structural position on that same event would
   * overwrite a value the paste already got right.
   */
  private handleContentChanged(event: ContentChangedEvent): void {
    if (!this.editor || event.source === 'Paste') return;
    renumberCompositeMarkers(this.editor);
  }
}

/**
 * Recompute every composite list marker's displayed number from its current
 * position in the DOM, in place.
 *
 * A composite marker — Word's "1.1", "1.1.1" style level text ("%1.%2…"),
 * which no single CSS counter-style can express (see HtmlListRenderer.ts) —
 * renders as a static `element`-mode span even under this plugin's `native`
 * marker mode. A real `<ol>`/`<li>` auto-continues on its own for a *simple*
 * counter (one browser `::marker`, one number); it does nothing for a
 * composite one, because there is no single counter to continue — the text
 * is just characters in a `contenteditable="false"` span, same as any other
 * static content, until something rewrites it. This is that something.
 *
 * There is no metadata left to work from by the time an edit happens: the
 * Word-derived template that produced "1.1" (`%1.%2`, decoded per level's own
 * number format) lived in `data-word-*` attributes that RoosterJS's paste
 * conversion already strips (the same mechanism that made the marker vanish
 * entirely before HtmlListRenderer.ts's `element` fallback — see
 * docs/RICH-TEXT-EDITOR.md § 3), and DOM node identity does not survive
 * RoosterJS's own DOM → Content Model → DOM round trip either, so nothing can
 * be cached against the original elements. Only the *current* marker text is
 * available, so recomputation has to be self-describing: split it on its own
 * digit runs to recover a template ("1.1." -> literal parts `["", ".", "."]`
 * around two digit placeholders), then refill those placeholders from the
 * item's actual position in the list — which is real, live DOM structure, so
 * it is always correct regardless of what RoosterJS did to everything else.
 * The literal separators (typically ".", sometimes a trailing suffix) are
 * preserved exactly as they were, never reinterpreted.
 *
 * Scope, honestly: this only recovers *decimal* composite numbering ("1.1"),
 * the overwhelmingly common real-world case (business/legal/technical
 * section numbering — including this file's own worked example). A composite
 * mixing formats (Word's rarer "I.1.a", roman-then-decimal-then-alpha) has no
 * digit run for its roman or alpha segments and is left untouched, same as
 * today: still a real, protected marker, just not one this function knows how
 * to renumber. And because recomputation always starts from a document's
 * current, position-based numbering, an edit anywhere resequences every
 * composite marker in the document to start from 1 at each level — a paste
 * that started mid-document at "3.4" stays "3.4" until the next edit, then
 * renumbers like everything else. Preserving that offset indefinitely would
 * need the same Word-level metadata this function has already established
 * does not survive the round trip.
 */
function renumberCompositeMarkers(editor: IEditor): void {
  // The item the caret is currently in, if any — see the note on
  // `activeItem` below for why new markers skip it. Only while the editor
  // itself has focus: a stale `Selection` range left over from before a
  // blur (verified directly — clicking outside the editor does not clear
  // it, only moves DOM focus) must not keep excluding that item forever,
  // since nothing can still be mid-keystroke in it once focus is gone.
  const activeItem = editor.getDOMHelper().hasFocus() ? currentListItem() : null;

  for (const list of editor.getDOMHelper().queryElements('ol, ul')) {
    const items = Array.from(list.children).filter(
      (child): child is HTMLElement => child.tagName === 'LI' && !child.classList.contains('wce-list-spacer'),
    );

    // A list is "managed" only if at least one of its current items already
    // has a composite marker — that marker is both the signal that this
    // level needs renumbering at all, and (kept as `sampleMarker`, not just
    // its text) the source of everything a freshly synthesised marker needs
    // to look identical to one that survived a paste: the digit-run
    // template AND the `style` that reserves this level's gutter width
    // (`display:inline-block;min-width:…`, inlined once by
    // `inlineGeneratedStyles` at paste time — a synthesised marker has no
    // other way to learn that width, since it never goes through that step).
    const sampleMarker = items.map(ownMarker).find((marker): marker is HTMLElement => marker !== null);
    if (!sampleMarker) continue;
    const template = digitRunTemplate(sampleMarker.textContent ?? '');
    if (template.placeholderCount < 1) continue;

    for (const item of items) {
      const positions = listPositions(item);
      // A shape mismatch (this item sits at a different nesting depth than
      // the template expects) means guessing would produce a wrong number
      // with total confidence — leave it alone rather than risk that.
      if (positions.length !== template.placeholderCount) continue;

      const next = fillDigitRunTemplate(template, positions);
      const marker = ownMarker(item);
      if (marker) {
        if (marker.textContent !== next) marker.textContent = next;
      } else if (item !== activeItem) {
        // A new item RoosterJS created by splitting an existing one (Enter
        // at the end of a composite-numbered item) inherits the item's
        // formatting but not its marker span — nothing about "the next
        // item needs a marker" is part of what RoosterJS split. This list
        // is already known composite (a sibling supplied the template
        // above), so this item belongs in the same sequence; give it one —
        // unless it is the item the caret is in right now (see `activeItem`
        // above): inserting into a node RoosterJS's own selection tracking
        // is actively managing races the next keystroke, verified directly
        // — the marker's DOM position drifted mid-word when tried. Leaving
        // the live line unnumbered until the caret moves on to the next
        // edit (typically the next Enter) avoids that race entirely, and
        // every already-settled line is still numbered immediately.
        insertionParent(item).prepend(createMarkerElement(next, sampleMarker));
      }
    }
  }
}

/** The `<li>` the caret is currently positioned in, if any. */
function currentListItem(): HTMLElement | null {
  const anchor = window.getSelection()?.anchorNode;
  if (!anchor) return null;
  const element = anchor.nodeType === Node.ELEMENT_NODE ? (anchor as Element) : anchor.parentElement;
  return element?.closest('li') ?? null;
}

/**
 * The `.wce-marker` that belongs to `item` itself, as opposed to one
 * belonging to a nested sub-list inside it.
 *
 * A plain `:scope > .wce-marker` check is not enough: verified directly,
 * RoosterJS does not keep the marker as a direct child of `<li>` in every
 * case — after some edits (a split from pressing Enter mid-item, observed
 * directly) it ends up one level deeper, inside a `<div role="presentation">`
 * RoosterJS itself wraps the item's content in. Searching all descendants
 * and confirming the match's nearest enclosing `<li>` is still this one
 * (not a deeper item's, reached through a nested `<ol>`/`<ul>`) is robust to
 * that either way, since an item's own marker — wrapped or not — always
 * comes before any nested list's content in document order, so a plain
 * `querySelector` finds it first regardless.
 */
function ownMarker(item: HTMLElement): HTMLElement | null {
  const marker = item.querySelector<HTMLElement>('.wce-marker');
  return marker && marker.closest('li') === item ? marker : null;
}

/**
 * Where a freshly synthesised marker should be inserted so it renders
 * inline with the item's text instead of on its own line.
 *
 * Mirrors `ownMarker`'s reasoning: an item RoosterJS just created by
 * splitting another sometimes wraps its whole text content in one
 * `role="presentation"` `<div>` rather than leaving it as a direct text
 * child of the `<li>`. Since that div is block-level, inserting the marker
 * as a sibling before it — the natural reading of "prepend to this item" —
 * would put the number and the text on two different lines. Descending into
 * a lone such wrapper keeps the marker on the same line as the text it
 * marks, matching every marker this engine's own renderer ever produces.
 */
function insertionParent(item: HTMLElement): HTMLElement {
  const only = item.firstElementChild;
  return only instanceof HTMLDivElement && item.children.length === 1 ? only : item;
}

/**
 * Match `HtmlListRenderer.ts`'s own `renderMarkerElement` output, so a
 * freshly synthesised marker is indistinguishable from one that survived a
 * paste — critically including its `style`: the gutter width
 * (`display:inline-block;min-width:…`) that keeps this level's text lined up
 * a fixed distance after the marker comes only from `style`, inlined once by
 * `inlineGeneratedStyles` at paste time. Nothing else in a freshly created
 * marker would carry it — verified directly: without copying it, a
 * synthesised marker rendered with no gap at all before the item's text,
 * unlike every marker that survived a paste.
 */
function createMarkerElement(text: string, like: HTMLElement): HTMLElement {
  const marker = document.createElement('span');
  marker.className = 'wce-marker';
  marker.style.cssText = like.style.cssText;
  marker.setAttribute('contenteditable', 'false');
  marker.setAttribute('aria-hidden', 'true');
  marker.textContent = text;
  return marker;
}

interface DigitRunTemplate {
  /** Text between/around the digit runs, one longer than the digit runs it separates. */
  literalParts: string[];
  placeholderCount: number;
}

/** Split "1.1." into the literal parts around its digit runs: `["", ".", "."]`, 2 placeholders. */
function digitRunTemplate(text: string): DigitRunTemplate {
  const literalParts = text.split(/\d+/);
  return { literalParts, placeholderCount: literalParts.length - 1 };
}

function fillDigitRunTemplate(template: DigitRunTemplate, values: number[]): string {
  let out = template.literalParts[0] ?? '';
  for (let index = 0; index < values.length; index++) {
    out += String(values[index]) + (template.literalParts[index + 1] ?? '');
  }
  return out;
}

/**
 * An item's 1-based position within its own list, at every nesting level from
 * the outermost list down to the item itself — e.g. the 2nd sub-item of the
 * 1st item returns `[1, 2]`.
 *
 * This is what a composite marker's digit placeholders are filled from: the
 * outer levels are exactly the ancestor counters Word's `%1.%2…` level text
 * references, and ordinary DOM sibling position is what stays correct after
 * any edit, independent of whatever RoosterJS did to attributes or node
 * identity elsewhere in the tree.
 */
function listPositions(item: Element): number[] {
  const positions: number[] = [];
  let current: Element | null = item;
  while (current) {
    const list: HTMLElement | null = current.parentElement;
    if (!list || (list.tagName !== 'OL' && list.tagName !== 'UL')) break;
    positions.unshift(siblingPosition(current, list));
    current = listContainer(list);
  }
  return positions;
}

/**
 * The `<li>` a nested `<ol>`/`<ul>` logically belongs to — the item whose
 * sub-list it is.
 *
 * `HtmlRenderer.ts` always emits proper nesting, `<li>…<ol>…</ol></li>`, and
 * that shape is what a static render or a download keeps. RoosterJS's own
 * content model does not: verified directly, after a real paste through this
 * plugin, a nested list under "Background" comes back out as a *sibling* of
 * `<li>Background</li>` — `<ol><li>Background</li><ol>…sub-items…</ol><li>
 * Scope</li></ol>` — not a child of it, because RoosterJS's list
 * representation tracks nesting as a per-item level stack rather than actual
 * DOM containment, and re-renders it as consecutive `<ol>` runs at
 * increasing indentation instead of one nested inside the other. (It does
 * this inconsistently — a *single*-item sub-list, like "Given"'s one
 * sub-point in this same document, came back properly nested. Handling both
 * shapes is simpler than predicting which one RoosterJS will choose.)
 *
 * So the owning item is either the list's real parent `<li>` (the shape this
 * engine's own renderer produces, and what a non-RoosterJS consumer would
 * keep), or — RoosterJS's shape — the nearest `<li>` immediately before this
 * `<ol>`/`<ul>` among its own siblings.
 */
function listContainer(list: Element): Element | null {
  const parent = list.parentElement;
  if (!parent) return null;
  if (parent.tagName === 'LI') return parent;
  if (parent.tagName !== 'OL' && parent.tagName !== 'UL') return null;

  let sibling: Element | null = list.previousElementSibling;
  while (sibling && sibling.tagName !== 'LI') sibling = sibling.previousElementSibling;
  return sibling;
}

/** 1-based position of `item` among its counted `<li>` siblings in `list`. */
function siblingPosition(item: Element, list: Element): number {
  let position = 0;
  for (const child of Array.from(list.children)) {
    if (child.tagName !== 'LI' || child.classList.contains('wce-list-spacer')) continue;
    position++;
    if (child === item) return position;
  }
  return position;
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
