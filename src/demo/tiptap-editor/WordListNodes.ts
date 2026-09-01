import { Node, mergeAttributes } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { splitListItem } from '@tiptap/pm/schema-list';
import type { Node as PMNode } from '@tiptap/pm/model';
import { expandLevelText, type WordNumberFormat } from '../../index.js';

/**
 * The core of the TipTap integration: two node types, `wordList` and
 * `wordListItem`, that carry Word's own numbering *declaration* — not its
 * rendered text — as node attributes, plus a plugin that draws every
 * marker's text fresh from those attributes on every document read.
 *
 * This is the fix for what RoosterJS's integration could never fully solve.
 * RoosterJS converts pasted content into its own fixed Content Model, which
 * has no field for "this item's marker is `%1.%2`, formatted decimal then
 * decimal" — only a curated set of formatting properties it already knows
 * about. Nothing pasted through it can carry Word-specific numbering data
 * through an edit, which is why that integration needed a separate,
 * DOM-scraping renumbering pass bolted on afterward (see
 * `../rooster-editor/WordClipboardEnginePlugin.ts`), and why that pass is
 * still limited to decimal composites — it has no schema to lean on, only
 * text it can pattern-match.
 *
 * TipTap (ProseMirror) has no such fixed model — *this file defines the
 * schema* — so the numbering declaration is never lost, never needs
 * recovering from rendered text, and covers every Word number format
 * (roman, alpha, ordinal, cardinal-text, any mix at any depth), not just
 * decimal chains. A marker is drawn by asking "what number is the Nth item
 * of the Mth nested list, in the format this level declares" — the same
 * question `expandLevelText`/`formatNumber` already answer for the static
 * renderer (`HtmlListRenderer.ts`) — every time the document changes, using
 * nothing but real ProseMirror document structure. There is no DOM to
 * scrape and no race with a framework re-rendering nodes it does not
 * recognise, because every node here is one this schema owns outright.
 */

export interface WordListAttrs {
  listId: string | null;
  level: number;
  startAt: number;
  ordered: boolean;
  style: string | null;
}

export interface WordListItemAttrs {
  listId: string | null;
  level: number;
  levelText: string | null;
  numberFormat: WordNumberFormat | null;
  markerFont: string | null;
  /** The literal marker text Word rendered, captured at parse time. Used
   * verbatim for a bullet or any level text with no `%N` placeholder — a
   * bullet's glyph never changes with position, so there is nothing to
   * recompute, and reusing the captured text sidesteps the private-use-area
   * placeholder (`U+F0B7`) `data-word-level-text` carries for a glyph this
   * engine could not map to a printable Unicode character. */
  initialMarkerText: string | null;
  /** The marker span's own `style` (`display:inline-block;min-width:…`) —
   * the reserved gutter width that makes Word's tab-stop gap survive. It has
   * to be captured separately from the marker's text: the `ignore: true`
   * parse rule below (needed so the marker's *text* is never read into the
   * document as ordinary content) drops the whole span, geometry included,
   * so nothing else has a chance to keep it. */
  markerStyle: string | null;
  style: string | null;
}

export const WordList = Node.create({
  name: 'wordList',
  group: 'block',
  content: 'wordListItem+',

  // Every attribute here is read back out explicitly in `renderHTML` below,
  // as exactly the plain HTML attribute Word/this engine would produce
  // (`style`, `start`) — none of them are meant to round-trip as their own
  // named HTML attribute (`listid="…"`, `startat="…"`), which is what
  // TipTap's own default `renderHTML` per attribute would otherwise do.
  // `renderHTML: () => ({})` suppresses that default.
  addAttributes() {
    return {
      listId: { default: null, parseHTML: (el) => el.getAttribute('data-word-list'), renderHTML: () => ({}) },
      level: {
        default: 0,
        parseHTML: (el) => Number.parseInt(el.getAttribute('data-word-level') ?? '1', 10) - 1,
        renderHTML: () => ({}),
      },
      startAt: {
        default: 1,
        parseHTML: (el) => Number.parseInt(el.getAttribute('start') ?? '1', 10) || 1,
        renderHTML: () => ({}),
      },
      ordered: { default: false, parseHTML: (el) => el.tagName === 'OL', renderHTML: () => ({}) },
      style: { default: null, parseHTML: (el) => el.getAttribute('style'), renderHTML: () => ({}) },
    };
  },

  parseHTML() {
    return [
      { tag: 'ol[data-word-list]', priority: 60 },
      { tag: 'ul[data-word-list]', priority: 60 },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const tag = node.attrs.ordered ? 'ol' : 'ul';
    const attrs: Record<string, string> = {};
    if (node.attrs.style) attrs.style = node.attrs.style;
    if (node.attrs.ordered && node.attrs.startAt > 1) attrs.start = String(node.attrs.startAt);
    return [tag, mergeAttributes(attrs, HTMLAttributes), 0];
  },
});

export const WordListItem = Node.create({
  name: 'wordListItem',
  group: 'block',
  content: 'paragraph+ wordList*',
  defining: true,

  // As with `wordList` above: every attribute is internal bookkeeping, read
  // back out explicitly (as `style`, plus a synthesised marker element) in
  // `renderHTML`/`createMarkerElement`, not meant to round-trip as its own
  // named HTML attribute.
  addAttributes() {
    return {
      listId: { default: null, parseHTML: (el) => el.getAttribute('data-word-list'), renderHTML: () => ({}) },
      level: {
        default: 0,
        parseHTML: (el) => Number.parseInt(el.getAttribute('data-word-level') ?? '1', 10) - 1,
        renderHTML: () => ({}),
      },
      levelText: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-word-level-text'),
        renderHTML: () => ({}),
      },
      numberFormat: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-word-number-format'),
        renderHTML: () => ({}),
      },
      markerFont: {
        default: null,
        parseHTML: (el) => el.getAttribute('data-word-marker-font'),
        renderHTML: () => ({}),
      },
      initialMarkerText: {
        default: null,
        parseHTML: (el) => el.querySelector(':scope > .wce-marker')?.textContent ?? null,
        renderHTML: () => ({}),
      },
      markerStyle: {
        default: null,
        parseHTML: (el) => el.querySelector(':scope > .wce-marker')?.getAttribute('style') ?? null,
        renderHTML: () => ({}),
      },
      style: { default: null, parseHTML: (el) => el.getAttribute('style'), renderHTML: () => ({}) },
    };
  },

  parseHTML() {
    return [{ tag: 'li.wce-list-item', priority: 60 }];
  },

  renderHTML({ node, HTMLAttributes }) {
    const attrs: Record<string, string> = { class: 'wce-list-item' };
    if (node.attrs.style) attrs.style = node.attrs.style;
    return ['li', mergeAttributes(attrs, HTMLAttributes), 0];
  },

  addProseMirrorPlugins() {
    return [wordListMarkerPlugin];
  },

  /**
   * Without this, Enter falls through to ProseMirror's default `splitBlock`
   * — which, because `content` above allows more than one `paragraph` per
   * item (`paragraph+`, for the rare genuinely multi-paragraph list item),
   * just adds a second paragraph *inside the same `wordListItem`* rather
   * than starting a new item. Verified directly: without this handler,
   * pressing Enter at the end of a pasted item silently added the typed
   * text to that same item, with no new marker anywhere — not broken
   * rendering, just never becoming a new list item at all.
   * `splitListItem`, prosemirror-schema-list's own command for exactly this
   * (used under the hood by every list implementation, this schema's own
   * disabled `listItem` extension included), splits the *item*, and with no
   * explicit `itemAttrs` passed, copies the current item's attrs onto the
   * new one — the correct default here, since a fresh item is always a
   * sibling in the very same Word list.
   */
  addKeyboardShortcuts() {
    return {
      Enter: () => splitListItem(this.type)(this.editor.state, this.editor.view.dispatch),
    };
  },
});

/**
 * The marker span survives paste as ordinary content unless something tells
 * ProseMirror's parser to skip it — the schema's own node/mark `parseHTML`
 * rules only say what *to* keep, not what to drop.
 * `WordClipboardExtension.ts` installs this alongside the schema-derived
 * rules, ahead of them (parse rules are tried in order, first match wins),
 * so the marker's *rendered* text is never read into the document — only
 * its geometry `style`, captured onto
 * `wordListItem.initialMarkerText`/`markerStyle` above by the node's own
 * `getAttrs`, which run before this rule's `ignore` takes effect for the
 * span's children.
 */
export const ignoreMarkerParseRule = { tag: 'span.wce-marker', ignore: true } as const;

const markerPluginKey = new PluginKey('wordListMarkers');

/**
 * Draw every list marker fresh, on every document read.
 *
 * `Decoration.widget`, not a NodeView on `wordListItem` itself: a NodeView
 * would need `ignoreMutation`/manual DOM management to keep the marker
 * `contenteditable="false"`, and would still need this exact computation
 * inside it. A widget decoration gets both for free — ProseMirror already
 * treats a widget as opaque, non-editable content, the same guarantee
 * `contenteditable="false"` gave the RoosterJS integration's marker spans,
 * except here it is the framework's own mechanism rather than an attribute
 * this schema has to remember to set.
 */
export const wordListMarkerPlugin = new Plugin({
  key: markerPluginKey,
  props: {
    decorations(state) {
      return computeMarkerDecorations(state.doc);
    },
  },
});

function computeMarkerDecorations(doc: PMNode): DecorationSet {
  const decorations: Decoration[] = [];

  function walk(node: PMNode, pos: number, counters: number[]): void {
    if (node.type.name === 'wordList') {
      let index = 0;
      node.forEach((item, offset) => {
        if (item.type.name !== 'wordListItem') return;
        index++;
        const itemPos = pos + 1 + offset;
        const ownCounters = [...counters, (node.attrs.startAt as number) + index - 1];

        emitMarker(item, itemPos, ownCounters, decorations);

        item.forEach((child, childOffset) => walk(child, itemPos + 1 + childOffset, ownCounters));
      });
      return;
    }

    node.forEach((child, offset) => walk(child, pos + 1 + offset, counters));
  }

  walk(doc, -1, []);
  return DecorationSet.create(doc, decorations);
}

function emitMarker(item: PMNode, itemPos: number, counters: number[], decorations: Decoration[]): void {
  const firstChild = item.firstChild;
  if (!firstChild || firstChild.type.name !== 'paragraph') return;

  const levelText = (item.attrs.levelText as string | null) ?? '';
  const hasPlaceholder = /%\d/.test(levelText);
  // A composite level text's *own* number format applies to every
  // placeholder it contains, including ones that name an ancestor level —
  // not each ancestor's own independent format. Verified directly against a
  // real Word document: a roman-numbered level 0 ("I.") followed by a
  // decimal-formatted level 1 whose level text is "%1.%2" renders "1.1", not
  // "I.1" — level 1's own `decimal` format governs both digits, because
  // Word's `%1` here does not mean "show the ancestor the way it shows
  // itself", it means "show the ancestor's current count, in the format
  // *this* level declared". So the same format is used for the whole
  // `counters` array, not a chain of each level's own.
  const format = (item.attrs.numberFormat as WordNumberFormat | null) ?? 'decimal';
  const formats: WordNumberFormat[] = new Array(counters.length).fill(format);
  const text = hasPlaceholder
    ? expandLevelText(levelText, counters, formats)
    : ((item.attrs.initialMarkerText as string | null) ?? levelText);
  if (!text) return;

  // itemPos + 1 enters the wordListItem's content, +1 again enters the first
  // paragraph's own content — where the marker widget needs to sit so it
  // renders as the first thing on the item's own line.
  const markerPos = itemPos + 2;
  decorations.push(
    Decoration.widget(markerPos, () => createMarkerElement(text, item.attrs as WordListItemAttrs), {
      side: -1,
      ignoreSelection: true,
    }),
  );
}

function createMarkerElement(text: string, attrs: WordListItemAttrs): HTMLElement {
  const marker = document.createElement('span');
  marker.className = 'wce-marker';
  marker.setAttribute('contenteditable', 'false');
  marker.setAttribute('aria-hidden', 'true');
  if (attrs.markerStyle) marker.setAttribute('style', attrs.markerStyle);
  if (attrs.markerFont) marker.style.fontFamily = attrs.markerFont;
  marker.textContent = text;
  return marker;
}
