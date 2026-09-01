import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { DOMParser as ProseMirrorDOMParser, Slice } from '@tiptap/pm/model';
import {
  clipboardPayloadFromHtml,
  detectWordHtml,
  parseWordClipboard,
  renderWordDocument,
} from '../../index.js';
import { ignoreMarkerParseRule } from './WordListNodes.js';

/**
 * The TipTap paste handler backed by word-clipboard-engine.
 *
 * Unlike `../rooster-editor/WordClipboardEnginePlugin.ts`, this one hands
 * ProseMirror's *own* `DOMParser` the rendered HTML directly, rather than
 * splicing a DOM fragment into the editor and hoping the editor's own
 * conversion keeps what matters. That difference is possible, and safe,
 * specifically because `WordListNodes.ts` gives ProseMirror's schema a real
 * node for a Word list item — parsing straight into the schema is no longer
 * lossy the way handing arbitrary attributes to RoosterJS's fixed Content
 * Model was.
 *
 * `element` marker mode, not `native`: this integration does not need the
 * browser's own `<ol>`/`::marker` auto-continuation RoosterJS's plugin
 * traded gutter spacing for, because `wordListMarkerPlugin`
 * (`WordListNodes.ts`) *is* this integration's auto-continuation — driven by
 * the schema, correct for every Word number format, not just decimal
 * composites. `element` mode also carries Word's exact hanging-indent gap in
 * its geometry (`padding-left`/`text-indent`/marker `min-width`), which is
 * what the code below inlines onto the pasted content before it ever reaches
 * `wordListItem`'s `parseHTML`.
 */
export const WordClipboardExtension = Extension.create({
  name: 'wordClipboard',

  addProseMirrorPlugins() {
    const schema = this.editor.schema;

    return [
      new Plugin({
        props: {
          handlePaste: (view, event) => {
            const rawHtml = event.clipboardData?.getData('text/html');
            if (!rawHtml) return false;

            // Non-Word HTML (a webpage, a screenshot's alt text) falls
            // through to ProseMirror's own default paste handling, which is
            // good at normalising arbitrary web markup in ways this engine
            // deliberately isn't.
            if (!detectWordHtml(rawHtml).isWord) return false;

            let html: string;
            let css: string;
            try {
              const wordDocument = parseWordClipboard(clipboardPayloadFromHtml(rawHtml));
              const rendered = renderWordDocument(wordDocument, {
                markerMode: 'element',
                includeWordMetadata: true,
              });
              html = rendered.html;
              css = rendered.css;
            } catch {
              // A parse failure must fall through to the editor's own paste
              // handling, not swallow the paste. `parseWordClipboard` runs
              // the security scrub before anything here sees a node, so a
              // caught failure means the payload was unparseable, not that
              // this handler is choosing to insert something unsafe.
              return false;
            }

            ensureWordClipboardStylesheet(css);

            const parsed = new DOMParser().parseFromString(html, 'text/html');
            const root = parsed.body.firstElementChild;
            if (!root) return false;

            // Same reasoning as the RoosterJS integration's own
            // `inlineGeneratedStyles`: the generated CSS is class-based,
            // meant to be installed once rather than repeated per element.
            // Here it additionally means `wordList`/`wordListItem`'s
            // `parseHTML` — which reads each element's own `style`
            // attribute, not classes resolved against a stylesheet — sees
            // Word's real geometry without needing to know this fragment's
            // generated class names at all.
            inlineGeneratedStyles(root);

            const parser = new ProseMirrorDOMParser(schema, [
              // Tried before the schema-derived rules (first match wins), so
              // the marker's own rendered text is never read into the
              // document as ordinary content — only `wordListItem`'s
              // `getAttrs` (which run against the still-intact DOM before
              // this rule discards the span) ever sees it.
              ignoreMarkerParseRule,
              ...ProseMirrorDOMParser.fromSchema(schema).rules,
            ]);
            // `parseSlice` (the usual choice for pasted content) computes
            // `openStart`/`openEnd` from how deep the first/last leaf sits
            // and trims wrapping nodes at those edges — the standard
            // "merge cleanly into whatever's at the caret" behaviour a paste
            // normally wants. Verified directly: it stripped the outer
            // `wordList` from every *top-level* item (the ones the trimming
            // considered "at the edge"), while a nested list further inside
            // the pasted content — never at the trimmed boundary — kept its
            // wrapper intact. `parse` (a full, self-contained document) plus
            // a `Slice` built with `openStart`/`openEnd` both `0` inserts
            // the structure exactly as rendered, no merging assumed.
            const parsedDocument = parser.parse(root, { preserveWhitespace: true });
            const slice = new Slice(parsedDocument.content, 0, 0);

            const tr = view.state.tr.replaceSelection(slice);
            view.dispatch(tr.scrollIntoView());

            event.preventDefault();
            return true;
          },
        },
      }),
    ];
  },
});

/**
 * Apply every declaration the generated stylesheet has for elements under
 * `root`, as inline `style`, without disturbing whatever inline style an
 * element already carries. See
 * `../rooster-editor/WordClipboardEnginePlugin.ts`'s `inlineGeneratedStyles`
 * for the full reasoning — identical mechanism, needed here for a related
 * but distinct reason: `wordList`/`wordListItem`'s `parseHTML` reads
 * `style` directly rather than resolving CSS classes.
 */
function inlineGeneratedStyles(root: Element): void {
  const sheet = installedStylesheet?.sheet;
  if (!sheet) return;

  for (const rule of Array.from(sheet.cssRules)) {
    if (!(rule instanceof CSSStyleRule)) continue;

    let matches: Element[];
    try {
      matches = root.matches(rule.selectorText)
        ? [root, ...Array.from(root.querySelectorAll(rule.selectorText))]
        : Array.from(root.querySelectorAll(rule.selectorText));
    } catch {
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

let installedStylesheet: HTMLStyleElement | null = null;

function ensureWordClipboardStylesheet(css: string): void {
  if (!installedStylesheet) {
    installedStylesheet = document.createElement('style');
    installedStylesheet.setAttribute('data-word-clipboard-engine', '');
    document.head.appendChild(installedStylesheet);
  }
  const existing = new Set(
    installedStylesheet.textContent!.split('\n').filter((line) => line.trim().length > 0),
  );
  const additions = css.split('\n').filter((line) => line.trim().length > 0 && !existing.has(line));
  if (additions.length > 0) {
    installedStylesheet.textContent += (installedStylesheet.textContent ? '\n' : '') + additions.join('\n');
  }
}
