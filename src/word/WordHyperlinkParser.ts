import type { WordBookmark, WordHyperlink } from '../model/Hyperlink.js';
import { attr, tagNameOf } from '../util/dom.js';
import { checkLinkUrl } from '../util/security.js';
import { DiagnosticCode } from '../diagnostics/UnsupportedFeature.js';
import { nextBookmarkId, nextHyperlinkId, type WordParseContext } from './WordParseContext.js';

/**
 * Hyperlinks and bookmarks.
 *
 * Word writes a link as `<a href="…">` with the *run formatting inside it*,
 * which is why links are stored out of line in the document and referenced by
 * id: flattening a link into a single styled run would lose the bold word in
 * the middle of it.
 *
 * Word also writes bookmarks as `<a name="_Toc12345">` — the same element,
 * with no href. Those are anchors, not links, and Word's own TOC bookmarks
 * (`_Toc…`, `_Ref…`, `_GoBack`, `_Hlk…`) are marked internal so a consumer can
 * choose to drop them without losing user-authored bookmarks.
 */

const INTERNAL_BOOKMARK = /^_(Toc|Ref|Hlk|GoBack|Hlt|MailOriginal|MailAutoSig|top)/i;

/** Parse an `<a>` element into a hyperlink, registering it on the context. */
export function parseHyperlink(element: Element, ctx: WordParseContext): WordHyperlink | null {
  const rawHref = attr(element, 'href');
  if (rawHref === undefined) return null;

  const check = checkLinkUrl(rawHref);
  const id = nextHyperlinkId(ctx);
  const link: WordHyperlink = {
    id,
    href: check.safe,
    rawHref,
  };

  if (check.blocked) {
    link.blocked = true;
    ctx.diagnostics.warn(
      DiagnosticCode.SECURITY_URL_BLOCKED,
      `Hyperlink target was not emitted: ${check.reason ?? 'unsupported scheme'} The original target is preserved on the model as rawHref.`,
      { details: { rawHref: rawHref.slice(0, 200) }, fidelity: 'APPROXIMATED' },
    );
  }

  if (rawHref.startsWith('#')) link.anchor = rawHref.slice(1);
  const title = attr(element, 'title');
  if (title) link.title = title;
  const target = attr(element, 'target');
  if (target) link.target = target;

  ctx.hyperlinks[id] = link;
  return link;
}

/** Parse an `<a name>` / `id` anchor into a bookmark. */
export function parseBookmark(element: Element, ctx: WordParseContext): WordBookmark | null {
  if (tagNameOf(element) !== 'a') return null;
  const name = attr(element, 'name') ?? (attr(element, 'href') === undefined ? attr(element, 'id') : undefined);
  if (!name) return null;

  const id = nextBookmarkId(ctx);
  const bookmark: WordBookmark = {
    id,
    name,
    internal: INTERNAL_BOOKMARK.test(name),
  };
  ctx.bookmarks[id] = bookmark;
  return bookmark;
}

/** True when the element is an anchor with no href (a bookmark target). */
export function isBookmarkAnchor(element: Element): boolean {
  return (
    tagNameOf(element) === 'a' &&
    attr(element, 'href') === undefined &&
    (attr(element, 'name') !== undefined || attr(element, 'id') !== undefined)
  );
}
