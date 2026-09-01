import type { DiagnosticCollector } from '../diagnostics/DiagnosticCollector.js';
import { DiagnosticCode } from '../diagnostics/UnsupportedFeature.js';
import { childNodesOf, isElement, nodePath, tagNameOf } from './dom.js';

/**
 * Clipboard HTML is untrusted input.
 *
 * Anyone who can put HTML on the clipboard — a malicious page the user copied
 * from, a crafted email — can put it into this parser. Two defences apply:
 *
 *  1. **Nothing executes.** The payload is parsed with `DOMParser`, which does
 *     not run scripts, load resources or execute inline handlers, and the
 *     resulting tree is never inserted into the live document. Scripts and
 *     event handlers are stripped from the working tree anyway, because the
 *     tree is also what the debug UI inspects.
 *
 *  2. **The renderer emits from the model, not from the input.** The output
 *     HTML is built by serialising the canonical model with escaped text and a
 *     fixed attribute whitelist, so an attribute in the input has no route to
 *     the output unless the model has a field for it.
 */

/** URL schemes permitted on a hyperlink. */
const SAFE_LINK_SCHEMES = new Set(['http:', 'https:', 'mailto:', 'tel:', 'ftp:', 'ftps:', 'sms:']);
/** URL schemes permitted on an image source. */
const SAFE_IMAGE_SCHEMES = new Set(['http:', 'https:', 'data:', 'blob:']);
/** Image MIME types permitted in a `data:` URI. */
const SAFE_IMAGE_MIME = /^image\/(png|jpe?g|gif|webp|bmp|tiff?|x-emf|x-wmf|svg\+xml)$/i;

/** Elements removed outright: they execute, navigate, or load remote content. */
const FORBIDDEN_TAGS = new Set([
  'script', 'iframe', 'frame', 'frameset', 'embed', 'applet', 'base', 'link',
  'meta', 'noscript', 'template', 'portal',
]);

/** Elements dropped from the working tree once the stylesheet has been mined. */
const STRIPPED_TAGS = new Set(['style', 'title', 'head']);

/**
 * Neutralise executing and resource-loading elements in the raw text, before
 * any DOM implementation acts on them.
 *
 * `DOMParser` documents are inert in a browser — no scripts run, no
 * subresources load — but the engine also runs under Node DOM shims, and not
 * all of them honour that. Relying on the host's parser to be inert makes a
 * security property depend on somebody else's correctness.
 *
 * The elements are **deleted from the source text**, never renamed. Renaming
 * looks tidier and is badly wrong: `<link>` is a void element, and a custom
 * element by that name is not. Word puts `<link rel=File-List …>` in the head
 * of every clipboard payload, so a renamed `<link>` is hoisted into the body,
 * stays open, swallows the entire document as its children, and is then
 * removed along with all of it. The symptom is a paste that yields zero
 * paragraphs — silently, because the payload parsed "successfully".
 *
 * The scrub applies to the working copy only. `document.rawHtml` keeps the
 * original bytes.
 */

/** Elements whose content is not document content: drop tag and children. */
const RAW_TEXT_TAGS = ['script', 'noscript', 'noframes', 'applet'];

/** Void elements: no children to lose, so only the tag itself is dropped. */
const VOID_TAGS = ['link', 'base', 'embed', 'frame'];

/**
 * Container elements: the element goes, its children stay. Deleting a
 * container's subtree is how content disappears, so it is never done here —
 * `sanitizeTree` still walks whatever is left.
 *
 * `<object>` is deliberately absent. It does not load anything in an inert
 * document, and it is how Word represents an embedded OLE object — removing it
 * here would destroy the very thing the OLE diagnostic exists to report.
 */
const CONTAINER_TAGS = ['iframe', 'frameset'];

/**
 * The attribute part of a start tag, tolerating `>` inside quoted values.
 *
 * Written so the alternatives cannot match the same text two ways. The obvious
 * spelling — `(?:"[^"]*"|'[^']*'|[^>])*` — is ambiguous, because `[^>]` also
 * matches the characters inside a quoted value, and the engine then explores
 * exponentially many partitions of the same input. That turned a 1 ms parse
 * into 771 ms on a small fixture and hung the suite outright on a large one.
 */
const TAG_ATTRIBUTES = `[^>"']*(?:(?:"[^"]*"|'[^']*')[^>"']*)*`;

/** Match a start tag, tolerating `>` inside quoted attribute values. */
function startTagPattern(tag: string, flags = 'gi'): RegExp {
  return new RegExp(`<${tag}\\b${TAG_ATTRIBUTES}>`, flags);
}

export interface PreScrubResult {
  html: string;
  /** What was removed, by tag name, for diagnostics. */
  removed: Record<string, number>;
}

export function preScrubRawHtml(html: string): PreScrubResult {
  let out = html;
  const removed: Record<string, number> = {};
  const count = (tag: string, n: number): void => {
    if (n > 0) removed[tag] = (removed[tag] ?? 0) + n;
  };

  for (const tag of RAW_TEXT_TAGS) {
    const paired = new RegExp(`<${tag}\\b${TAG_ATTRIBUTES}>[\\s\\S]*?<\\/${tag}\\s*>`, 'gi');
    let n = 0;
    out = out.replace(paired, () => {
      n++;
      return '';
    });
    // An unclosed raw-text element runs to the end of the document — that is
    // how the HTML tokenizer reads it, so removing to the end is correct.
    const unclosed = new RegExp(`<${tag}\\b${TAG_ATTRIBUTES}>[\\s\\S]*$`, 'i');
    if (unclosed.test(out)) {
      out = out.replace(unclosed, '');
      n++;
    }
    count(tag, n);
  }

  for (const tag of VOID_TAGS) {
    let n = 0;
    out = out.replace(startTagPattern(tag), () => {
      n++;
      return '';
    });
    count(tag, n);
  }

  for (const tag of CONTAINER_TAGS) {
    let n = 0;
    out = out.replace(startTagPattern(tag), () => {
      n++;
      return '';
    });
    out = out.replace(new RegExp(`<\\/${tag}\\s*>`, 'gi'), '');
    count(tag, n);
  }

  // `<meta http-equiv=refresh>` is a navigation instruction, not metadata.
  // Matched as a whole tag and then filtered, rather than with a single
  // pattern that has to look for the attribute in among the others.
  let refresh = 0;
  out = out.replace(startTagPattern('meta'), (tag) => {
    if (!/http-equiv\s*=\s*["']?refresh/i.test(tag)) return tag;
    refresh++;
    return '';
  });
  count('meta', refresh);

  return { html: out, removed };
}

/** The diagnostic code a pre-scrub removal is reported under. */
export function preScrubDiagnosticCode(tag: string): string {
  return RAW_TEXT_TAGS.includes(tag)
    ? DiagnosticCode.SECURITY_SCRIPT_REMOVED
    : DiagnosticCode.SECURITY_EXTERNAL_RESOURCE;
}

/** Move an element's children into its place, then leave it empty. */
function unwrapChildren(element: Element): void {
  const parent = element.parentNode;
  if (!parent) return;
  while (element.firstChild) parent.insertBefore(element.firstChild, element);
}

export interface SanitizeResult {
  removedElements: number;
  removedAttributes: number;
}

/**
 * Strip executable and resource-loading constructs from a parsed tree.
 *
 * Operates on the working clone; the raw HTML string is untouched, so a
 * security researcher can still see exactly what arrived.
 */
export function sanitizeTree(root: Node, diagnostics: DiagnosticCollector): SanitizeResult {
  const result: SanitizeResult = { removedElements: 0, removedAttributes: 0 };
  const walk = (node: Node): void => {
    for (const child of childNodesOf(node)) {
      if (!isElement(child)) continue;
      const tag = tagNameOf(child);
      if (FORBIDDEN_TAGS.has(tag)) {
        // Anything reaching here slipped past the pre-scrub, which means the
        // host parser put it somewhere unexpected — so its children may well
        // be real document content that a mis-parse swept inside it. They are
        // unwrapped rather than deleted, except for the raw-text elements
        // whose children are code and not content. Deleting a subtree on the
        // strength of an element name is how a whole paste disappears.
        const isRawText = tag === 'script' || tag === 'noscript' || tag === 'applet';
        if (!isRawText) unwrapChildren(child);
        child.parentNode?.removeChild(child);
        result.removedElements++;
        diagnostics.warn(
          tag === 'script' ? DiagnosticCode.SECURITY_SCRIPT_REMOVED : DiagnosticCode.SECURITY_EXTERNAL_RESOURCE,
          `<${tag}> was removed from the clipboard payload. Clipboard HTML is never executed and never loads remote resources.`,
          { location: { tagName: tag, path: nodePath(child) }, fidelity: 'EQUIVALENT' },
        );
        continue;
      }
      if (STRIPPED_TAGS.has(tag)) {
        // The stylesheet has already been mined from the raw text; the element
        // itself has no place in the content tree.
        child.parentNode?.removeChild(child);
        result.removedElements++;
        continue;
      }
      result.removedAttributes += stripDangerousAttributes(child, diagnostics);
      walk(child);
    }
  };
  walk(root);
  return result;
}

function stripDangerousAttributes(element: Element, diagnostics: DiagnosticCollector): number {
  let removed = 0;
  for (const attribute of Array.from(element.attributes)) {
    const name = attribute.name.toLowerCase();
    if (name.startsWith('on')) {
      element.removeAttribute(attribute.name);
      removed++;
      diagnostics.warn(
        DiagnosticCode.SECURITY_EVENT_HANDLER_REMOVED,
        `Inline event handler "${name}" was removed from <${tagNameOf(element)}>.`,
        { location: { tagName: tagNameOf(element) }, fidelity: 'EQUIVALENT' },
      );
      continue;
    }
    if ((name === 'href' || name === 'src' || name === 'xlink:href' || name === 'action') &&
        isDangerousUrl(attribute.value)) {
      element.removeAttribute(attribute.name);
      removed++;
      diagnostics.warn(
        DiagnosticCode.SECURITY_URL_BLOCKED,
        `A "${schemeOf(attribute.value) ?? 'malformed'}" URL was blocked on <${tagNameOf(element)}>.`,
        { location: { tagName: tagNameOf(element) }, fidelity: 'EQUIVALENT' },
      );
    }
  }
  return removed;
}

function schemeOf(url: string): string | undefined {
  const match = /^\s*([a-z][a-z0-9+.-]*):/i.exec(stripControlCharacters(url));
  return match ? `${match[1]!.toLowerCase()}:` : undefined;
}

/** Control characters and entities are the classic scheme-smuggling trick. */
function stripControlCharacters(value: string): string {
  // eslint-disable-next-line no-control-regex
  return value.replace(/[\u0000-\u0020\u007f-\u009f]/g, '');
}

function isDangerousUrl(url: string): boolean {
  const scheme = schemeOf(url);
  if (!scheme) return false;
  return scheme === 'javascript:' || scheme === 'vbscript:' || scheme === 'jscript:';
}

export interface UrlCheck {
  /** The URL to use, or an empty string when it must not be used. */
  safe: string;
  blocked: boolean;
  reason?: string;
}

/**
 * Screen a hyperlink target.
 *
 * `file:///` targets are common in Word documents and are not an injection
 * risk, but browsers refuse to navigate to them from an http page, so they are
 * reported rather than emitted as a link that silently does nothing.
 */
export function checkLinkUrl(url: string): UrlCheck {
  const trimmed = (url ?? '').trim();
  if (!trimmed) return { safe: '', blocked: false };
  if (trimmed.startsWith('#')) return { safe: trimmed, blocked: false };

  const scheme = schemeOf(trimmed);
  if (!scheme) {
    // Relative URL. Harmless, and Word emits them for same-document targets.
    return { safe: trimmed, blocked: false };
  }
  if (SAFE_LINK_SCHEMES.has(scheme)) return { safe: trimmed, blocked: false };
  if (scheme === 'file:') {
    return {
      safe: '',
      blocked: true,
      reason: 'file:// links point at the author\'s local disk and cannot be followed from a web page.',
    };
  }
  return { safe: '', blocked: true, reason: `The "${scheme}" URL scheme is not permitted.` };
}

/** Screen an image source. */
export function checkImageUrl(url: string): UrlCheck {
  const trimmed = (url ?? '').trim();
  if (!trimmed) return { safe: '', blocked: false };

  const scheme = schemeOf(trimmed);
  if (!scheme) return { safe: trimmed, blocked: false };
  if (scheme === 'data:') {
    const mime = /^data:([^;,]+)/i.exec(trimmed)?.[1] ?? '';
    if (SAFE_IMAGE_MIME.test(mime.trim())) return { safe: trimmed, blocked: false };
    return { safe: '', blocked: true, reason: `data: URI of type "${mime}" is not a permitted image type.` };
  }
  if (SAFE_IMAGE_SCHEMES.has(scheme)) return { safe: trimmed, blocked: false };
  if (scheme === 'file:' || scheme === 'cid:') {
    // Not dangerous, just unusable — the image parser reports it properly.
    return { safe: '', blocked: false };
  }
  return { safe: '', blocked: true, reason: `The "${scheme}" URL scheme is not permitted for images.` };
}

/** Limits that keep a hostile or enormous paste bounded. */
export interface ParseLimits {
  /** Maximum characters of raw HTML accepted. */
  maxHtmlLength: number;
  /** Maximum DOM nodes traversed. */
  maxNodes: number;
  /** Maximum block nesting depth (containers, tables, lists). */
  maxDepth: number;
  /** Maximum number of top-level and nested blocks produced. */
  maxBlocks: number;
  /** Maximum table nesting depth. */
  maxTableDepth: number;
}

export const DEFAULT_LIMITS: ParseLimits = {
  maxHtmlLength: 32 * 1024 * 1024,
  maxNodes: 500_000,
  maxDepth: 64,
  maxBlocks: 200_000,
  maxTableDepth: 16,
};
