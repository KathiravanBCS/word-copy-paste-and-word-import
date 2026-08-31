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
 * Tags neutralised in the raw text *before* the DOM ever sees them.
 *
 * `DOMParser` documents are inert in a browser — no scripts run, no
 * subresources load — but the engine also runs under Node DOM shims, and not
 * all of them honour that. Relying on the host's parser to be inert is a
 * dependency on somebody else's correctness for a security property, so these
 * elements are renamed to an inert custom element in the source text first.
 *
 * The rename is applied to the working copy only. `document.rawHtml` keeps the
 * original bytes, so the payload can still be inspected exactly as it arrived.
 */
const NEUTRALISED_TAGS = ['script', 'iframe', 'frame', 'frameset', 'embed', 'applet', 'link', 'base'];

export const NEUTRALISED_PREFIX = 'wce-blocked-';

export function preScrubRawHtml(html: string): string {
  let out = html;
  for (const tag of NEUTRALISED_TAGS) {
    const open = new RegExp(`<${tag}(?=[\\s/>])`, 'gi');
    const close = new RegExp(`</${tag}(?=[\\s>])`, 'gi');
    out = out.replace(open, `<${NEUTRALISED_PREFIX}${tag}`);
    out = out.replace(close, `</${NEUTRALISED_PREFIX}${tag}`);
    // The tag with nothing after the name, e.g. `<script>` / `</script>`.
    out = out.replace(new RegExp(`<${tag}>`, 'gi'), `<${NEUTRALISED_PREFIX}${tag}>`);
    out = out.replace(new RegExp(`</${tag}>`, 'gi'), `</${NEUTRALISED_PREFIX}${tag}>`);
  }
  // `<meta http-equiv=refresh>` is a navigation instruction, not metadata.
  out = out.replace(/<meta\b([^>]*http-equiv\s*=\s*["']?refresh)/gi, `<${NEUTRALISED_PREFIX}meta$1`);
  return out;
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
      if (tag.startsWith(NEUTRALISED_PREFIX)) {
        const original = tag.slice(NEUTRALISED_PREFIX.length);
        child.parentNode?.removeChild(child);
        result.removedElements++;
        diagnostics.warn(
          original === 'script'
            ? DiagnosticCode.SECURITY_SCRIPT_REMOVED
            : DiagnosticCode.SECURITY_EXTERNAL_RESOURCE,
          `<${original}> was neutralised before parsing and removed. Clipboard HTML is never executed and never loads remote resources.`,
          { location: { tagName: original }, fidelity: 'EQUIVALENT' },
        );
        continue;
      }
      if (FORBIDDEN_TAGS.has(tag)) {
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
