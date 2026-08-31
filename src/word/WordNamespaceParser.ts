import type { WordDocumentMetadata } from '../model/WordMetadata.js';

/**
 * Office XML namespaces.
 *
 * Word declares its namespaces on `<html>` and then uses prefixed elements
 * throughout: `<o:p>` (the paragraph mark), `<w:WordDocument>` (settings),
 * `<v:shape>` (drawings), `<m:oMath>` (equations), `<st1:place>` (smart tags).
 *
 * None of these belong in browser HTML, but each one carries information —
 * so they are read here first, and removed later.
 */

export const OFFICE_NAMESPACES: Record<string, string> = {
  o: 'urn:schemas-microsoft-com:office:office',
  w: 'urn:schemas-microsoft-com:office:word',
  x: 'urn:schemas-microsoft-com:office:excel',
  p: 'urn:schemas-microsoft-com:office:powerpoint',
  v: 'urn:schemas-microsoft-com:vml',
  m: 'http://schemas.microsoft.com/office/2004/12/omml',
  st1: 'urn:schemas-microsoft-com:office:smarttags',
};

/** Namespace prefixes whose elements are structural noise once mined. */
export const KNOWN_OFFICE_PREFIXES = new Set(['o', 'w', 'x', 'p', 'v', 'm', 'st1', 'st2', 'office']);

/** Split `o:p` into `{ prefix: 'o', local: 'p' }`. */
export function splitQualifiedName(tagName: string): { prefix?: string; local: string } {
  const name = tagName.toLowerCase();
  const colon = name.indexOf(':');
  if (colon === -1) return { local: name };
  return { prefix: name.slice(0, colon), local: name.slice(colon + 1) };
}

/** True when the element belongs to an Office namespace. */
export function isOfficeNamespacedElement(tagName: string): boolean {
  const { prefix } = splitQualifiedName(tagName);
  return prefix !== undefined && KNOWN_OFFICE_PREFIXES.has(prefix);
}

/**
 * `<o:p>` is Word's paragraph-mark element. It is empty in a normal paragraph
 * and contains `&nbsp;` in an empty one, which is the only way to tell an
 * intentionally blank paragraph from a structural artefact.
 */
export function isParagraphMark(tagName: string): boolean {
  return tagName.toLowerCase() === 'o:p';
}

/** Read the namespace declarations off the `<html>` element in raw markup. */
export function parseNamespaceDeclarations(html: string): Record<string, string> {
  const namespaces: Record<string, string> = {};
  const htmlTag = /<html\b([^>]*)>/i.exec(html);
  const attributes = htmlTag?.[1] ?? '';
  const pattern = /xmlns:([a-zA-Z0-9_-]+)\s*=\s*["']([^"']*)["']/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(attributes)) !== null) {
    namespaces[match[1]!.toLowerCase()] = match[2]!;
  }
  const defaultNs = /xmlns\s*=\s*["']([^"']*)["']/.exec(attributes);
  if (defaultNs) namespaces[''] = defaultNs[1]!;
  return namespaces;
}

/**
 * Read `<o:DocumentProperties>` and `<w:WordDocument>` out of the Office
 * conditional comment blocks.
 *
 * These are XML fragments living inside an HTML comment, so they are parsed
 * with a small tag scanner rather than the DOM. The values are useful for
 * telling apart Word builds and for reporting what the source document was.
 */
export function parseOfficeMetadata(
  conditionalContents: string[],
): Pick<WordDocumentMetadata, 'documentProperties' | 'wordSettings'> {
  const documentProperties: Record<string, string> = {};
  const wordSettings: Record<string, string> = {};

  for (const content of conditionalContents) {
    collectSimpleElements(content, /<o:([A-Za-z0-9_]+)>([\s\S]*?)<\/o:\1>/g, documentProperties);
    collectSimpleElements(content, /<w:([A-Za-z0-9_]+)>([\s\S]*?)<\/w:\1>/g, wordSettings);
    // Self-closing settings such as `<w:TrackMoves/>` are flags.
    const flags = /<w:([A-Za-z0-9_]+)\s*\/>/g;
    let match: RegExpExecArray | null;
    while ((match = flags.exec(content)) !== null) {
      const key = match[1]!;
      if (!(key in wordSettings)) wordSettings[key] = 'true';
    }
  }
  return { documentProperties, wordSettings };
}

function collectSimpleElements(
  content: string,
  pattern: RegExp,
  target: Record<string, string>,
): void {
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(content)) !== null) {
    const key = match[1]!;
    const value = (match[2] ?? '').trim();
    // Only keep leaf values; a nested block is structure, not a property.
    if (value.includes('<')) continue;
    if (value.length > 0 && value.length < 512) target[key] = decodeBasicEntities(value);
  }
}

function decodeBasicEntities(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&');
}

/** Read `<meta>` values Word writes into the head. */
export function parseMetaTags(html: string): {
  generator?: string;
  progId?: string;
  charset?: string;
} {
  const result: { generator?: string; progId?: string; charset?: string } = {};

  const generator = /<meta[^>]+name=["']?Generator["']?[^>]*content=["']?([^"'>]+)/i.exec(html);
  if (generator) result.generator = generator[1]!.trim();

  const progId = /<meta[^>]+name=["']?ProgId["']?[^>]*content=["']?([^"'>]+)/i.exec(html);
  if (progId) result.progId = progId[1]!.trim();

  const charsetAttr = /<meta[^>]+charset=["']?([\w-]+)/i.exec(html);
  if (charsetAttr) {
    result.charset = charsetAttr[1]!.trim();
  } else {
    const contentType = /<meta[^>]+http-equiv=["']?Content-Type["']?[^>]*content=["']?[^"'>]*charset=([\w-]+)/i.exec(
      html,
    );
    if (contentType) result.charset = contentType[1]!.trim();
  }
  return result;
}

/** Collect `class="WordSectionN"` names present in the payload. */
export function parseSectionNames(html: string): string[] {
  const sections = new Set<string>();
  const pattern = /class=["']?(WordSection\d+)/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(html)) !== null) {
    sections.add(match[1]!);
  }
  return [...sections];
}
