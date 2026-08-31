/**
 * DOM helpers.
 *
 * The engine runs in a browser and under Node (for tests and server-side
 * conversion), so it never reaches for a global `document`. A `DOMParser` is
 * resolved once, from the global scope, and everything else works off the
 * parsed document that comes back.
 */

export const NODE_ELEMENT = 1;
export const NODE_TEXT = 3;
export const NODE_COMMENT = 8;

/** Locate a usable DOMParser, or explain precisely what is missing. */
export function getDomParser(): DOMParser {
  const globalScope = globalThis as unknown as { DOMParser?: new () => DOMParser };
  if (typeof globalScope.DOMParser === 'function') {
    return new globalScope.DOMParser();
  }
  throw new Error(
    'No DOMParser available. In a browser this is built in; under Node run the engine in a ' +
      'DOM environment (vitest `environment: "happy-dom"`, jsdom, or linkedom) before calling parseWordClipboard().',
  );
}

/**
 * Parse an HTML string into a document.
 *
 * `text/html` parsing is deliberate: it is what a browser does with clipboard
 * HTML, so Word's bogus conditional comments and unclosed tags land in exactly
 * the shape the engine's rules were written against. XML parsing would reject
 * the payload outright.
 */
export function parseHtmlDocument(html: string): Document {
  return getDomParser().parseFromString(html, 'text/html');
}

export function isElement(node: Node): node is Element {
  return node.nodeType === NODE_ELEMENT;
}

export function isTextNode(node: Node): node is Text {
  return node.nodeType === NODE_TEXT;
}

export function isCommentNode(node: Node): node is Comment {
  return node.nodeType === NODE_COMMENT;
}

/** Lower-cased tag name, with the namespace prefix intact (`o:p`, `v:shape`). */
export function tagNameOf(element: Element): string {
  return element.tagName.toLowerCase();
}

/** Case-insensitive attribute read that tolerates missing attributes. */
export function attr(element: Element, name: string): string | undefined {
  const value = element.getAttribute(name);
  return value === null ? undefined : value;
}

/** Class list as a lower-cased array. */
export function classList(element: Element): string[] {
  const value = element.getAttribute('class');
  if (!value) return [];
  return value.split(/\s+/).filter(Boolean);
}

export function hasClass(element: Element, name: string): boolean {
  const lower = name.toLowerCase();
  return classList(element).some((c) => c.toLowerCase() === lower);
}

/** Children as a stable array (the live NodeList changes under mutation). */
export function childNodesOf(node: Node): Node[] {
  return Array.from(node.childNodes);
}

/** Build a short CSS-ish path to a node, for diagnostics. */
export function nodePath(node: Node, maxDepth = 6): string {
  const parts: string[] = [];
  let current: Node | null = node;
  let depth = 0;
  while (current && depth < maxDepth) {
    if (isElement(current)) {
      const tag = tagNameOf(current);
      const classes = classList(current);
      parts.unshift(classes.length ? `${tag}.${classes[0]}` : tag);
    } else if (isTextNode(current)) {
      parts.unshift('#text');
    } else if (isCommentNode(current)) {
      parts.unshift('#comment');
    }
    current = current.parentNode;
    depth++;
  }
  return parts.join(' > ');
}

/** A truncated excerpt of a node's markup, safe to put in a diagnostic. */
export function excerpt(node: Node, limit = 200): string {
  let text: string;
  if (isElement(node)) text = node.outerHTML;
  else if (isCommentNode(node)) text = `<!--${node.data}-->`;
  else text = node.textContent ?? '';
  text = text.replace(/\s+/g, ' ').trim();
  return text.length > limit ? `${text.slice(0, limit)}…` : text;
}

/** Count every node in a subtree, used to enforce the parse budget. */
export function countNodes(node: Node, limit = Number.MAX_SAFE_INTEGER): number {
  let count = 0;
  const stack: Node[] = [node];
  while (stack.length > 0) {
    const current = stack.pop()!;
    count++;
    if (count >= limit) return count;
    const children = current.childNodes;
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      if (child) stack.push(child);
    }
  }
  return count;
}

/** Elements that introduce a block box in Word's output. */
const BLOCK_LEVEL_TAGS = new Set([
  'address', 'article', 'aside', 'blockquote', 'center', 'dd', 'div', 'dl', 'dt',
  'fieldset', 'figcaption', 'figure', 'footer', 'form', 'h1', 'h2', 'h3', 'h4',
  'h5', 'h6', 'header', 'hr', 'li', 'main', 'nav', 'ol', 'p', 'pre', 'section',
  'table', 'tbody', 'td', 'tfoot', 'th', 'thead', 'tr', 'ul',
]);

export function isBlockLevelTag(tagName: string): boolean {
  return BLOCK_LEVEL_TAGS.has(tagName);
}

/** True when a subtree contains any block-level element. */
export function containsBlockLevel(element: Element): boolean {
  const children = element.childNodes;
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    if (!child || !isElement(child)) continue;
    if (isBlockLevelTag(tagNameOf(child))) return true;
    if (containsBlockLevel(child)) return true;
  }
  return false;
}

/**
 * Collapse HTML whitespace the way a browser would, while leaving
 * non-breaking spaces alone — Word uses U+00A0 as real content (it is how an
 * empty paragraph keeps its height and how `mso-spacerun` indents text).
 */
export function collapseWhitespace(text: string): string {
  return text.replace(/[\t\n\r\f ]+/g, ' ');
}

/** True when a string is whitespace only, ignoring non-breaking spaces. */
export function isCollapsibleWhitespace(text: string): boolean {
  return /^[\t\n\r\f ]*$/.test(text);
}

/** Escape text for safe insertion into HTML output. */
export function escapeHtmlText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/** Escape a value for use inside a double-quoted HTML attribute. */
export function escapeHtmlAttribute(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
