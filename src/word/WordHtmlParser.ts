import type { WordDocument } from '../model/Document.js';
import type { WordBlock, WordContainer, WordUnsupportedObject } from '../model/Block.js';
import type { ParagraphFormatting, RunFormatting } from '../model/Style.js';
import type { WordDocumentMetadata } from '../model/WordMetadata.js';
import type { ClipboardPayload } from '../clipboard/ClipboardPayload.js';
import { detectWordHtml } from '../detection/WordDetector.js';
import { DiagnosticCollector } from '../diagnostics/DiagnosticCollector.js';
import { DiagnosticCode } from '../diagnostics/UnsupportedFeature.js';
import {
  attr,
  childNodesOf,
  classList,
  excerpt,
  hasClass,
  isBlockLevelTag,
  isCollapsibleWhitespace,
  isCommentNode,
  isElement,
  isTextNode,
  nodePath,
  parseHtmlDocument,
  tagNameOf,
} from '../util/dom.js';
import { preScrubRawHtml, sanitizeTree } from '../util/security.js';
import { parseInlineStyle } from './WordCssTokenizer.js';
import {
  parseParagraphFormattingFromCss,
  parseRunFormattingFromCss,
} from './WordFormattingParser.js';
import { parseWordStyleSheet } from './WordStyleParser.js';
import {
  extractStyleBlocks,
  extractHiddenConditionalBlocks,
} from './WordConditionalCommentParser.js';
import {
  parseMetaTags,
  parseNamespaceDeclarations,
  parseOfficeMetadata,
  parseSectionNames,
} from './WordNamespaceParser.js';
import {
  collectClaimedShapeIds,
  collectVmlShapes,
  describeUnsupportedObject,
  isVmlSubPart,
  parseVmlShapeImage,
} from './WordImageParser.js';
import { parseParagraph, resolveElementCss } from './WordParagraphParser.js';
import { parseTable } from './WordTableParser.js';
import { extractComments, extractNotes } from './WordCommentParser.js';
import { buildListItem, extractRenderedMarker, findListReference } from './WordListParser.js';
import { parseListReference } from './WordListStyleParser.js';
import {
  chargeBlock,
  chargeNode,
  createParseContext,
  type ParseOptions,
  type WordParseContext,
} from './WordParseContext.js';

/**
 * The orchestrator.
 *
 * The pipeline is fixed and the order is the point of the whole design:
 *
 *     raw HTML (never mutated)
 *       -> stylesheet mined from raw text      (@list, @font-face, mso-*)
 *       -> metadata mined from raw text        (namespaces, office XML, meta)
 *       -> VML shapes mined from raw text      (inside conditional comments)
 *       -> DOM parsed from raw text            (a working clone)
 *       -> fragment boundary honoured
 *       -> security scrub of the working clone
 *       -> comments and notes lifted out
 *       -> content walked into the model       (markers lifted before runs)
 *       -> Word-only representation discarded
 *
 * Cleanup is last. Every earlier stage is extraction, so nothing is thrown
 * away before something has had the chance to understand it.
 */

/** Result of parsing, before normalization. */
export interface WordHtmlParseResult {
  document: WordDocument;
  context: WordParseContext;
}

export function parseWordHtml(
  rawHtml: string,
  options: ParseOptions = {},
  payload?: ClipboardPayload,
): WordHtmlParseResult {
  const diagnostics = new DiagnosticCollector(options.maxDiagnostics);

  const detection = detectWordHtml(rawHtml, {
    ...(options.detectionThreshold !== undefined ? { threshold: options.detectionThreshold } : {}),
  });
  if (!detection.isWord && !options.forceWord) {
    diagnostics.info(
      DiagnosticCode.NOT_WORD_CONTENT,
      `The payload did not look like Word content (confidence ${detection.confidence.toFixed(2)}). It was parsed with the generic HTML path; Word-specific rules were still applied where signals were present.`,
      { details: { confidence: detection.confidence, signals: detection.signals.length }, fidelity: 'EQUIVALENT' },
    );
  }

  // --- 1. mine the stylesheet from raw text ------------------------------
  const styleBlocks = extractStyleBlocks(rawHtml);
  const sheet = parseWordStyleSheet(styleBlocks, diagnostics);

  const ctx = createParseContext(diagnostics, options, sheet, payload?.images ?? []);

  // --- 2. mine metadata ---------------------------------------------------
  const conditional = extractHiddenConditionalBlocks(rawHtml);
  const officeMetadata = parseOfficeMetadata(
    conditional.filter((b) => b.payload === 'office-metadata' || b.payload === 'word-settings').map((b) => b.content),
  );
  const meta = parseMetaTags(rawHtml);
  const metadata: WordDocumentMetadata = {
    namespaces: parseNamespaceDeclarations(rawHtml),
    documentProperties: officeMetadata.documentProperties,
    wordSettings: officeMetadata.wordSettings,
    sections: parseSectionNames(rawHtml),
    fragmentBoundaryFound: false,
    rawHtmlLength: rawHtml.length,
    sourceApplication: detection.source,
  };
  if (meta.generator) metadata.generator = meta.generator;
  if (meta.progId) metadata.progId = meta.progId;
  if (meta.charset) metadata.charset = meta.charset;
  if (detection.wordVersion !== undefined) metadata.wordVersion = detection.wordVersion;

  // --- 3. mine VML from raw text (it lives inside comments) ---------------
  collectVmlShapes(rawHtml, ctx);
  collectClaimedShapeIds(rawHtml, ctx);

  // --- 4. parse the DOM ---------------------------------------------------
  if (rawHtml.length > ctx.limits.maxHtmlLength) {
    diagnostics.error(
      DiagnosticCode.LIMIT_DOCUMENT_TRUNCATED,
      `The clipboard payload is ${rawHtml.length} characters, beyond the ${ctx.limits.maxHtmlLength} limit. It was truncated before parsing.`,
    );
  }
  const truncated =
    rawHtml.length > ctx.limits.maxHtmlLength ? rawHtml.slice(0, ctx.limits.maxHtmlLength) : rawHtml;
  // Neutralise executing/loading elements in the text before any DOM
  // implementation gets to act on them. `rawHtml` on the document is untouched.
  const workingHtml = preScrubRawHtml(truncated);
  const dom = parseHtmlDocument(workingHtml);
  const body = dom.body ?? dom.documentElement;

  // --- 5. honour the CF_HTML fragment boundary ---------------------------
  metadata.fragmentBoundaryFound = pruneToFragment(body, diagnostics);

  // --- 6. security scrub of the working clone ----------------------------
  sanitizeTree(body, diagnostics);

  // --- 7. lift comments and notes out of the flow ------------------------
  extractComments(body, ctx);
  extractNotes(body, ctx);

  // --- 8. walk the content into the model --------------------------------
  const blocks = parseBlockChildren(body, ctx, { depth: 0 });

  // --- 9. account for anything left over ---------------------------------
  appendOrphanShapes(blocks, ctx);

  const document: WordDocument = {
    blocks,
    styles: sheet,
    lists: Object.values(sheet.lists),
    images: ctx.images,
    hyperlinks: ctx.hyperlinks,
    bookmarks: ctx.bookmarks,
    diagnostics: diagnostics.all(),
    metadata,
    detection,
    rawHtml,
  };
  if (payload?.text !== undefined) document.rawText = payload.text;

  return { document, context: ctx };
}

/* -------------------------------------------------------------------------
 * Fragment boundary
 * ---------------------------------------------------------------------- */

/**
 * Trim the tree to the CF_HTML fragment.
 *
 * The clipboard's `text/html` flavour is a whole document, but only the part
 * between `<!--StartFragment-->` and `<!--EndFragment-->` is what the user
 * selected. Everything else is the surrounding page or Word's own scaffolding.
 *
 * The trim is done on the DOM, by removing siblings outward from each marker,
 * rather than by slicing the HTML string. Slicing breaks the moment the
 * selection starts inside a table: `<tr>…</tr>` reparsed on its own loses its
 * rows entirely.
 */
function pruneToFragment(body: Element, diagnostics: DiagnosticCollector): boolean {
  const start = findComment(body, /^\s*StartFragment\s*$/i);
  const end = findComment(body, /^\s*EndFragment\s*$/i);
  if (!start || !end) {
    if (start || end) {
      diagnostics.info(
        DiagnosticCode.WORD_FRAGMENT_BOUNDARY_MISSING,
        'Only one of the StartFragment/EndFragment markers was present; the whole body was used as the content.',
        { fidelity: 'EQUIVALENT' },
      );
    }
    return false;
  }

  let node: Node | null = start;
  while (node && node !== body) {
    const parent: Node | null = node.parentNode;
    if (!parent) break;
    while (node.previousSibling) parent.removeChild(node.previousSibling);
    node = parent;
  }

  node = end;
  while (node && node !== body) {
    const parent: Node | null = node.parentNode;
    if (!parent) break;
    while (node.nextSibling) parent.removeChild(node.nextSibling);
    node = parent;
  }

  start.parentNode?.removeChild(start);
  end.parentNode?.removeChild(end);
  return true;
}

function findComment(root: Node, pattern: RegExp): Comment | null {
  const stack: Node[] = [root];
  while (stack.length > 0) {
    const node = stack.shift()!;
    if (isCommentNode(node) && pattern.test(node.data ?? '')) return node;
    stack.push(...childNodesOf(node));
  }
  return null;
}

/* -------------------------------------------------------------------------
 * Block walking
 * ---------------------------------------------------------------------- */

export interface BlockParseOptions {
  depth: number;
  inheritedRun?: RunFormatting;
  inheritedParagraph?: ParagraphFormatting;
  /** Section name in effect, used to detect section breaks. */
  sectionState?: { current?: string };
}

/**
 * Walk an element's children into blocks.
 *
 * Inline content that appears directly between block elements is gathered into
 * an implicit paragraph rather than dropped, which is what keeps stray text in
 * a `<div>` or a table cell from disappearing.
 */
export function parseBlockChildren(
  parent: Element,
  ctx: WordParseContext,
  options: BlockParseOptions,
): WordBlock[] {
  const blocks: WordBlock[] = [];
  const pending: Node[] = [];
  const sectionState = options.sectionState ?? {};

  const flush = (): void => {
    if (pending.length === 0) return;
    const nodes = pending.splice(0, pending.length);
    if (nodes.every((n) => isTextNode(n) && isCollapsibleWhitespace(n.data ?? ''))) return;
    const wrapper = parent.ownerDocument.createElement('p');
    for (const node of nodes) wrapper.appendChild(node.cloneNode(true));
    const paragraph = parseParagraph(wrapper, ctx, buildParagraphOptions(options));
    if (!paragraph.empty || paragraph.runs.length > 0) blocks.push(paragraph);
  };

  for (const child of childNodesOf(parent)) {
    if (!chargeNode(ctx)) break;

    if (isTextNode(child)) {
      pending.push(child);
      continue;
    }
    if (isCommentNode(child)) {
      continue;
    }
    if (!isElement(child)) continue;

    const tag = tagNameOf(child);
    if (!isBlockLevelTag(tag) && !isBlockLikeSpecial(child, tag)) {
      pending.push(child);
      continue;
    }

    flush();
    const produced = parseBlockElement(child, ctx, { ...options, sectionState });
    for (const block of produced) {
      if (!chargeBlock(ctx)) return blocks;
      blocks.push(block);
    }
  }
  flush();
  return blocks;
}

function buildParagraphOptions(options: BlockParseOptions): {
  inheritedRun?: RunFormatting;
  inheritedParagraph?: ParagraphFormatting;
} {
  const result: { inheritedRun?: RunFormatting; inheritedParagraph?: ParagraphFormatting } = {};
  if (options.inheritedRun) result.inheritedRun = options.inheritedRun;
  if (options.inheritedParagraph) result.inheritedParagraph = options.inheritedParagraph;
  return result;
}

/** Elements that are not block-level in HTML but behave as blocks in Word. */
function isBlockLikeSpecial(element: Element, tag: string): boolean {
  if (tag === 'object' || tag === 'o:oleobject') return true;
  if (tag.startsWith('v:') && tag !== 'v:imagedata') return true;
  if (tag === 'br') {
    const style = parseInlineStyle(attr(element, 'style'));
    return /always|left|right/i.test(style['page-break-before'] ?? '');
  }
  return false;
}

function parseBlockElement(
  element: Element,
  ctx: WordParseContext,
  options: BlockParseOptions,
): WordBlock[] {
  const tag = tagNameOf(element);

  switch (tag) {
    case 'p':
    case 'h1':
    case 'h2':
    case 'h3':
    case 'h4':
    case 'h5':
    case 'h6':
    case 'dt':
    case 'dd':
      return [parseParagraph(element, ctx, buildParagraphOptions(options))];

    case 'div':
    case 'section':
    case 'article':
    case 'main':
    case 'header':
    case 'footer':
    case 'aside':
    case 'nav':
    case 'form':
    case 'fieldset':
    case 'figure':
    case 'center':
      return parseDivLike(element, ctx, options);

    case 'blockquote':
      return [parseContainer(element, ctx, options, 'blockquote')];

    case 'ol':
    case 'ul':
    case 'dl':
      return parseHtmlList(element, ctx, options, 0);

    case 'table':
      return [
        parseTable(element, ctx, {
          depth: options.depth,
          ...(options.inheritedRun ? { inheritedRun: options.inheritedRun } : {}),
          ...(options.inheritedParagraph ? { inheritedParagraph: options.inheritedParagraph } : {}),
          parseBlocks: (el, c, o) => parseBlockChildren(el, c, o),
        }),
      ];

    case 'hr':
      return [parseHorizontalRule(element, ctx)];

    case 'br':
      return [{ type: 'page-break', breakType: 'page' }];

    case 'pre':
      return [parseParagraph(element, ctx, buildParagraphOptions(options))];

    case 'object':
    case 'o:oleobject':
      return [buildUnsupported(element, ctx, 'ole-object', DiagnosticCode.WORD_OLE_OBJECT)];

    case 'li':
      // A stray `<li>` outside a list container.
      return [parseParagraph(element, ctx, buildParagraphOptions(options))];

    case 'tbody':
    case 'thead':
    case 'tfoot':
    case 'tr':
    case 'td':
    case 'th':
      // Table parts reached outside a table (a partial copy). Keep the content.
      return parseBlockChildren(element, ctx, options);

    default:
      break;
  }

  if (tag.startsWith('v:')) return parseVmlBlock(element, ctx);

  return parseBlockChildren(element, ctx, options);
}

/**
 * `<div>` handling, which is where Word's section structure lives.
 *
 * `div.WordSection1`, `div.WordSection2` … are page-setup sections. A change
 * of section is a section break, which HTML cannot express as page setup, so
 * it becomes a page break plus a diagnostic saying what was lost.
 */
function parseDivLike(
  element: Element,
  ctx: WordParseContext,
  options: BlockParseOptions,
): WordBlock[] {
  const blocks: WordBlock[] = [];
  const sectionState = options.sectionState ?? {};
  const sectionClass = classList(element).find((c) => /^WordSection\d+$/i.test(c));

  if (sectionClass) {
    if (sectionState.current && sectionState.current !== sectionClass) {
      blocks.push({ type: 'page-break', breakType: 'section', sectionName: sectionClass });
      ctx.diagnostics.warn(
        DiagnosticCode.WORD_SECTION_BREAK_APPROXIMATED,
        `A Word section break (${sectionState.current} to ${sectionClass}) was represented as a page break. Per-section page setup — size, margins, headers, columns — has no HTML equivalent.`,
        { details: { from: sectionState.current, to: sectionClass } },
      );
    }
    sectionState.current = sectionClass;

    if (!ctx.options.keepSectionContainers) {
      blocks.push(...parseBlockChildren(element, ctx, { ...options, sectionState }));
      return blocks;
    }
    return [...blocks, parseContainer(element, ctx, { ...options, sectionState }, 'section')];
  }

  const { css } = resolveElementCss(element, ctx);
  const msoElement = (css['mso-element'] ?? '').toLowerCase();

  if (msoElement === 'frame' || msoElement.startsWith('para-border-div') === false && isTextBox(css)) {
    ctx.diagnostics.warn(
      DiagnosticCode.WORD_TEXT_BOX_APPROXIMATED,
      'A Word text box or frame was flattened into a bordered container. Its floating position relative to the page is not represented.',
      { location: { tagName: tagNameOf(element) } },
    );
    return [parseContainer(element, ctx, options, 'text-box')];
  }

  // Word wraps a bordered/shaded paragraph in a div carrying the border. That
  // div is meaningful; a bare structural div is not.
  const hasVisualFormatting =
    css['border'] !== undefined ||
    css['border-top'] !== undefined ||
    css['border-bottom'] !== undefined ||
    css['border-left'] !== undefined ||
    css['border-right'] !== undefined ||
    css['background'] !== undefined ||
    css['background-color'] !== undefined ||
    msoElement.startsWith('para-border-div');

  if (hasVisualFormatting) return [parseContainer(element, ctx, options, 'div')];

  const containsBlock = childNodesOf(element).some(
    (n) => isElement(n) && (isBlockLevelTag(tagNameOf(n)) || isBlockLikeSpecial(n, tagNameOf(n))),
  );
  if (containsBlock) {
    return parseBlockChildren(element, ctx, { ...options, sectionState });
  }
  // A div holding only inline content is a paragraph.
  return [parseParagraph(element, ctx, buildParagraphOptions(options))];
}

function isTextBox(css: Record<string, string>): boolean {
  const position = (css['position'] ?? '').toLowerCase();
  return position === 'absolute' && css['mso-position-horizontal'] !== undefined;
}

function parseContainer(
  element: Element,
  ctx: WordParseContext,
  options: BlockParseOptions,
  role: WordContainer['role'],
): WordContainer {
  const { css } = resolveElementCss(element, ctx);
  const container: WordContainer = {
    type: 'container',
    role,
    blocks: [],
    formatting: parseParagraphFormattingFromCss(css),
  };
  const className = attr(element, 'class');
  if (className) container.className = className;

  const runFormatting = parseRunFormattingFromCss(css);
  // As in table cells: a container's background is not a text highlight.
  delete runFormatting.highlight;
  const childOptions: BlockParseOptions = {
    depth: options.depth + 1,
    ...(options.sectionState ? { sectionState: options.sectionState } : {}),
  };
  const mergedRun = { ...(options.inheritedRun ?? {}), ...runFormatting };
  if (Object.keys(mergedRun).length > 0) childOptions.inheritedRun = mergedRun;
  if (options.inheritedParagraph) childOptions.inheritedParagraph = options.inheritedParagraph;

  container.blocks = parseBlockChildren(element, ctx, childOptions);
  return container;
}

/**
 * Real `<ol>`/`<ul>` lists.
 *
 * Word desktop does not emit these, but Word Online, Outlook and content that
 * has passed through another editor do. They are converted to the same
 * `listItem` model as `mso-list` paragraphs so downstream code has exactly one
 * representation of a list to deal with — with `fromHtmlList` recording where
 * it came from.
 */
function parseHtmlList(
  element: Element,
  ctx: WordParseContext,
  options: BlockParseOptions,
  level: number,
): WordBlock[] {
  const blocks: WordBlock[] = [];
  const tag = tagNameOf(element);
  const ordered = tag === 'ol';
  const { css } = resolveElementCss(element, ctx);
  const listId = css['mso-list'] ? (parseListReference(css['mso-list'])?.listId ?? `html-${level}`) : `html-${level}`;

  const startAttr = attr(element, 'start');
  const start = startAttr ? Number.parseInt(startAttr, 10) : undefined;
  const typeAttr = attr(element, 'type');

  let index = 0;
  for (const child of childNodesOf(element)) {
    if (!isElement(child)) continue;
    const childTag = tagNameOf(child);

    if (childTag === 'ol' || childTag === 'ul') {
      // A nested list that is a sibling of the items rather than inside one.
      blocks.push(...parseHtmlList(child, ctx, options, level + 1));
      continue;
    }
    if (childTag !== 'li' && childTag !== 'dt' && childTag !== 'dd') continue;

    // Pull nested lists out of the item so the item's own text stays clean.
    const nested: Element[] = [];
    for (const grandchild of childNodesOf(child)) {
      if (isElement(grandchild) && (tagNameOf(grandchild) === 'ol' || tagNameOf(grandchild) === 'ul')) {
        nested.push(grandchild);
        grandchild.parentNode?.removeChild(grandchild);
      }
    }

    const paragraph = parseParagraph(child, ctx, buildParagraphOptions(options));
    if (!paragraph.listItem) {
      const { css: itemCss } = resolveElementCss(child, ctx);
      const reference = findListReference(itemCss);
      const extracted = extractRenderedMarker(child, ctx);
      if (reference) {
        paragraph.listItem = buildListItem(reference, extracted, itemCss, ctx, child);
      } else {
        paragraph.listItem = {
          listId,
          level,
          fromHtmlList: true,
          marker: {
            type: ordered ? 'number' : 'bullet',
            source: 'html-list',
            ...(ordered
              ? {
                  numberFormat: numberFormatFromType(typeAttr),
                  levelText: '%1.',
                  ...(start !== undefined && index === 0 ? { startAt: start } : {}),
                }
              : { glyph: bulletForLevel(level) }),
          },
        };
      }
    }
    if (paragraph.listItem) {
      paragraph.listItem.level = Math.max(paragraph.listItem.level, level);
      paragraph.listItem.fromHtmlList = true;
    }
    blocks.push(paragraph);
    index++;

    for (const nestedList of nested) {
      blocks.push(...parseHtmlList(nestedList, ctx, options, level + 1));
    }
  }
  return blocks;
}

function numberFormatFromType(type: string | undefined): 'decimal' | 'lower-alpha' | 'upper-alpha' | 'lower-roman' | 'upper-roman' {
  switch ((type ?? '1').trim()) {
    case 'a':
      return 'lower-alpha';
    case 'A':
      return 'upper-alpha';
    case 'i':
      return 'lower-roman';
    case 'I':
      return 'upper-roman';
    default:
      return 'decimal';
  }
}

/** The browser's own default bullet sequence, matching Word's first levels. */
function bulletForLevel(level: number): string {
  const glyphs = ['•', 'o', '▪'];
  return glyphs[level % glyphs.length] ?? '•';
}

function parseHorizontalRule(element: Element, ctx: WordParseContext): WordBlock {
  const { css } = resolveElementCss(element, ctx);
  const formatting = parseParagraphFormattingFromCss(css);
  const rule: WordBlock = { type: 'horizontal-rule' };
  if (formatting.borders) rule.borders = formatting.borders;
  return rule;
}

/**
 * A live (non-comment) VML element at block level.
 *
 * As in the inline walker: the shape registry built from the raw text is the
 * source of truth, so a shape an `<img>` twin already claims produces nothing
 * here, and sub-parts (`<v:shapetype>`, `<v:stroke>`, …) are boilerplate rather
 * than drawings.
 */
function parseVmlBlock(element: Element, ctx: WordParseContext): WordBlock[] {
  const tag = tagNameOf(element);
  if (isVmlSubPart(tag)) return [];

  const shapeId = attr(element, 'id') ?? attr(element, 'o:spid');
  if (shapeId && ctx.claimedShapeIds.has(shapeId)) return [];

  const shape = shapeId ? ctx.vmlShapes.get(shapeId) : undefined;
  if (shape?.consumed) return [];
  if (shape?.imageSrc) {
    const image = parseVmlShapeImage(shape, ctx);
    if (image) {
      return [{ type: 'image-block', imageId: image.id, placement: image.placement, formatting: {} }];
    }
  }
  if (shape) shape.consumed = true;
  return [buildUnsupported(element, ctx, 'vml-shape', DiagnosticCode.WORD_VML_OBJECT)];
}

function buildUnsupported(
  element: Element,
  ctx: WordParseContext,
  objectType: string,
  code: string,
): WordUnsupportedObject {
  const block: WordUnsupportedObject = {
    type: 'unsupported',
    objectType,
    code,
    rawMetadata: describeUnsupportedObject(element),
    rawMarkup: excerpt(element, 2000),
  };
  const fallback = (element.textContent ?? '').replace(/\s+/g, ' ').trim();
  if (fallback) block.fallbackText = fallback;

  ctx.diagnostics.error(
    code,
    `A Word ${objectType.replace('-', ' ')} was found. It has no HTML equivalent, so its metadata and markup are preserved on the model and an inert placeholder is rendered in its place.`,
    { location: { tagName: tagNameOf(element), path: nodePath(element), excerpt: excerpt(element, 200) } },
  );
  return block;
}

/**
 * VML shapes that no `<img>` and no block claimed.
 *
 * These are drawings — a rectangle, an arrow, a canvas — that Word rendered
 * but the HTML flavour has no element for. They are appended as unsupported
 * objects rather than being dropped, so the fidelity report can count them.
 */
function appendOrphanShapes(blocks: WordBlock[], ctx: WordParseContext): void {
  for (const shape of ctx.vmlShapes.values()) {
    if (shape.consumed) continue;
    if (shape.id && ctx.claimedShapeIds.has(shape.id)) continue;
    if (shape.imageSrc) {
      const image = parseVmlShapeImage(shape, ctx);
      if (image) {
        blocks.push({ type: 'image-block', imageId: image.id, placement: image.placement, formatting: {} });
        continue;
      }
    }
    ctx.diagnostics.warn(
      DiagnosticCode.WORD_VML_SHAPE_APPROXIMATED,
      `A VML drawing (<${shape.tagName}>) was present in the payload with no picture inside it. Vector drawings have no HTML equivalent; the markup is preserved on the model and nothing is rendered for it.`,
      { details: { shapeId: shape.id, tagName: shape.tagName } },
    );
    blocks.push({
      type: 'unsupported',
      objectType: 'vml-shape',
      code: DiagnosticCode.WORD_VML_SHAPE_APPROXIMATED,
      rawMetadata: { id: shape.id, tagName: shape.tagName, style: shape.style ?? '' },
      rawMarkup: shape.raw,
    });
    shape.consumed = true;
  }
}

/** True when the element is Word's own list-paragraph class. */
export function isListParagraphClass(element: Element): boolean {
  return classList(element).some((c) => /^MsoListParagraph/i.test(c)) || hasClass(element, 'MsoListBullet');
}
