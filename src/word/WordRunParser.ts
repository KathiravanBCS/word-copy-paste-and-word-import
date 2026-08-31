import type { RunFormatting } from '../model/Style.js';
import type { WordRun, WordTextRun } from '../model/Run.js';
import {
  attr,
  childNodesOf,
  collapseWhitespace,
  excerpt,
  isCommentNode,
  isElement,
  isTextNode,
  nodePath,
  tagNameOf,
} from '../util/dom.js';
import { DiagnosticCode } from '../diagnostics/UnsupportedFeature.js';
import { parseInlineStyle } from './WordCssTokenizer.js';
import {
  mergeRunFormatting,
  parseRunFormattingFromCss,
  primaryFontFamily,
} from './WordFormattingParser.js';
import { parseWordColor } from './WordColorParser.js';
import { lengthFromPt } from './WordLengthParser.js';
import { normaliseStyleId, resolveStyleChain } from './WordStyleParser.js';
import { isVmlSubPart, parseImageElement, parseVmlShapeImage } from './WordImageParser.js';
import { isBookmarkAnchor, parseBookmark, parseHyperlink } from './WordHyperlinkParser.js';
import { chargeNode, type WordParseContext } from './WordParseContext.js';
import { classifyComment } from './WordConditionalCommentParser.js';

/**
 * The inline walker: turns a paragraph's inline content into runs.
 *
 * The rule that shapes this file is that **formatting is never flattened
 * across the paragraph**. `Hello <b>world</b>` is two runs, not one paragraph
 * with a bold flag. Word nests spans several deep for a single word, and each
 * level may change one property, so formatting is threaded down as an
 * inherited context and each text node captures the exact state in effect
 * where it sits.
 *
 * Whitespace is collapsed the way a browser would, with two exceptions Word
 * depends on: `mso-spacerun:yes` spans hold significant runs of spaces, and
 * `mso-tab-count` spans are tabs written as spaces. Non-breaking spaces are
 * never collapsed — Word uses them as real content.
 */

/** Mutable whitespace state carried across sibling nodes in one paragraph. */
interface InlineState {
  runs: WordRun[];
  /** True when nothing has been emitted yet, so a leading space is dropped. */
  atStart: boolean;
  /** True when the last emitted character was a collapsible space. */
  pendingSpace: boolean;
  depth: number;
}

export interface InlineContext {
  /** Formatting inherited from ancestors. */
  formatting: RunFormatting;
  /** Hyperlink id in effect, if any. */
  hyperlinkId?: string;
  /** Style names in effect, outermost first. */
  styleChain: string[];
  /** True inside an `mso-spacerun` span, where spaces are significant. */
  preserveSpace: boolean;
  /** Bookmarks opened but not yet attached to a run. */
  pendingBookmarks: string[];
}

export function createInlineContext(formatting: RunFormatting = {}): InlineContext {
  return { formatting, styleChain: [], preserveSpace: false, pendingBookmarks: [] };
}

/** Parse a list of inline nodes into runs. */
export function parseInlineNodes(
  nodes: Node[],
  ctx: WordParseContext,
  inline: InlineContext,
): WordRun[] {
  const state: InlineState = { runs: [], atStart: true, pendingSpace: false, depth: 0 };
  for (const node of nodes) {
    walkInline(node, ctx, inline, state);
  }
  trimTrailingSpace(state);
  return state.runs;
}

function walkInline(
  node: Node,
  ctx: WordParseContext,
  inline: InlineContext,
  state: InlineState,
): void {
  if (!chargeNode(ctx)) return;
  if (state.depth > ctx.limits.maxDepth) {
    ctx.diagnostics.warn(
      DiagnosticCode.LIMIT_DEPTH_EXCEEDED,
      `Inline nesting exceeded ${ctx.limits.maxDepth} levels; deeper content was not descended into.`,
    );
    return;
  }

  if (isTextNode(node)) {
    emitText(node.data ?? '', ctx, inline, state);
    return;
  }
  if (isCommentNode(node)) {
    // A conditional comment reaching the inline walker means the paragraph
    // parser has already taken what it needed from it (list markers). VML
    // hidden here still deserves a note rather than a silent drop.
    const info = classifyComment(node.data ?? '');
    if (info.payload === 'vml' && info.content) {
      ctx.diagnostics.info(
        DiagnosticCode.WORD_VML_OBJECT,
        'A VML drawing was found inside a paragraph. Where it wrapped a picture the picture was extracted; the VML markup itself is preserved on the model and not rendered.',
        { location: { path: nodePath(node) }, fidelity: 'EQUIVALENT' },
      );
    }
    return;
  }
  if (!isElement(node)) return;

  const element = node;
  const tag = tagNameOf(element);

  switch (tag) {
    case 'br':
      emitBreak(element, ctx, inline, state);
      return;
    case 'img':
      emitImage(element, ctx, inline, state);
      return;
    case 'a':
      emitAnchor(element, ctx, inline, state);
      return;
    case 'o:p':
      // Word's paragraph mark. It carries no content of its own; an
      // `&nbsp;` inside it means "this paragraph is intentionally empty",
      // which the paragraph parser reads separately.
      return;
    case 'script':
    case 'style':
      return;
    default:
      break;
  }

  if (tag.startsWith('v:')) {
    emitVmlInline(element, ctx, inline, state);
    return;
  }
  if (tag === 'object' || tag === 'o:oleobject' || tag === 'embed') {
    ctx.diagnostics.error(
      DiagnosticCode.WORD_OLE_OBJECT,
      'An embedded object was found inside a paragraph. OLE objects have no HTML equivalent; the object metadata is preserved on the model and its fallback text is kept.',
      { location: { tagName: tag, path: nodePath(element), excerpt: excerpt(element) } },
    );
    // Keep whatever visible fallback Word rendered rather than dropping it.
    const nested = deriveContext(element, ctx, inline);
    state.depth++;
    for (const child of childNodesOf(element)) walkInline(child, ctx, nested, state);
    state.depth--;
    return;
  }

  const nested = deriveContext(element, ctx, inline);

  // `mso-tab-count` spans are tabs Word wrote as runs of spaces.
  const tabCount = readTabCount(element);
  if (tabCount > 0) {
    for (let i = 0; i < tabCount; i++) {
      state.runs.push({ type: 'tab', formatting: nested.formatting });
    }
    state.atStart = false;
    state.pendingSpace = false;
    return;
  }

  const special = readSpecialCharacter(element);
  if (special === 'line-break' || special === 'page-break') {
    state.runs.push({
      type: 'break',
      breakType: special === 'page-break' ? 'page' : 'line',
      formatting: nested.formatting,
    });
    state.atStart = false;
    state.pendingSpace = false;
    return;
  }
  if (special === 'comment') {
    ctx.diagnostics.info(
      DiagnosticCode.WORD_UNSUPPORTED_FIELD,
      'A Word comment reference was found. Comment anchors are not represented in the output; the comment text, when present in the payload, is reported separately.',
      { location: { path: nodePath(element) }, fidelity: 'APPROXIMATED' },
    );
    return;
  }

  if (tag === 'del' || tag === 'ins') {
    ctx.diagnostics.info(
      DiagnosticCode.WORD_REVISION_MARK_FLATTENED,
      `Tracked-change markup (<${tag}>) was flattened into ordinary formatting; the revision itself is not represented.`,
      { location: { tagName: tag }, fidelity: 'APPROXIMATED' },
    );
  }

  state.depth++;
  for (const child of childNodesOf(element)) {
    walkInline(child, ctx, nested, state);
  }
  state.depth--;
}

/** Build the inline context for an element from its tag, attributes and style. */
function deriveContext(
  element: Element,
  ctx: WordParseContext,
  inline: InlineContext,
): InlineContext {
  const tag = tagNameOf(element);
  let formatting: RunFormatting = inline.formatting;
  const styleChain = inline.styleChain;

  // 1. Class-based character style from the clipboard stylesheet.
  const classes = (attr(element, 'class') ?? '').split(/\s+/).filter(Boolean);
  let nextStyleChain = styleChain;
  for (const className of classes) {
    const definition = ctx.sheet.styles[normaliseStyleId(className)];
    if (!definition) continue;
    const declarations = resolveStyleChain(ctx.sheet, definition.id);
    formatting = mergeRunFormatting(formatting, parseRunFormattingFromCss(declarations));
    if (!nextStyleChain.includes(definition.name)) {
      nextStyleChain = [...nextStyleChain, definition.name];
    }
  }

  // 2. Presentational elements. Word emits these alongside CSS.
  const tagFormatting = formattingForTag(tag, element);
  if (tagFormatting) formatting = mergeRunFormatting(formatting, tagFormatting);

  // 3. Inline style wins over everything above it.
  const style = parseInlineStyle(attr(element, 'style'));
  if (Object.keys(style).length > 0) {
    formatting = mergeRunFormatting(formatting, parseRunFormattingFromCss(style));
  }

  // 4. `lang` attribute.
  const lang = attr(element, 'lang');
  if (lang) formatting = { ...formatting, language: lang };

  const preserveSpace =
    inline.preserveSpace ||
    /yes/i.test(style['mso-spacerun'] ?? '') ||
    /pre/i.test(style['white-space'] ?? '');

  const next: InlineContext = {
    formatting,
    styleChain: nextStyleChain,
    preserveSpace,
    pendingBookmarks: inline.pendingBookmarks,
  };
  if (inline.hyperlinkId) next.hyperlinkId = inline.hyperlinkId;
  return next;
}

function formattingForTag(tag: string, element: Element): RunFormatting | null {
  switch (tag) {
    case 'b':
    case 'strong':
      return { bold: true };
    case 'i':
    case 'em':
    case 'cite':
    case 'var':
    case 'address':
      return { italic: true };
    case 'u':
    case 'ins':
      return { underline: 'single' };
    case 's':
    case 'strike':
    case 'del':
      return { strike: true };
    case 'sup':
      return { verticalAlign: 'super' };
    case 'sub':
      return { verticalAlign: 'sub' };
    case 'code':
    case 'kbd':
    case 'samp':
    case 'tt':
      return { fontFamily: 'monospace' };
    case 'small':
      return {};
    case 'font':
      return formattingForFontElement(element);
    default:
      return null;
  }
}

/** `<font face size color>` — Word still emits this in some payloads. */
function formattingForFontElement(element: Element): RunFormatting {
  const formatting: RunFormatting = {};
  const face = attr(element, 'face');
  if (face) {
    formatting.fontFamilyRaw = face;
    formatting.fontFamily = primaryFontFamily(face);
  }
  const color = parseWordColor(attr(element, 'color'));
  if (color) formatting.color = color;
  const size = attr(element, 'size');
  if (size) {
    // The legacy 1..7 scale, in points, as Word's own exporter uses it.
    const scale: Record<string, number> = { '1': 7.5, '2': 10, '3': 12, '4': 13.5, '5': 18, '6': 24, '7': 36 };
    const pt = scale[size.trim()];
    if (pt) formatting.fontSize = lengthFromPt(pt);
  }
  return formatting;
}

function readTabCount(element: Element): number {
  const style = parseInlineStyle(attr(element, 'style'));
  const value = style['mso-tab-count'];
  if (!value) return 0;
  const count = Number.parseInt(value, 10);
  return Number.isFinite(count) && count > 0 ? Math.min(count, 64) : 0;
}

function readSpecialCharacter(element: Element): string | null {
  const style = parseInlineStyle(attr(element, 'style'));
  const special = style['mso-special-character'];
  if (!special) return null;
  const lower = special.trim().toLowerCase();
  if (lower === 'line-break') {
    return /always/i.test(style['page-break-before'] ?? '') ? 'page-break' : 'line-break';
  }
  if (lower === 'comment') return 'comment';
  return lower;
}

function emitText(
  data: string,
  ctx: WordParseContext,
  inline: InlineContext,
  state: InlineState,
): void {
  if (data.length === 0) return;

  let text: string;
  if (inline.preserveSpace) {
    text = data;
  } else {
    text = collapseWhitespace(data);
    if (text === ' ') {
      // A whitespace-only node between elements: remember it, but do not emit
      // a run for it until real content follows.
      if (!state.atStart) state.pendingSpace = true;
      return;
    }
    if (text.startsWith(' ')) {
      if (state.atStart || state.pendingSpace) text = text.slice(1);
    }
  }
  if (text.length === 0) return;

  if (state.pendingSpace && !state.atStart) {
    appendText(' ', ctx, inline, state);
    state.pendingSpace = false;
  }
  appendText(text, ctx, inline, state);
  state.atStart = false;
  state.pendingSpace = false;
}

function appendText(
  text: string,
  _ctx: WordParseContext,
  inline: InlineContext,
  state: InlineState,
): void {
  const previous = state.runs[state.runs.length - 1];
  // Merge only into a run created from the *same* inline context, so a
  // formatting boundary always starts a new run.
  if (
    previous &&
    previous.type === 'text' &&
    previous.formatting === inline.formatting &&
    previous.hyperlinkId === inline.hyperlinkId
  ) {
    previous.text += text;
    return;
  }
  const run: WordTextRun = { type: 'text', text, formatting: inline.formatting };
  if (inline.hyperlinkId) run.hyperlinkId = inline.hyperlinkId;
  if (inline.styleChain.length > 0) run.styleChain = inline.styleChain;
  if (inline.pendingBookmarks.length > 0) {
    run.bookmarks = [...inline.pendingBookmarks];
    inline.pendingBookmarks.length = 0;
  }
  state.runs.push(run);
}

function emitBreak(
  element: Element,
  ctx: WordParseContext,
  inline: InlineContext,
  state: InlineState,
): void {
  const nested = deriveContext(element, ctx, inline);
  const style = parseInlineStyle(attr(element, 'style'));
  const isPageBreak = /always|left|right/i.test(style['page-break-before'] ?? '');
  state.runs.push({
    type: 'break',
    breakType: isPageBreak ? 'page' : 'line',
    formatting: nested.formatting,
  });
  state.atStart = false;
  state.pendingSpace = false;
}

function emitImage(
  element: Element,
  ctx: WordParseContext,
  inline: InlineContext,
  state: InlineState,
): void {
  const nested = deriveContext(element, ctx, inline);
  const image = parseImageElement(element, ctx);
  const run: WordRun = { type: 'image', imageId: image.id, formatting: nested.formatting };
  if (inline.hyperlinkId) run.hyperlinkId = inline.hyperlinkId;
  state.runs.push(run);
  state.atStart = false;
  state.pendingSpace = false;
}

function emitAnchor(
  element: Element,
  ctx: WordParseContext,
  inline: InlineContext,
  state: InlineState,
): void {
  if (isBookmarkAnchor(element)) {
    const bookmark = parseBookmark(element, ctx);
    const nested = deriveContext(element, ctx, inline);
    if (bookmark) nested.pendingBookmarks = [...inline.pendingBookmarks, bookmark.name];
    state.depth++;
    for (const child of childNodesOf(element)) walkInline(child, ctx, nested, state);
    state.depth--;
    // An empty bookmark anchor still marks a position; hand the name to the
    // next run rather than losing it.
    if (bookmark && nested.pendingBookmarks.length > 0) {
      inline.pendingBookmarks.push(bookmark.name);
    }
    return;
  }

  const link = parseHyperlink(element, ctx);
  const nested = deriveContext(element, ctx, inline);
  if (link) nested.hyperlinkId = link.id;
  state.depth++;
  for (const child of childNodesOf(element)) walkInline(child, ctx, nested, state);
  state.depth--;
}

/**
 * A live (non-comment) VML element inside a paragraph.
 *
 * Every VML shape in the payload has already been harvested from the raw text
 * by `collectVmlShapes`, whichever way the host's HTML parser treated the
 * conditional comment around it. So the live element is not a second source of
 * truth — it is a duplicate, and the only questions are whether it has an
 * `<img>` twin (in which case the twin renders the picture) and whether it
 * holds text worth keeping.
 */
function emitVmlInline(
  element: Element,
  ctx: WordParseContext,
  inline: InlineContext,
  state: InlineState,
): void {
  const tag = tagNameOf(element);
  if (isVmlSubPart(tag)) return;

  const shapeId = attr(element, 'id') ?? attr(element, 'o:spid');
  const shape = shapeId ? ctx.vmlShapes.get(shapeId) : undefined;
  const claimed = shapeId ? ctx.claimedShapeIds.has(shapeId) : false;

  if (claimed) {
    // The `<img>` twin renders this picture. Nothing to do here.
    return;
  }

  if (shape?.imageSrc && !shape.consumed) {
    const image = parseVmlShapeImage(shape, ctx);
    if (image) {
      const nested = deriveContext(element, ctx, inline);
      state.runs.push({ type: 'image', imageId: image.id, formatting: nested.formatting });
      state.atStart = false;
      state.pendingSpace = false;
      ctx.diagnostics.info(
        DiagnosticCode.WORD_VML_OBJECT,
        'A VML picture shape with no <img> twin was converted to an image. The original VML markup is preserved on the image model.',
        { location: { tagName: tag }, fidelity: 'EQUIVALENT' },
      );
      return;
    }
  }
  if (shape?.consumed) return;

  ctx.diagnostics.warn(
    DiagnosticCode.WORD_VML_SHAPE_APPROXIMATED,
    `A VML drawing (<${tag}>) has no HTML equivalent and was not rendered. Its markup is preserved in the diagnostic details.`,
    {
      location: { tagName: tag, path: nodePath(element), excerpt: excerpt(element, 400) },
      fidelity: 'UNSUPPORTED',
    },
  );
  if (shape) shape.consumed = true;

  // Keep any text Word drew inside the shape (a VML text box).
  const nested = deriveContext(element, ctx, inline);
  state.depth++;
  for (const child of childNodesOf(element)) walkInline(child, ctx, nested, state);
  state.depth--;
}

function trimTrailingSpace(state: InlineState): void {
  for (let i = state.runs.length - 1; i >= 0; i--) {
    const run = state.runs[i]!;
    if (run.type !== 'text') return;
    const trimmed = run.text.replace(/ +$/, '');
    if (trimmed.length === 0) {
      state.runs.splice(i, 1);
      continue;
    }
    run.text = trimmed;
    return;
  }
}

/** True when a run list contains no visible content. */
export function runsAreEmpty(runs: WordRun[]): boolean {
  for (const run of runs) {
    if (run.type === 'image' || run.type === 'break' || run.type === 'tab') return false;
    if (run.type === 'text' || run.type === 'field' || run.type === 'note') {
      if (run.text.replace(/[\s\u00a0]/g, '').length > 0) return false;
    }
  }
  return true;
}

/** Concatenate the text of a run list, for marker heuristics and diagnostics. */
export function runsToText(runs: WordRun[]): string {
  let out = '';
  for (const run of runs) {
    if (run.type === 'text' || run.type === 'field' || run.type === 'note') out += run.text;
    else if (run.type === 'tab') out += '\t';
    else if (run.type === 'break') out += '\n';
  }
  return out;
}
