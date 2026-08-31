import type { WordStyleSheet } from '../model/Style.js';
import type { WordImage } from '../model/Image.js';
import type { WordBookmark, WordHyperlink } from '../model/Hyperlink.js';
import type { ClipboardImageItem } from '../clipboard/ClipboardPayload.js';
import type { DiagnosticCollector } from '../diagnostics/DiagnosticCollector.js';
import { DEFAULT_LIMITS, type ParseLimits } from '../util/security.js';
import { createEmptyStyleSheet } from './WordStyleParser.js';

/** A VML shape harvested from a conditional comment or the live tree. */
export interface VmlShape {
  id: string;
  /** `src` from `<v:imagedata>`, when the shape wraps a picture. */
  imageSrc?: string;
  /** `o:title` from `<v:imagedata>`. */
  title?: string;
  /** Inline style on the shape, e.g. `width:451.2pt;height:184.8pt`. */
  style?: string;
  /** `<v:imagedata cropleft="…">` percentages, as written. */
  crop?: { top?: string; right?: string; bottom?: string; left?: string };
  /** The shape's element name, e.g. `v:shape`, `v:rect`. */
  tagName: string;
  /** The shape's raw markup, preserved for diagnostics. */
  raw: string;
  /** True once the shape has been consumed by an `<img>`. */
  consumed?: boolean;
}

export interface ParseOptions {
  /**
   * Treat the payload as Word even if detection says otherwise. Useful for
   * fixtures and for pasting a saved Word HTML file that lost its meta tags.
   */
  forceWord?: boolean;
  /** Detection threshold overrides. */
  detectionThreshold?: number;
  /** Parse limits. */
  limits?: Partial<ParseLimits>;
  /**
   * When Word's rendered marker span is missing, recover the marker from
   * leading text in the paragraph. Default true; every recovery is diagnosed.
   */
  recoverMarkersFromText?: boolean;
  /**
   * Assign unmatched clipboard image blobs to unresolved `file:///` images in
   * document order. Default true; every assignment is diagnosed, because the
   * pairing is positional rather than certain.
   */
  matchClipboardImages?: boolean;
  /** Keep Word's `WordSectionN` divs as container blocks. Default false. */
  keepSectionContainers?: boolean;
  /** Maximum diagnostics retained. */
  maxDiagnostics?: number;
}

/**
 * Mutable state shared by every content parser during a single parse.
 *
 * Passing one context around is what keeps the parsers pure functions of
 * (node, context) rather than a class hierarchy with hidden state, and it is
 * what makes the id allocation for images, links and bookmarks deterministic —
 * which in turn is what makes golden model fixtures stable.
 */
export interface WordParseContext {
  sheet: WordStyleSheet;
  diagnostics: DiagnosticCollector;
  images: Record<string, WordImage>;
  hyperlinks: Record<string, WordHyperlink>;
  bookmarks: Record<string, WordBookmark>;
  /** VML shapes keyed by shape id. */
  vmlShapes: Map<string, VmlShape>;
  /**
   * Shape ids that an `<img v:shapes="…">` twin claims.
   *
   * Word emits every picture twice — once as VML for Word/IE, once as an
   * `<img>` for everyone else — and which of the two the DOM exposes as live
   * elements depends on how the host parses downlevel-hidden conditional
   * comments. Knowing up front which shapes have an `<img>` twin is what stops
   * the same picture being modelled twice.
   */
  claimedShapeIds: Set<string>;
  /** Image blobs from the clipboard, consumed in document order. */
  clipboardImages: ClipboardImageItem[];
  limits: ParseLimits;
  options: ParseOptions;
  /** Mutable counters. */
  state: {
    imageSeq: number;
    hyperlinkSeq: number;
    bookmarkSeq: number;
    clipboardImageCursor: number;
    nodesVisited: number;
    blocksProduced: number;
    budgetExceeded: boolean;
  };
}

export function createParseContext(
  diagnostics: DiagnosticCollector,
  options: ParseOptions = {},
  sheet: WordStyleSheet = createEmptyStyleSheet(),
  clipboardImages: ClipboardImageItem[] = [],
): WordParseContext {
  return {
    sheet,
    diagnostics,
    images: {},
    hyperlinks: {},
    bookmarks: {},
    vmlShapes: new Map(),
    claimedShapeIds: new Set(),
    clipboardImages,
    limits: { ...DEFAULT_LIMITS, ...(options.limits ?? {}) },
    options,
    state: {
      imageSeq: 0,
      hyperlinkSeq: 0,
      bookmarkSeq: 0,
      clipboardImageCursor: 0,
      nodesVisited: 0,
      blocksProduced: 0,
      budgetExceeded: false,
    },
  };
}

export function nextImageId(ctx: WordParseContext): string {
  return `img-${++ctx.state.imageSeq}`;
}

export function nextHyperlinkId(ctx: WordParseContext): string {
  return `link-${++ctx.state.hyperlinkSeq}`;
}

export function nextBookmarkId(ctx: WordParseContext): string {
  return `bm-${++ctx.state.bookmarkSeq}`;
}

/**
 * Charge a node against the parse budget.
 *
 * Returns false once the budget is spent, at which point every parser stops
 * descending. A 100-page paste is legitimate; an adversarially nested one is
 * not, and neither should be able to wedge the tab.
 */
export function chargeNode(ctx: WordParseContext): boolean {
  ctx.state.nodesVisited++;
  if (ctx.state.nodesVisited <= ctx.limits.maxNodes) return true;
  if (!ctx.state.budgetExceeded) {
    ctx.state.budgetExceeded = true;
    ctx.diagnostics.error(
      'LIMIT_NODE_BUDGET_EXCEEDED',
      `The pasted document exceeded the node budget of ${ctx.limits.maxNodes}; parsing stopped early and the result is truncated.`,
      { details: { limit: ctx.limits.maxNodes } },
    );
  }
  return false;
}

export function chargeBlock(ctx: WordParseContext): boolean {
  ctx.state.blocksProduced++;
  if (ctx.state.blocksProduced <= ctx.limits.maxBlocks) return true;
  if (!ctx.state.budgetExceeded) {
    ctx.state.budgetExceeded = true;
    ctx.diagnostics.error(
      'LIMIT_DOCUMENT_TRUNCATED',
      `The pasted document exceeded the block budget of ${ctx.limits.maxBlocks}; the result is truncated.`,
      { details: { limit: ctx.limits.maxBlocks } },
    );
  }
  return false;
}
