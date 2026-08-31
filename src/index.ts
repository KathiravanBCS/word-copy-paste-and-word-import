/**
 * word-clipboard-engine
 *
 * Microsoft Word clipboard HTML -> canonical Word Document Model -> clean HTML.
 *
 *     const payload  = captureFromPasteEvent(event);
 *     const document = parseWordClipboard(payload);   // the source of truth
 *     const html     = renderWordDocument(document).html;
 *
 * The model in the middle is the point. Everything downstream — this renderer,
 * an editor adapter, a DOCX writer — reads the model and never re-interprets
 * Word's semantics. Nothing in the core knows about any editor.
 */

import type { WordDocument } from './model/Document.js';
import type { ClipboardPayload } from './clipboard/ClipboardPayload.js';
import { clipboardPayloadFromHtml } from './clipboard/ClipboardPayload.js';
import { parseWordHtml } from './word/WordHtmlParser.js';
import { normalizeDocument, type NormalizeOptions } from './normalization/NormalizeDocument.js';
import type { ParseOptions } from './word/WordParseContext.js';
import { buildFidelityReport, type FidelityReport } from './diagnostics/FidelityReport.js';

/* -------------------------------------------------------------------------
 * Primary API
 * ---------------------------------------------------------------------- */

export interface ParseClipboardOptions extends ParseOptions {
  /** Normalisation options, or `false` to skip normalisation entirely. */
  normalize?: NormalizeOptions | false;
}

/**
 * Parse a captured clipboard payload into the canonical model.
 *
 * This is the function to call from a paste handler. It never mutates the
 * payload, and the returned document keeps the untouched `rawHtml` so a real
 * paste can always be re-examined.
 */
export function parseWordClipboard(
  payload: ClipboardPayload,
  options: ParseClipboardOptions = {},
): WordDocument {
  const { document } = parseWordHtml(payload.html ?? '', options, payload);
  if (options.normalize !== false) {
    normalizeDocument(document, options.normalize ?? {});
  }
  return document;
}

/** Parse a raw Word HTML string. Convenient for fixtures, tests and servers. */
export function parseWordHtmlString(
  html: string,
  options: ParseClipboardOptions = {},
): WordDocument {
  return parseWordClipboard(clipboardPayloadFromHtml(html), options);
}

/** Build the fidelity report for a parsed document. */
export function getFidelityReport(document: WordDocument): FidelityReport {
  return buildFidelityReport(document);
}

/* -------------------------------------------------------------------------
 * Re-exports
 * ---------------------------------------------------------------------- */

// Model — the contract every consumer codes against.
export type * from './model/index.js';
export { isParagraph, isTable, isContainer, walkBlocks } from './model/Block.js';

// Clipboard capture.
export {
  captureFromPasteEvent,
  captureFromNavigatorClipboard,
  installPasteCapture,
  materialiseImages,
  type CaptureOptions,
} from './clipboard/ClipboardCapture.js';
export {
  clipboardPayloadFromHtml,
  emptyClipboardPayload,
  type ClipboardPayload,
  type ClipboardImageItem,
  type RawClipboardDocument,
} from './clipboard/ClipboardPayload.js';

// Detection.
export {
  detectWordHtml,
  detectWordPayload,
  type WordDetectionResult,
  type DetectionOptions,
  type WordSource,
} from './detection/WordDetector.js';
export {
  detectSignals,
  WORD_SIGNALS,
  SIGNAL_WEIGHTS,
  type DetectedSignal,
  type SignalStrength,
} from './detection/WordSignalDetector.js';

// Parsing internals, exposed for tooling and for building other back ends.
export { parseWordHtml, type WordHtmlParseResult } from './word/WordHtmlParser.js';
export { type ParseOptions, type WordParseContext, type VmlShape } from './word/WordParseContext.js';
export {
  parseWordStyleSheet,
  parseWordCss,
  findStyle,
  resolveStyleChain,
  normaliseStyleId,
  createEmptyStyleSheet,
} from './word/WordStyleParser.js';
export {
  parseListRules,
  parseListReference,
  parseListLevel,
  expandLevelText,
  formatNumber,
  toAlpha,
  toRoman,
  toOrdinal,
  toCardinalText,
  toOrdinalText,
} from './word/WordListStyleParser.js';
export {
  parseWordLength,
  parsePercent,
  parseBoxShorthand,
  lengthFromPx,
  lengthFromPt,
  lengthFromTwips,
  lengthToCss,
  pxToPt,
  ptToPx,
  PX_PER_INCH,
  PT_PER_INCH,
  TWIPS_PER_INCH,
} from './word/WordLengthParser.js';
export { parseWordColor, parseHighlight } from './word/WordColorParser.js';
export {
  resolveSymbolGlyph,
  decodeMsoLevelText,
  isSymbolFont,
  type SymbolMappingResult,
} from './word/WordSymbolFonts.js';
export {
  classifyComment,
  extractStyleBlocks,
  extractHiddenConditionalBlocks,
  type ConditionalCommentInfo,
} from './word/WordConditionalCommentParser.js';
export { tokenizeCss, parseInlineStyle, parseDeclarations } from './word/WordCssTokenizer.js';

// Normalisation.
export {
  normalizeDocument,
  type NormalizeOptions,
  type NormalizeResult,
} from './normalization/NormalizeDocument.js';
export {
  normalizeLists,
  buildListTree,
  type ListTreeNode,
  type ListTreeItem,
} from './normalization/NormalizeLists.js';
export {
  computeListIndentation,
  type ListIndentation,
} from './normalization/NormalizeUnits.js';

// Rendering.
export {
  renderWordDocument,
  renderWordDocumentToHtml,
  type RenderOptions,
  type RenderResult,
} from './rendering/HtmlRenderer.js';
export {
  renderStandaloneHtml,
  toHtmlBlob,
  suggestFileName,
  type DocumentRenderOptions,
  type DocumentRenderResult,
} from './rendering/HtmlDocumentRenderer.js';
export { type ListMarkerMode } from './rendering/HtmlListRenderer.js';

// Diagnostics.
export {
  DiagnosticCode,
  DiagnosticExplanations,
  type WordDiagnostic,
  type DiagnosticSeverity,
  type FidelityClass,
} from './diagnostics/UnsupportedFeature.js';
export { DiagnosticCollector } from './diagnostics/DiagnosticCollector.js';
export {
  buildFidelityReport,
  formatFidelityReport,
  type FidelityReport,
} from './diagnostics/FidelityReport.js';

// Security helpers, so a host application can apply the same rules.
export {
  checkLinkUrl,
  checkImageUrl,
  sanitizeTree,
  DEFAULT_LIMITS,
  type ParseLimits,
} from './util/security.js';
