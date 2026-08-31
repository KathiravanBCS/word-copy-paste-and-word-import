import type { ListItemInfo, ListItemMarker, WordListReference } from '../model/List.js';
import type { WordListLevel } from '../model/ListLevel.js';
import type { RunFormatting } from '../model/Style.js';
import {
  attr,
  childNodesOf,
  isCommentNode,
  isElement,
  nodePath,
  tagNameOf,
} from '../util/dom.js';
import { DiagnosticCode } from '../diagnostics/UnsupportedFeature.js';
import { parseInlineStyle } from './WordCssTokenizer.js';
import { parseRunFormattingFromCss } from './WordFormattingParser.js';
import { parseWordLength } from './WordLengthParser.js';
import { parseListReference } from './WordListStyleParser.js';
import { isSymbolFont, resolveSymbolGlyph } from './WordSymbolFonts.js';
import { isEndIf, isListMarkerOpen } from './WordConditionalCommentParser.js';
import type { WordParseContext } from './WordParseContext.js';

/**
 * Recovering a list item from a Word paragraph.
 *
 * Word does not emit `<ol>`/`<ul>` from the desktop app. A list item is an
 * ordinary paragraph carrying an `mso-list` declaration, with the marker Word
 * *drew* sitting inside the paragraph as literal text:
 *
 *     <p class=MsoListParagraphCxSpFirst
 *        style='text-indent:-.25in;mso-list:l0 level1 lfo1'>
 *       <![if !supportLists]>
 *         <span style='font-family:Symbol'>
 *           <span style='mso-list:Ignore'>·
 *             <span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp; </span>
 *           </span>
 *         </span>
 *       <![endif]>
 *       Parent
 *     </p>
 *
 * Leaving that marker where it is produces `"·      Parent"` as paragraph
 * text — a bullet glyph welded into a text node, unselectable as a list, wrong
 * when re-indented, and wrong when re-numbered. That is the single defect this
 * module exists to prevent.
 *
 * So the marker is *lifted out*: removed from the content nodes and turned
 * into structured marker data, combined with the `@list` definition, which
 * says what the marker means. What comes back is a list item whose content is
 * `"Parent"` and whose marker is `{ type: 'bullet', glyph: '•' }`.
 */

export interface ExtractedMarker {
  marker: ListItemMarker;
  /** The nodes that carried the marker, already detached from the paragraph. */
  removedNodes: Node[];
}

/** Find the `mso-list` reference on a paragraph, from inline style or class style. */
export function findListReference(
  css: Record<string, string>,
): WordListReference | null {
  return parseListReference(css['mso-list']);
}

/**
 * Remove Word's rendered marker from a paragraph element and return it.
 *
 * Two markup shapes are handled, because Word emits both depending on version
 * and on whether the copy went through the "filtered" HTML path:
 *
 *  1. `<![if !supportLists]> … <![endif]>` bracketing the marker.
 *  2. A bare `<span style='mso-list:Ignore'>` with no conditional comment.
 *
 * The element is mutated — it is the working clone, never the raw payload.
 */
export function extractRenderedMarker(
  paragraph: Element,
  ctx: WordParseContext,
): ExtractedMarker | null {
  const bracketed = extractBracketedMarker(paragraph);
  if (bracketed) return buildMarkerFromNodes(bracketed.markerNodes, bracketed.removedNodes, ctx);

  const ignoreSpan = findIgnoreSpan(paragraph);
  if (ignoreSpan) {
    // Remove the outermost wrapper that exists only to hold the marker, so no
    // empty span is left behind to produce a phantom run.
    let target: Element = ignoreSpan;
    while (
      target.parentElement &&
      target.parentElement !== paragraph &&
      target.parentElement.childNodes.length === 1
    ) {
      target = target.parentElement;
    }
    target.parentNode?.removeChild(target);
    return buildMarkerFromNodes([target], [target], ctx);
  }

  return null;
}

function extractBracketedMarker(
  paragraph: Element,
): { markerNodes: Node[]; removedNodes: Node[] } | null {
  const children = childNodesOf(paragraph);
  let openIndex = -1;
  for (let i = 0; i < children.length; i++) {
    const node = children[i]!;
    if (isCommentNode(node) && isListMarkerOpen(node.data ?? '')) {
      openIndex = i;
      break;
    }
  }
  if (openIndex === -1) return null;

  const removedNodes: Node[] = [children[openIndex]!];
  const markerNodes: Node[] = [];
  for (let i = openIndex + 1; i < children.length; i++) {
    const node = children[i]!;
    removedNodes.push(node);
    if (isCommentNode(node) && isEndIf(node.data ?? '')) break;
    markerNodes.push(node);
  }
  for (const node of removedNodes) node.parentNode?.removeChild(node);
  return { markerNodes, removedNodes };
}

/** Depth-first search for a `span[style*="mso-list:Ignore"]`. */
function findIgnoreSpan(root: Element): Element | null {
  const stack: Node[] = childNodesOf(root);
  while (stack.length > 0) {
    const node = stack.shift()!;
    if (!isElement(node)) continue;
    const style = attr(node, 'style');
    if (style && /mso-list\s*:\s*ignore/i.test(style)) return node;
    stack.push(...childNodesOf(node));
  }
  return null;
}

function buildMarkerFromNodes(
  markerNodes: Node[],
  removedNodes: Node[],
  ctx: WordParseContext,
): ExtractedMarker {
  const text = markerText(markerNodes);
  const formatting = markerFormatting(markerNodes);
  const font = formatting.fontFamily;

  const marker: ListItemMarker = { type: 'number', source: 'mso-list-ignore' };
  if (Object.keys(formatting).length > 0) marker.formatting = formatting;
  if (text) marker.text = text;

  if (looksLikeBullet(text, font)) {
    marker.type = 'bullet';
    const resolved = resolveSymbolGlyph(text, font);
    marker.glyph = resolved.glyph;
    marker.rawGlyph = resolved.rawGlyph;
    if (resolved.font) marker.font = resolved.font;
    if (resolved.mapped) {
      marker.fontMapped = true;
      ctx.diagnostics.info(
        DiagnosticCode.WORD_SYMBOL_FONT_MAPPED,
        `A list bullet drawn as "${resolved.rawGlyph}" in ${resolved.font ?? 'a symbol font'} was mapped to "${resolved.glyph}". The raw glyph and its font are preserved on the marker.`,
        { details: { font: resolved.font ?? '', glyph: resolved.glyph }, fidelity: 'EQUIVALENT' },
      );
    } else if (resolved.unmapped) {
      ctx.diagnostics.warn(
        DiagnosticCode.WORD_SYMBOL_FONT_UNMAPPED,
        `A list bullet is a ${resolved.font ?? 'symbol'} font byte with no known Unicode equivalent; it is rendered as-is in its original font.`,
        { details: { font: resolved.font ?? '', codePoint: resolved.codePoint ?? -1 } },
      );
    }
  }

  return { marker, removedNodes };
}

/**
 * The marker's own text, with Word's spacer stripped.
 *
 * Word pads the marker to the tab stop with a small-font run of non-breaking
 * spaces inside the same span. Those are layout, not marker text.
 */
function markerText(nodes: Node[]): string {
  let text = '';
  for (const node of nodes) text += node.textContent ?? '';
  return text.replace(/^[\s\u00a0]+/, '').replace(/[\s\u00a0]+$/, '');
}

/** Formatting declared on the marker span and its wrappers. */
function markerFormatting(nodes: Node[]): RunFormatting {
  let formatting: RunFormatting = {};
  const visit = (node: Node, depth: number): void => {
    if (!isElement(node) || depth > 8) return;
    const style = parseInlineStyle(attr(node, 'style'));
    // The innermost spacer span (`font:7.0pt "Times New Roman"`) is layout, so
    // its font must not be mistaken for the marker's font.
    const isSpacer = /^\s*\d/.test(style['font'] ?? '') && (node.textContent ?? '').trim() === '';
    if (!isSpacer && Object.keys(style).length > 0) {
      formatting = { ...formatting, ...parseRunFormattingFromCss(style) };
    }
    for (const child of childNodesOf(node)) visit(child, depth + 1);
  };
  for (const node of nodes) visit(node, 0);
  return formatting;
}

/**
 * Decide whether a rendered marker is a bullet or a number.
 *
 * The list definition answers this properly and is consulted first by the
 * caller; this is the fallback for payloads where the `@list` rule is absent
 * (a partial copy, Word Online, a payload that went through another editor).
 */
function looksLikeBullet(text: string, font: string | undefined): boolean {
  if (!text) return false;
  if (isSymbolFont(font)) return true;
  const codePoint = text.codePointAt(0) ?? 0;
  // Private use area: a symbol font byte, whatever the font attribute says.
  if (codePoint >= 0xf000 && codePoint <= 0xf0ff) return true;
  if (text.length > 2) return false;
  // Word's level-2 default bullet really is the letter "o" in Courier New.
  if (text === 'o' && /courier/i.test(font ?? '')) return true;
  if (/^[0-9a-z]/i.test(text)) return false;
  // A single non-alphanumeric character is a bullet glyph.
  return /^[^\w\s]$/u.test(text);
}

/**
 * Build the list item for a paragraph.
 *
 * `reference` comes from the `mso-list` declaration and `extracted` from the
 * rendered marker. Where both exist they corroborate each other: the
 * definition says how the list numbers, the rendered marker says what Word
 * actually drew. The definition wins on *format*, the rendered text wins on
 * *literal appearance*, and any disagreement is a diagnostic rather than a
 * silent choice.
 */
export function buildListItem(
  reference: WordListReference,
  extracted: ExtractedMarker | null,
  css: Record<string, string>,
  ctx: WordParseContext,
  element: Element,
): ListItemInfo {
  const definition = ctx.sheet.lists[reference.listId];
  const levelDefinition: WordListLevel | undefined = definition?.levels[reference.level];

  if (!definition) {
    ctx.diagnostics.warn(
      DiagnosticCode.WORD_LIST_DEFINITION_MISSING,
      `Paragraph references list "${reference.listId}" but the clipboard stylesheet contains no @list ${reference.listId} rule. The marker Word rendered is used as the authority instead.`,
      { details: { listId: reference.listId, level: reference.level + 1 }, location: { path: nodePath(element) } },
    );
  } else if (!levelDefinition) {
    ctx.diagnostics.warn(
      DiagnosticCode.WORD_LIST_LEVEL_MISSING,
      `Paragraph references ${reference.listId} level ${reference.level + 1}, which the @list rule does not define.`,
      { details: { listId: reference.listId, level: reference.level + 1 } },
    );
  }

  const marker = combineMarker(extracted?.marker, levelDefinition, ctx, element);

  const item: ListItemInfo = {
    listId: reference.listId,
    level: reference.level,
    marker,
  };
  if (reference.lfo) item.lfo = reference.lfo;
  if (levelDefinition) item.levelDefinition = levelDefinition;

  const marginLeft = parseWordLength(css['margin-left'], { defaultUnit: 'pt' });
  if (marginLeft) item.marginLeft = marginLeft;
  const textIndent = parseWordLength(css['text-indent'], { defaultUnit: 'pt' });
  if (textIndent) item.textIndent = textIndent;

  return item;
}

function combineMarker(
  rendered: ListItemMarker | undefined,
  definition: WordListLevel | undefined,
  ctx: WordParseContext,
  element: Element,
): ListItemMarker {
  if (!rendered && !definition) {
    ctx.diagnostics.warn(
      DiagnosticCode.WORD_LIST_MARKER_HEURISTIC,
      'A list paragraph carried neither a rendered marker nor a list level definition; it is modelled as an unnumbered list item.',
      { location: { path: nodePath(element) } },
    );
    return { type: 'none', source: 'text-heuristic' };
  }

  if (!rendered && definition) {
    // Word Online omits the marker spans. The definition alone is enough to
    // describe the list; the literal text is computed at render time from the
    // level text and the counter state.
    const marker: ListItemMarker = {
      type: definition.type,
      source: 'list-definition',
    };
    if (definition.bulletGlyph) marker.glyph = definition.bulletGlyph;
    if (definition.bulletGlyphRaw) marker.rawGlyph = definition.bulletGlyphRaw;
    if (definition.bulletFont) marker.font = definition.bulletFont;
    if (definition.bulletFontMapped) marker.fontMapped = true;
    if (definition.levelText) marker.levelText = definition.levelText;
    if (definition.numberFormat) marker.numberFormat = definition.numberFormat;
    if (definition.startAt !== undefined) marker.startAt = definition.startAt;
    if (definition.markerFormatting) marker.formatting = definition.markerFormatting;
    return marker;
  }

  const marker: ListItemMarker = { ...rendered! };

  if (definition) {
    // The definition is authoritative about *kind* and *format*.
    marker.type = definition.type;
    if (definition.numberFormat) marker.numberFormat = definition.numberFormat;
    if (definition.levelText) marker.levelText = definition.levelText;
    if (definition.startAt !== undefined) marker.startAt = definition.startAt;

    if (definition.type === 'bullet') {
      // Prefer the definition's glyph: `mso-level-text:\F0B7` states the font
      // byte exactly, where the rendered span holds whatever the ANSI code page
      // mapped it to.
      if (definition.bulletGlyph) marker.glyph = definition.bulletGlyph;
      if (definition.bulletGlyphRaw) marker.rawGlyph = definition.bulletGlyphRaw;
      if (definition.bulletFont) marker.font = definition.bulletFont;
      if (definition.bulletFontMapped) marker.fontMapped = true;
    } else if (marker.glyph) {
      // The heuristic guessed bullet, the definition says number. Trust the
      // definition and drop the glyph.
      delete marker.glyph;
      delete marker.rawGlyph;
    }

    if (definition.markerFormatting) {
      marker.formatting = { ...definition.markerFormatting, ...(marker.formatting ?? {}) };
    }
  } else if (marker.type === 'number' && marker.text) {
    // No definition: infer the number format from what Word drew, and say so.
    const inferred = inferNumberFormat(marker.text);
    if (inferred) {
      marker.numberFormat = inferred.format;
      marker.levelText = inferred.levelText;
    }
    ctx.diagnostics.warn(
      DiagnosticCode.WORD_LIST_MARKER_HEURISTIC,
      `No @list definition was available, so the numbering format for marker "${marker.text}" was inferred from the rendered text. Verify against the source document.`,
      { details: { marker: marker.text, inferred: inferred?.format ?? 'unknown' } },
    );
  }

  return marker;
}

/**
 * Infer a number format from a rendered marker such as `IV.`, `c)`, `1.1`.
 *
 * Used only when the `@list` definition is missing. The level text produced
 * keeps the separators Word drew, so `1.1` stays `%1.%2` rather than being
 * re-invented as some other scheme.
 */
export function inferNumberFormat(
  text: string,
): { format: ListItemMarker['numberFormat']; levelText: string } | null {
  const trimmed = text.trim();
  if (!trimmed) return null;

  // Multi-part numbers: 1.1, 1.1.1, 2.3.4.
  const multi = /^(\d+(?:\.\d+)+)([.)\]]?)$/.exec(trimmed);
  if (multi) {
    const parts = multi[1]!.split('.');
    const levelText = parts.map((_, i) => `%${i + 1}`).join('.') + (multi[2] ?? '');
    return { format: 'decimal', levelText };
  }

  const single = /^([(\[]?)([0-9]+|[IVXLCDM]+|[ivxlcdm]+|[A-Z]|[a-z])([.)\]]?)$/.exec(trimmed);
  if (!single) return null;
  const prefix = single[1] ?? '';
  const body = single[2]!;
  const suffix = single[3] ?? '';
  const levelText = `${prefix}%1${suffix}`;

  if (/^\d+$/.test(body)) {
    return { format: body.length > 1 && body.startsWith('0') ? 'decimal-leading-zero' : 'decimal', levelText };
  }
  if (/^[IVXLCDM]+$/.test(body)) {
    // A single `I` is ambiguous between upper-roman and upper-alpha. Roman is
    // the far more common Word outline default and matches `I.` headings.
    return { format: 'upper-roman', levelText };
  }
  if (/^[ivxlcdm]+$/.test(body)) return { format: 'lower-roman', levelText };
  if (/^[A-Z]$/.test(body)) return { format: 'upper-alpha', levelText };
  if (/^[a-z]$/.test(body)) return { format: 'lower-alpha', levelText };
  return null;
}

/**
 * Recover a marker from leading paragraph text.
 *
 * The last resort, for payloads where Word's structure is gone entirely (text
 * pasted through another application, a `.htm` file edited by hand). Disabled
 * by `recoverMarkersFromText: false`, and always diagnosed, because unlike
 * everything else in this module it *is* a guess.
 */
export function recoverMarkerFromText(
  text: string,
  ctx: WordParseContext,
): { marker: ListItemMarker; remainingText: string } | null {
  if (ctx.options.recoverMarkersFromText === false) return null;

  const bulletMatch = /^([\u2022\u25aa\u25cf\u25e6\u00b7\u2043\u2023\u25a0\u25cb\u2013-])[ \t\u00a0]+(.*)$/u.exec(text);
  if (bulletMatch) {
    ctx.diagnostics.warn(
      DiagnosticCode.WORD_LIST_MARKER_HEURISTIC,
      `A leading "${bulletMatch[1]}" was recovered as a bullet marker from paragraph text because the payload carried no Word list structure.`,
      { details: { glyph: bulletMatch[1]! } },
    );
    return {
      marker: { type: 'bullet', glyph: bulletMatch[1]!, rawGlyph: bulletMatch[1]!, source: 'text-heuristic' },
      remainingText: bulletMatch[2] ?? '',
    };
  }

  const numberMatch = /^((?:\d+\.)*\d+[.)]?|[IVXLCDM]+[.)]|[ivxlcdm]+[.)]|[A-Za-z][.)])[ \t\u00a0]+(.*)$/.exec(text);
  if (numberMatch) {
    const inferred = inferNumberFormat(numberMatch[1]!);
    ctx.diagnostics.warn(
      DiagnosticCode.WORD_LIST_MARKER_HEURISTIC,
      `A leading "${numberMatch[1]}" was recovered as a number marker from paragraph text because the payload carried no Word list structure.`,
      { details: { marker: numberMatch[1]! } },
    );
    const marker: ListItemMarker = { type: 'number', text: numberMatch[1]!, source: 'text-heuristic' };
    if (inferred) {
      marker.numberFormat = inferred.format;
      marker.levelText = inferred.levelText;
    }
    return { marker, remainingText: numberMatch[2] ?? '' };
  }

  return null;
}

/** True when an element is a real HTML list container. */
export function isHtmlListElement(element: Element): boolean {
  const tag = tagNameOf(element);
  return tag === 'ol' || tag === 'ul';
}
