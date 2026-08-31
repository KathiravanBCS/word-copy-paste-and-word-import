import type { WordParagraph } from '../model/Paragraph.js';
import type { ParagraphFormatting, RunFormatting } from '../model/Style.js';
import type { WordRun } from '../model/Run.js';
import { attr, childNodesOf, classList, tagNameOf } from '../util/dom.js';
import { parseInlineStyle } from './WordCssTokenizer.js';
import {
  mergeParagraphFormatting,
  mergeRunFormatting,
  parseAlignment,
  parseParagraphFormattingFromCss,
  parseRunFormattingFromCss,
} from './WordFormattingParser.js';
import { humaniseClassName, normaliseStyleId, resolveStyleChain } from './WordStyleParser.js';
import {
  buildListItem,
  extractRenderedMarker,
  findListReference,
  recoverMarkerFromText,
} from './WordListParser.js';
import { createInlineContext, parseInlineNodes, runsAreEmpty, runsToText } from './WordRunParser.js';
import type { WordParseContext } from './WordParseContext.js';

/**
 * Paragraph parsing.
 *
 * Order matters here, and it is the order the whole architecture is built
 * around: **semantic extraction happens before cleanup**.
 *
 *   1. Resolve the paragraph's formatting from the clipboard stylesheet and
 *      its inline style.
 *   2. Decide whether it is a list item (the `mso-list` declaration).
 *   3. *Lift the rendered list marker out of the content* — before any run is
 *      created, so a bullet glyph can never end up inside a text node.
 *   4. Parse the remaining inline content into runs.
 *
 * Reversing steps 3 and 4 is exactly the mistake that produces `"• Item"` as
 * paragraph text.
 */

const HEADING_TAGS: Record<string, number> = { h1: 1, h2: 2, h3: 3, h4: 4, h5: 5, h6: 6 };

export interface ParagraphParseOptions {
  /** Formatting inherited from an enclosing container or table cell. */
  inheritedRun?: RunFormatting;
  inheritedParagraph?: ParagraphFormatting;
  /** Override the block role, e.g. when parsing an `<li>`. */
  forceListLevel?: number;
}

/**
 * Merge a paragraph's effective CSS: element defaults, then class-based Word
 * styles (through their `mso-style-parent` chain), then the inline style.
 */
export function resolveElementCss(
  element: Element,
  ctx: WordParseContext,
): { css: Record<string, string>; styleName?: string; styleId?: string } {
  const tag = tagNameOf(element);
  let css: Record<string, string> = {};
  let styleName: string | undefined;
  let styleId: string | undefined;

  const elementStyle = ctx.sheet.styles[`element:${tag}`];
  if (elementStyle) {
    css = { ...css, ...resolveStyleChain(ctx.sheet, elementStyle.id) };
  }

  for (const className of classList(element)) {
    // Word's contextual-spacing variants (`MsoListParagraphCxSpFirst`,
    // `…CxSpMiddle`, `…CxSpLast`) say where a paragraph sits in a run of
    // same-styled paragraphs, not that it is a different style. Word normally
    // emits a rule for each variant, but not always — and when it does not,
    // falling back to the base style is the difference between a list keeping
    // its half-inch indent and losing it entirely.
    const id = normaliseStyleId(className);
    const definition =
      ctx.sheet.styles[id] ?? ctx.sheet.styles[normaliseStyleId(baseStyleName(className))];
    if (!definition) continue;
    css = { ...css, ...resolveStyleChain(ctx.sheet, definition.id) };
    styleName = definition.name;
    styleId = definition.id;
  }

  // A class Word did not define a rule for is still the style's name — Word
  // omits rules for styles whose formatting matched the default.
  if (!styleName) {
    const firstClass = classList(element)[0];
    if (firstClass) {
      styleName = humaniseClassName(firstClass);
      styleId = normaliseStyleId(firstClass);
    }
  }

  const inline = parseInlineStyle(attr(element, 'style'));
  css = { ...css, ...inline };

  const result: { css: Record<string, string>; styleName?: string; styleId?: string } = { css };
  if (styleName) result.styleName = styleName;
  if (styleId) result.styleId = styleId;
  return result;
}

/** `MsoListParagraphCxSpFirst` -> `MsoListParagraph`. */
function baseStyleName(className: string): string {
  return className.replace(/CxSp(First|Middle|Last)$/i, '');
}

/** Parse a `<p>`, `<h1>`..`<h6>`, `<li>` or an implicit paragraph. */
export function parseParagraph(
  element: Element,
  ctx: WordParseContext,
  options: ParagraphParseOptions = {},
): WordParagraph {
  const tag = tagNameOf(element);
  const { css, styleName, styleId } = resolveElementCss(element, ctx);

  let formatting = parseParagraphFormattingFromCss(css);
  if (options.inheritedParagraph) {
    formatting = mergeParagraphFormatting(options.inheritedParagraph, formatting);
  }

  // Presentational attributes Word still emits on paragraphs and cells.
  const align = parseAlignment(attr(element, 'align'));
  if (align && !formatting.alignment) formatting.alignment = align;

  let runFormatting = parseRunFormattingFromCss(css);
  if (options.inheritedRun) {
    runFormatting = mergeRunFormatting(options.inheritedRun, runFormatting);
  }

  const paragraph: WordParagraph = { type: 'paragraph', runs: [], formatting };
  if (styleName) paragraph.styleName = styleName;
  if (styleId) paragraph.styleId = styleId;
  const classes = classList(element);
  if (classes.length > 0) paragraph.classNames = classes;

  const headingLevel = detectHeadingLevel(tag, styleName, classes, css);
  if (headingLevel) paragraph.headingLevel = headingLevel;

  // --- list handling, before any run is produced -------------------------
  const reference = findListReference(css) ?? (options.forceListLevel !== undefined ? null : null);
  const extracted = extractRenderedMarker(element, ctx);

  if (reference) {
    paragraph.listItem = buildListItem(reference, extracted, css, ctx, element);
  } else if (extracted) {
    // A marker span with no `mso-list` declaration: the paragraph is a list
    // item whose reference did not survive. Keep the marker; NormalizeLists
    // will group it with its neighbours.
    ctx.diagnostics.warn(
      'WORD_LIST_DEFINITION_MISSING',
      'A paragraph carried a rendered list marker but no mso-list declaration; it is modelled as a list item using the marker alone.',
      { location: { tagName: tag } },
    );
    paragraph.listItem = {
      listId: 'orphan',
      level: 0,
      marker: extracted.marker,
    };
  }

  // --- runs ---------------------------------------------------------------
  const inline = createInlineContext(runFormatting);
  paragraph.runs = parseInlineNodes(childNodesOf(element), ctx, inline);

  // Last-resort marker recovery, only for payloads with no Word list structure.
  if (!paragraph.listItem && paragraph.runs.length > 0) {
    maybeRecoverMarkerFromText(paragraph, ctx);
  }

  if (runsAreEmpty(paragraph.runs)) paragraph.empty = true;

  const bookmarks = collectBookmarks(paragraph.runs);
  if (bookmarks.length > 0) paragraph.bookmarks = bookmarks;

  return paragraph;
}

/**
 * Heading detection uses three independent signals and never a font size.
 *
 * A real `<h1>` element is definitive. So is a style whose Word name is
 * `Heading N`. `mso-outline-level` on its own is *not* — Word puts it on
 * outline-numbered list paragraphs too, and treating those as headings turns
 * a numbered list into a document outline.
 */
function detectHeadingLevel(
  tag: string,
  styleName: string | undefined,
  classes: string[],
  css: Record<string, string>,
): number | undefined {
  const fromTag = HEADING_TAGS[tag];
  if (fromTag) return fromTag;

  if (styleName) {
    const match = /^heading\s*([1-9])$/i.exec(styleName.trim());
    if (match) return Math.min(6, Number.parseInt(match[1]!, 10));
  }

  for (const className of classes) {
    const match = /^MsoHeading([1-9])$/i.exec(className);
    if (match) return Math.min(6, Number.parseInt(match[1]!, 10));
  }

  // An outline level combined with a heading-shaped style name corroborates.
  const outline = css['mso-outline-level'];
  if (outline && styleName && /heading|title/i.test(styleName)) {
    const level = Number.parseInt(outline, 10);
    if (Number.isFinite(level) && level >= 1 && level <= 9) return Math.min(6, level);
  }
  return undefined;
}

function maybeRecoverMarkerFromText(paragraph: WordParagraph, ctx: WordParseContext): void {
  const first = paragraph.runs[0];
  if (!first || first.type !== 'text') return;
  const text = runsToText(paragraph.runs);
  const recovered = recoverMarkerFromText(text, ctx);
  if (!recovered) return;

  // Trim exactly the marker characters off the first run, leaving every other
  // run and all formatting untouched.
  const consumed = text.length - recovered.remainingText.length;
  let remaining = consumed;
  for (const run of paragraph.runs) {
    if (remaining <= 0) break;
    if (run.type !== 'text') break;
    if (run.text.length <= remaining) {
      remaining -= run.text.length;
      run.text = '';
    } else {
      run.text = run.text.slice(remaining);
      remaining = 0;
    }
  }
  paragraph.runs = paragraph.runs.filter((run) => run.type !== 'text' || run.text.length > 0);
  paragraph.listItem = {
    listId: 'text-recovered',
    level: 0,
    marker: recovered.marker,
  };
}

function collectBookmarks(runs: WordRun[]): string[] {
  const names: string[] = [];
  for (const run of runs) {
    if (run.type === 'text' && run.bookmarks) {
      for (const name of run.bookmarks) if (!names.includes(name)) names.push(name);
    }
  }
  return names;
}
