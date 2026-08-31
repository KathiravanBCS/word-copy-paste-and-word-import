import type { WordDocument } from '../model/Document.js';
import type { WordBlock, WordContainer, WordUnsupportedObject } from '../model/Block.js';
import type { WordParagraph } from '../model/Paragraph.js';
import type { WordRun } from '../model/Run.js';
import type { RunFormatting } from '../model/Style.js';
import type { WordHyperlink } from '../model/Hyperlink.js';
import { buildListTree, type ListTreeItem, type ListTreeNode } from '../normalization/NormalizeLists.js';
import { escapeHtmlAttribute, escapeHtmlText } from '../util/dom.js';
import {
  joinStyles,
  renderParagraphStyle,
  renderRunStyle,
  type StyleRenderOptions,
} from './HtmlStyleRenderer.js';
import {
  CssRegistry,
  compileLevelStyle,
  indentationFor,
  listTagFor,
  renderListItemAttributes,
  renderMarkerElement,
  type ListMarkerMode,
  type ListRenderOptions,
} from './HtmlListRenderer.js';
import { renderTable, tableCss, type TableRenderOptions } from './HtmlTableRenderer.js';
import { imagePlaceholderCss, renderImage, type ImageRenderOptions } from './HtmlImageRenderer.js';
import type { ListIndentation } from '../normalization/NormalizeUnits.js';

/**
 * The renderer.
 *
 * It reads the model and nothing else. It does not consult the raw payload, it
 * does not re-parse anything, and — the point of the whole architecture — it
 * does not reinterpret Word semantics. If the model says the marker is `I.`,
 * the output says `I.`. If the model says the bullet glyph is `§` in
 * Wingdings, the output draws `§` in Wingdings. The renderer has no numbering
 * scheme of its own to impose.
 *
 * Output is built by string construction from escaped model values rather than
 * by cloning input nodes, which is also what makes it safe: an attribute in
 * the clipboard payload has no path into the output unless the model has a
 * field for it.
 */

export interface RenderOptions extends StyleRenderOptions, ListRenderOptions, ImageRenderOptions, TableRenderOptions {
  /** How list markers are produced. Default `native`. */
  markerMode?: ListMarkerMode;
  /** CSS class prefix for generated classes. Default `wce`. */
  classPrefix?: string;
  /**
   * Where the generated CSS goes.
   *  `separate`  — returned in `css`, not embedded (default; the caller places it)
   *  `inline`    — wrapped in a `<style>` element at the top of the fragment
   *  `none`      — not generated; lists fall back to `element` marker mode
   */
  cssMode?: 'separate' | 'inline' | 'none';
  /** Emit `data-word-*` attributes carrying the original Word values. */
  includeWordMetadata?: boolean;
  /** Wrap the output in a container element with the class prefix. Default true. */
  wrapper?: boolean;
  /** Pretty-print with newlines between blocks. Default false. */
  pretty?: boolean;
}

export interface RenderResult {
  /** The rendered fragment. */
  html: string;
  /** The generated stylesheet (counter styles, list geometry, defaults). */
  css: string;
}

interface RenderContext {
  document: WordDocument;
  options: Required<Pick<RenderOptions, 'classPrefix' | 'markerMode'>> & RenderOptions;
  css: CssRegistry;
  /** Indentation of the list currently being rendered, for relative nesting. */
  listStack: ListIndentation[];
}

export function renderWordDocument(document: WordDocument, options: RenderOptions = {}): RenderResult {
  const prefix = options.classPrefix ?? 'wce';
  const cssMode = options.cssMode ?? 'separate';
  const markerMode: ListMarkerMode = cssMode === 'none' ? 'element' : (options.markerMode ?? 'native');

  const ctx: RenderContext = {
    document,
    options: { ...options, classPrefix: prefix, markerMode },
    css: new CssRegistry(prefix),
    listStack: [],
  };

  const body = renderBlocks(document.blocks, ctx);

  const css =
    cssMode === 'none'
      ? ''
      : [baseCss(prefix), imagePlaceholderCss(prefix), tableCss(prefix), ctx.css.toCss()]
          .filter(Boolean)
          .join('\n');

  const wrapped =
    options.wrapper === false
      ? body
      : `<div class="${prefix}-document">${body}</div>`;

  const html = cssMode === 'inline' && css ? `<style>${css}</style>${wrapped}` : wrapped;
  return { html, css };
}

/** Convenience: the fragment only, with CSS inlined. */
export function renderWordDocumentToHtml(document: WordDocument, options: RenderOptions = {}): string {
  return renderWordDocument(document, { ...options, cssMode: options.cssMode ?? 'inline' }).html;
}

/* -------------------------------------------------------------------------
 * Blocks
 * ---------------------------------------------------------------------- */

function renderBlocks(blocks: WordBlock[], ctx: RenderContext): string {
  const parts: string[] = [];
  let index = 0;

  while (index < blocks.length) {
    const block = blocks[index]!;

    // A run of list paragraphs is rendered as one nested list structure.
    if (block.type === 'paragraph' && block.listItem) {
      const tree = buildListTree(blocks, index);
      if (tree) {
        parts.push(renderList(tree.node, ctx, null));
        index = tree.nextIndex;
        continue;
      }
    }

    parts.push(renderBlock(block, ctx));
    index++;
  }
  return parts.join(ctx.options.pretty ? '\n' : '');
}

function renderBlock(block: WordBlock, ctx: RenderContext): string {
  switch (block.type) {
    case 'paragraph':
      return renderParagraph(block, ctx);
    case 'table':
      return renderTable(block, { renderBlocks: (blocks) => renderBlocks(blocks, ctx) }, ctx.options);
    case 'image-block':
      return renderImageBlock(block.imageId, ctx);
    case 'page-break':
      return renderPageBreak(block.breakType, ctx);
    case 'horizontal-rule':
      return `<hr class="${ctx.options.classPrefix}-rule">`;
    case 'container':
      return renderContainer(block, ctx);
    case 'unsupported':
      return renderUnsupported(block, ctx);
    default:
      return '';
  }
}

function renderParagraph(paragraph: WordParagraph, ctx: RenderContext): string {
  const tag = paragraph.headingLevel ? `h${paragraph.headingLevel}` : 'p';
  const style = renderParagraphStyle(paragraph.formatting, ctx.options);
  const attributes = paragraphAttributes(paragraph, style, ctx);
  const content = renderRuns(paragraph.runs, ctx);

  // Word's deliberately empty paragraphs are blank lines in the document; a
  // truly empty <p> collapses, so it keeps a non-breaking space.
  const body = content || (paragraph.empty ? '&nbsp;' : '');
  return `<${tag}${attributes}>${body}</${tag}>`;
}

function paragraphAttributes(paragraph: WordParagraph, style: string, ctx: RenderContext): string {
  const parts: string[] = [];
  const classes = [`${ctx.options.classPrefix}-p`];
  if (paragraph.styleId) classes.push(`${ctx.options.classPrefix}-style-${sanitiseClass(paragraph.styleId)}`);
  parts.push(` class="${escapeHtmlAttribute(classes.join(' '))}"`);
  if (style) parts.push(` style="${escapeHtmlAttribute(style)}"`);
  if (ctx.options.includeWordMetadata && paragraph.styleName) {
    parts.push(` data-word-style="${escapeHtmlAttribute(paragraph.styleName)}"`);
  }
  if (paragraph.bookmarks?.length) {
    parts.push(` id="${escapeHtmlAttribute(paragraph.bookmarks[0]!)}"`);
  }
  return parts.join('');
}

function renderContainer(container: WordContainer, ctx: RenderContext): string {
  const tag = container.role === 'blockquote' ? 'blockquote' : 'div';
  const style = renderParagraphStyle(container.formatting, ctx.options);
  const classes = [`${ctx.options.classPrefix}-${container.role === 'text-box' ? 'textbox' : 'block'}`];
  const attributes = [
    ` class="${escapeHtmlAttribute(classes.join(' '))}"`,
    style ? ` style="${escapeHtmlAttribute(style)}"` : '',
  ].join('');
  return `<${tag}${attributes}>${renderBlocks(container.blocks, ctx)}</${tag}>`;
}

function renderPageBreak(breakType: string, ctx: RenderContext): string {
  const prefix = ctx.options.classPrefix;
  ctx.css.add(
    `.${prefix}-page-break { break-before: page; page-break-before: always; height: 0; margin: 0; border: 0; }`,
  );
  return `<div class="${prefix}-page-break" role="separator" aria-label="${
    breakType === 'section' ? 'Section break' : 'Page break'
  }" data-word-break="${escapeHtmlAttribute(breakType)}"></div>`;
}

function renderImageBlock(imageId: string, ctx: RenderContext): string {
  const image = ctx.document.images[imageId];
  if (!image) return '';
  return `<p class="${ctx.options.classPrefix}-p">${renderImage(image, ctx.options)}</p>`;
}

/**
 * An unsupported object is rendered as an inert, visible placeholder.
 *
 * The alternatives are worse: dropping it makes content vanish with no trace,
 * and emitting Word's original markup would put VML or an `<object>` into the
 * page — which is both broken and a security decision this engine does not
 * make on the caller's behalf.
 */
function renderUnsupported(block: WordUnsupportedObject, ctx: RenderContext): string {
  const prefix = ctx.options.classPrefix;
  ctx.css.add(
    `.${prefix}-unsupported { display: block; padding: 6px 10px; border: 1px dashed currentColor; ` +
      `border-radius: 4px; opacity: 0.65; font-size: 12px; }`,
  );
  const label = describeObjectType(block.objectType);
  const fallback = block.fallbackText ? escapeHtmlText(block.fallbackText) : '';
  return (
    `<div class="${prefix}-unsupported" data-word-object="${escapeHtmlAttribute(block.objectType)}" ` +
    `data-word-code="${escapeHtmlAttribute(block.code)}" role="note">` +
    `${escapeHtmlText(label)}${fallback ? `: ${fallback}` : ''}</div>`
  );
}

function describeObjectType(objectType: string): string {
  switch (objectType) {
    case 'ole-object':
      return 'Embedded object (not representable in HTML)';
    case 'vml-shape':
      return 'Word drawing (not representable in HTML)';
    default:
      return `Unsupported Word object: ${objectType}`;
  }
}

/* -------------------------------------------------------------------------
 * Lists
 * ---------------------------------------------------------------------- */

function renderList(node: ListTreeNode, ctx: RenderContext, parent: ListIndentation | null): string {
  const first = node.items.find((item) => item.item);
  if (!first?.item || !first.block) {
    // A synthesised level with nothing but a deeper list under it.
    return node.items.map((item) => renderListItem(item, ctx, null, null)).join('');
  }

  const indentation = indentationFor(first.item, first.block.formatting);
  const level = compileLevelStyle(node, first.item, indentation, parent, ctx.css, ctx.options);

  const tag = listTagFor(node);
  const attributes: string[] = [
    ` class="${escapeHtmlAttribute(`${ctx.options.classPrefix}-list ${level.listClass}`)}"`,
  ];
  if (node.ordered && node.startAt !== undefined && node.startAt !== 1) {
    attributes.push(` start="${node.startAt}"`);
  }
  if (ctx.options.includeWordMetadata) {
    attributes.push(` data-word-list="${escapeHtmlAttribute(node.listId)}"`);
    attributes.push(` data-word-level="${node.level + 1}"`);
    attributes.push(` data-word-instance="${escapeHtmlAttribute(node.listInstanceId)}"`);
  }

  ctx.listStack.push(indentation);
  const items = node.items
    .map((item) => renderListItem(item, ctx, level.strategy, indentation))
    .join('');
  ctx.listStack.pop();

  return `<${tag}${attributes.join('')}>${items}</${tag}>`;
}

function renderListItem(
  item: ListTreeItem,
  ctx: RenderContext,
  strategy: ReturnType<typeof compileLevelStyle>['strategy'] | null,
  indentation: ListIndentation | null,
): string {
  const children = item.children
    .map((child) => renderList(child, ctx, indentation))
    .join('');

  if (item.spacer || !item.item || !item.block) {
    // The empty rung of a level jump. It carries no marker and no content, so
    // the nesting depth Word declared is preserved without inventing an item.
    return `<li class="${ctx.options.classPrefix}-list-spacer" style="list-style:none">${children}</li>`;
  }

  const attributes = strategy
    ? renderListItemAttributes(item.item, strategy, ctx.options)
    : '';
  const marker =
    strategy?.kind === 'element' ? renderMarkerElement(item.item, ctx.options) : '';

  const paragraph = item.block;
  const style = renderListItemStyle(paragraph, ctx);
  const content = renderRuns(paragraph.runs, ctx);

  return (
    `<li class="${ctx.options.classPrefix}-list-item"${attributes}` +
    `${style ? ` style="${escapeHtmlAttribute(style)}"` : ''}>` +
    `${marker}${content}${children}</li>`
  );
}

/**
 * Paragraph formatting for a list item, minus the indentation.
 *
 * The indentation has already been turned into the list's own geometry;
 * re-emitting the paragraph's `margin-left` on the `<li>` would apply it twice.
 */
function renderListItemStyle(paragraph: WordParagraph, ctx: RenderContext): string {
  const formatting = { ...paragraph.formatting };
  delete formatting.marginLeft;
  delete formatting.textIndent;
  return renderParagraphStyle(formatting, ctx.options);
}

/* -------------------------------------------------------------------------
 * Runs
 * ---------------------------------------------------------------------- */

function renderRuns(runs: WordRun[], ctx: RenderContext): string {
  const parts: string[] = [];
  let openLink: WordHyperlink | null = null;

  const closeLink = (): void => {
    if (openLink) {
      parts.push('</a>');
      openLink = null;
    }
  };

  for (const run of runs) {
    const linkId = 'hyperlinkId' in run ? run.hyperlinkId : undefined;
    const link = linkId ? ctx.document.hyperlinks[linkId] : undefined;

    if (link !== openLink) {
      closeLink();
      if (link) {
        parts.push(renderLinkOpen(link, ctx));
        openLink = link;
      }
    }
    parts.push(renderRun(run, ctx));
  }
  closeLink();
  return parts.join('');
}

function renderLinkOpen(link: WordHyperlink, ctx: RenderContext): string {
  const attributes: string[] = [` class="${ctx.options.classPrefix}-link"`];
  if (link.blocked || !link.href) {
    // The target could not be emitted. A `<span>` would lose the fact that
    // this text was a link at all, so the element stays and carries the
    // original target as data for anyone who can do something with it.
    attributes.push(` data-word-unresolved-href="${escapeHtmlAttribute(link.rawHref)}"`);
    attributes.push(` role="link"`);
    attributes.push(` aria-disabled="true"`);
  } else {
    attributes.push(` href="${escapeHtmlAttribute(link.href)}"`);
    if (link.target) attributes.push(` target="${escapeHtmlAttribute(link.target)}"`);
    if (link.target === '_blank') attributes.push(' rel="noopener noreferrer"');
  }
  if (link.title) attributes.push(` title="${escapeHtmlAttribute(link.title)}"`);
  return `<a${attributes.join('')}>`;
}

function renderRun(run: WordRun, ctx: RenderContext): string {
  switch (run.type) {
    case 'text':
      return renderTextRun(run.text, run.formatting, run.bookmarks, ctx);
    case 'break':
      return run.breakType === 'page' ? renderPageBreak('page', ctx) : '<br>';
    case 'tab':
      return renderTab(ctx);
    case 'image': {
      const image = ctx.document.images[run.imageId];
      return image ? renderImage(image, ctx.options) : '';
    }
    case 'field':
      return renderTextRun(run.text, run.formatting, undefined, ctx);
    case 'note':
      return renderTextRun(run.text, run.formatting, undefined, ctx);
    default:
      return '';
  }
}

function renderTextRun(
  text: string,
  formatting: RunFormatting,
  bookmarks: string[] | undefined,
  ctx: RenderContext,
): string {
  if (text.length === 0) return '';
  let inner = escapeHtmlText(text);

  // Prefer semantic elements over a styled span: <strong> and <em> carry
  // meaning that a font-weight declaration does not, and they survive being
  // pasted onward into an editor that strips inline CSS. Whatever has no
  // element form is left on a span.
  const semantic: string[] = [];
  const residual: RunFormatting = { ...formatting };

  if (formatting.bold) {
    semantic.push('strong');
    delete residual.bold;
  }
  if (formatting.italic) {
    semantic.push('em');
    delete residual.italic;
  }
  if (formatting.underline === 'single') {
    semantic.push('u');
    delete residual.underline;
  }
  if (formatting.strike && !formatting.doubleStrike) {
    semantic.push('s');
    delete residual.strike;
  }
  if (formatting.verticalAlign === 'super') {
    semantic.push('sup');
    delete residual.verticalAlign;
  } else if (formatting.verticalAlign === 'sub') {
    semantic.push('sub');
    delete residual.verticalAlign;
  }

  const residualStyle = renderRunStyle(residual, ctx.options);
  for (const tag of semantic.reverse()) inner = `<${tag}>${inner}</${tag}>`;

  const anchor = bookmarks?.length ? ` id="${escapeHtmlAttribute(bookmarks[0]!)}"` : '';
  if (!residualStyle && !anchor) return inner;
  return `<span${anchor}${residualStyle ? ` style="${escapeHtmlAttribute(residualStyle)}"` : ''}>${inner}</span>`;
}

function renderTab(ctx: RenderContext): string {
  const prefix = ctx.options.classPrefix;
  ctx.css.add(`.${prefix}-tab { display: inline-block; min-width: 36pt; }`);
  return `<span class="${prefix}-tab"></span>`;
}

/* -------------------------------------------------------------------------
 * Base stylesheet
 * ---------------------------------------------------------------------- */

function baseCss(prefix: string): string {
  return [
    `.${prefix}-document { }`,
    `.${prefix}-p { margin: 0; }`,
    `.${prefix}-list { }`,
    `.${prefix}-list-item { }`,
    `.${prefix}-list-spacer { }`,
  ].join('\n');
}

function sanitiseClass(value: string): string {
  return value.replace(/[^a-zA-Z0-9_-]/g, '-').toLowerCase();
}

/** Combine style strings; exported for adapters that extend the renderer. */
export { joinStyles };
