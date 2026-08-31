import type { WordImage } from '../model/Image.js';
import { escapeHtmlAttribute, escapeHtmlText } from '../util/dom.js';
import { StyleBuilder, renderLength, type StyleRenderOptions } from './HtmlStyleRenderer.js';
import { roundTo } from '../word/WordLengthParser.js';

/**
 * Image rendering.
 *
 * The rule that matters: **an unresolved image never becomes a broken
 * `<img>`.** Word's `file:///…/clip_image001.png` references resolve on
 * exactly one machine in the world, and emitting them produces content that
 * looks correct to whoever pasted it (their browser may even have the file
 * cached) and is broken for everyone else. So an unresolved image is rendered
 * as a visible, labelled placeholder that carries the original reference, and
 * the fidelity report counts it.
 */

export interface ImageRenderOptions extends StyleRenderOptions {
  /** CSS class prefix. Default `wce`. */
  classPrefix?: string;
  /**
   * Render unresolved images as a visible placeholder. Default true.
   * When false they are omitted entirely — still never as a broken `<img>`.
   */
  placeholderForUnresolved?: boolean;
}

export function renderImage(image: WordImage, options: ImageRenderOptions = {}): string {
  const prefix = options.classPrefix ?? 'wce';

  if (image.resolution === 'unresolved') {
    if (options.placeholderForUnresolved === false) return '';
    return renderPlaceholder(image, prefix, options);
  }

  const style = new StyleBuilder();
  if (image.width) style.set('width', renderLength(image.width, options));
  if (image.height) style.set('height', renderLength(image.height, options));
  if (image.placement === 'floating') {
    style.set('float', 'left');
    style.set('margin', '0 8pt 4pt 0');
  }
  applyCrop(image, style);

  const attributes: string[] = [
    ` src="${escapeHtmlAttribute(image.src)}"`,
    ` alt="${escapeHtmlAttribute(image.alt ?? '')}"`,
  ];
  if (image.title) attributes.push(` title="${escapeHtmlAttribute(image.title)}"`);
  if (!style.isEmpty) attributes.push(` style="${escapeHtmlAttribute(style.toString())}"`);
  if (options.includeWordMetadata) {
    attributes.push(` data-word-source="${escapeHtmlAttribute(image.originalSource)}"`);
    attributes.push(` data-word-origin="${escapeHtmlAttribute(image.origin)}"`);
    if (image.shapeId) attributes.push(` data-word-shape="${escapeHtmlAttribute(image.shapeId)}"`);
  }
  return `<img class="${prefix}-image"${attributes.join('')}>`;
}

/**
 * A crop is expressed by Word as percentages taken off each edge of the source
 * picture. CSS has no crop property, so the picture is oversized inside a
 * clipping wrapper — an approximation, and diagnosed as one at parse time.
 */
function applyCrop(image: WordImage, style: StyleBuilder): void {
  const crop = image.crop;
  if (!crop) return;
  const top = crop.top ?? 0;
  const right = crop.right ?? 0;
  const bottom = crop.bottom ?? 0;
  const left = crop.left ?? 0;
  if (top === 0 && right === 0 && bottom === 0 && left === 0) return;
  style.set(
    'clip-path',
    `inset(${roundTo(top, 3)}% ${roundTo(right, 3)}% ${roundTo(bottom, 3)}% ${roundTo(left, 3)}%)`,
  );
}

function renderPlaceholder(
  image: WordImage,
  prefix: string,
  options: ImageRenderOptions,
): string {
  const style = new StyleBuilder();
  if (image.width) style.set('width', renderLength(image.width, options));
  if (image.height) style.set('height', renderLength(image.height, options));

  const label = image.alt || describeReference(image);
  const attributes: string[] = [
    ` class="${prefix}-image-placeholder"`,
    ` role="img"`,
    ` aria-label="${escapeHtmlAttribute(`Unresolved image: ${label}`)}"`,
    ` title="${escapeHtmlAttribute(`This image could not be recovered from the clipboard. Original reference: ${image.originalSource || '(none)'}`)}"`,
  ];
  if (!style.isEmpty) attributes.push(` style="${escapeHtmlAttribute(style.toString())}"`);
  if (options.includeWordMetadata !== false) {
    attributes.push(` data-word-source="${escapeHtmlAttribute(image.originalSource)}"`);
    attributes.push(` data-word-unresolved="true"`);
  }
  return `<span${attributes.join('')}><span class="${prefix}-image-placeholder-label">${escapeHtmlText(label)}</span></span>`;
}

function describeReference(image: WordImage): string {
  const source = image.originalSource;
  if (!source) return 'image';
  const fileName = /([^/\\?#]+)(?:[?#].*)?$/.exec(source)?.[1];
  return fileName ? decodeURIComponent(fileName) : source.slice(0, 60);
}

/** Default styling for placeholders, added to the generated stylesheet. */
export function imagePlaceholderCss(prefix = 'wce'): string {
  return [
    `.${prefix}-image { max-width: 100%; height: auto; }`,
    `.${prefix}-image-placeholder {`,
    `  display: inline-flex;`,
    `  align-items: center;`,
    `  justify-content: center;`,
    `  box-sizing: border-box;`,
    `  max-width: 100%;`,
    `  padding: 8px;`,
    `  border: 1px dashed currentColor;`,
    `  border-radius: 4px;`,
    `  opacity: 0.65;`,
    `  font-size: 11px;`,
    `  line-height: 1.3;`,
    `  text-align: center;`,
    `  overflow: hidden;`,
    `  vertical-align: middle;`,
    `}`,
    `.${prefix}-image-placeholder-label { display: block; word-break: break-word; }`,
  ].join('\n');
}
