import type { ImageCrop, WordImage } from '../model/Image.js';
import { attr, excerpt, tagNameOf } from '../util/dom.js';
import { checkImageUrl } from '../util/security.js';
import { DiagnosticCode } from '../diagnostics/UnsupportedFeature.js';
import { parseWordLength } from './WordLengthParser.js';
import { parseInlineStyle } from './WordCssTokenizer.js';
import { nextImageId, type VmlShape, type WordParseContext } from './WordParseContext.js';

/**
 * Word clipboard images.
 *
 * What Word actually puts in the HTML flavour is rarely a usable image:
 *
 *   `<img width=451 height=185 src="file:///C:/Users/x/AppData/Local/Temp/
 *    msohtmlclip1/01/clip_image001.png" v:shapes="Picture_x0020_1">`
 *
 * That path exists on the author's machine and nowhere else. A browser will
 * silently render a broken image, which is the single most common way a
 * paste-from-Word implementation looks like it works and does not.
 *
 * So the rules here are:
 *
 *   - A `file:///` or `cid:` reference is *never* emitted as an `<img src>`.
 *     It becomes an `unresolved` image with a diagnostic.
 *   - Before giving up, the parser looks for actual image bytes in the
 *     clipboard payload. Word supplies these as separate clipboard items on
 *     several platforms, and pairing them recovers the picture completely.
 *   - The VML twin of the image (`<v:shape><v:imagedata>`) is consulted for
 *     size and crop, which the `<img>` element does not carry.
 */

export interface ParseImageOptions {
  /** Placement, when the caller already knows (a floated shape, say). */
  placement?: WordImage['placement'];
}

/** Parse an `<img>` element into the document's image table. */
export function parseImageElement(
  element: Element,
  ctx: WordParseContext,
  options: ParseImageOptions = {},
): WordImage {
  const rawSrc = attr(element, 'src') ?? '';
  const style = parseInlineStyle(attr(element, 'style'));
  const shapeId = attr(element, 'v:shapes') ?? attr(element, 'o:spid');
  const shape = shapeId ? findShape(ctx, shapeId) : undefined;

  const id = nextImageId(ctx);
  const image: WordImage = {
    id,
    src: '',
    originalSource: rawSrc,
    resolution: 'unresolved',
    origin: 'none',
    placement: options.placement ?? placementFromStyle(style),
  };
  if (shapeId) image.shapeId = shapeId;
  if (shape) {
    shape.consumed = true;
    image.vml = shape.raw;
    const crop = parseCrop(shape.crop);
    if (crop) {
      image.crop = crop;
      ctx.diagnostics.warn(
        DiagnosticCode.WORD_IMAGE_CROP_APPROXIMATED,
        'Word cropped this image inside the document. HTML has no crop primitive, so the crop percentages are preserved on the model and applied with a CSS clip on render.',
        { details: { imageId: id } },
      );
    }
  }

  applyDimensions(image, element, style, shape);

  const alt = attr(element, 'alt');
  if (alt) image.alt = alt;
  const title = attr(element, 'title') ?? shape?.title;
  if (title) image.title = title;

  resolveSource(image, rawSrc || shape?.imageSrc || '', ctx);

  ctx.images[id] = image;
  return image;
}

/** Build an image from a bare VML shape that had no `<img>` twin. */
export function parseVmlShapeImage(shape: VmlShape, ctx: WordParseContext): WordImage | null {
  if (!shape.imageSrc) return null;
  const id = nextImageId(ctx);
  const image: WordImage = {
    id,
    src: '',
    originalSource: shape.imageSrc,
    resolution: 'unresolved',
    origin: 'none',
    placement: 'inline',
    shapeId: shape.id,
    vml: shape.raw,
  };
  if (shape.title) image.title = shape.title;
  const crop = parseCrop(shape.crop);
  if (crop) image.crop = crop;
  applyDimensions(image, null, shape.style ? parseInlineStyle(shape.style) : {}, shape);
  resolveSource(image, shape.imageSrc, ctx);
  shape.consumed = true;
  ctx.images[id] = image;
  return image;
}

function findShape(ctx: WordParseContext, shapeId: string): VmlShape | undefined {
  const direct = ctx.vmlShapes.get(shapeId);
  if (direct) return direct;
  // Word escapes spaces in shape ids as `_x0020_`; the `v:shapes` attribute and
  // the `<v:shape id>` do not always agree on the escaping.
  const decoded = shapeId.replace(/_x0020_/g, ' ');
  for (const [key, shape] of ctx.vmlShapes) {
    if (key.replace(/_x0020_/g, ' ') === decoded) return shape;
  }
  return undefined;
}

function placementFromStyle(style: Record<string, string>): WordImage['placement'] {
  const position = (style['position'] ?? '').toLowerCase();
  const float = (style['float'] ?? style['mso-position-horizontal'] ?? '').toLowerCase();
  if (position === 'absolute' || position === 'relative') return 'floating';
  if (float === 'left' || float === 'right') return 'floating';
  return 'inline';
}

function applyDimensions(
  image: WordImage,
  element: Element | null,
  style: Record<string, string>,
  shape?: VmlShape,
): void {
  const shapeStyle = shape?.style ? parseInlineStyle(shape.style) : {};

  // Priority: CSS on the img, then the HTML width/height attributes (which are
  // pixels), then the VML shape's own style (which is in points).
  const width =
    parseWordLength(style['width'], { defaultUnit: 'pt' }) ??
    parseWordLength(element ? attr(element, 'width') : undefined, { defaultUnit: 'px' }) ??
    parseWordLength(shapeStyle['width'], { defaultUnit: 'pt' });
  const height =
    parseWordLength(style['height'], { defaultUnit: 'pt' }) ??
    parseWordLength(element ? attr(element, 'height') : undefined, { defaultUnit: 'px' }) ??
    parseWordLength(shapeStyle['height'], { defaultUnit: 'pt' });

  if (width) image.width = width;
  if (height) image.height = height;

  const naturalWidth = parseWordLength(style['mso-width-source'], { defaultUnit: 'px' });
  if (naturalWidth) image.naturalWidth = naturalWidth;
}

function parseCrop(crop: VmlShape['crop']): ImageCrop | undefined {
  if (!crop) return undefined;
  const result: ImageCrop = {};
  const convert = (value: string | undefined): number | undefined => {
    if (value === undefined) return undefined;
    // VML crops are either a percentage (`12.5%`) or a fraction expressed in
    // 1/65536ths (`8192f`).
    const percentMatch = /^([-\d.]+)%$/.exec(value.trim());
    if (percentMatch) {
      const n = Number.parseFloat(percentMatch[1]!);
      return Number.isFinite(n) ? n : undefined;
    }
    const fixedMatch = /^([-\d.]+)f$/i.exec(value.trim());
    if (fixedMatch) {
      const n = Number.parseFloat(fixedMatch[1]!);
      return Number.isFinite(n) ? (n / 65536) * 100 : undefined;
    }
    const plain = Number.parseFloat(value);
    return Number.isFinite(plain) ? plain * 100 : undefined;
  };
  const top = convert(crop.top);
  const right = convert(crop.right);
  const bottom = convert(crop.bottom);
  const left = convert(crop.left);
  if (top) result.top = top;
  if (right) result.right = right;
  if (bottom) result.bottom = bottom;
  if (left) result.left = left;
  return Object.keys(result).length > 0 ? result : undefined;
}

/**
 * Work out whether the image can actually be shown, and say so honestly.
 */
function resolveSource(image: WordImage, rawSrc: string, ctx: WordParseContext): void {
  image.originalSource = rawSrc;
  if (!rawSrc) {
    claimClipboardImage(image, ctx, 'The <img> element carried no src.');
    return;
  }

  const check = checkImageUrl(rawSrc);
  const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(rawSrc)?.[1]?.toLowerCase();

  if (check.blocked) {
    ctx.diagnostics.warn(
      DiagnosticCode.SECURITY_URL_BLOCKED,
      `Image source was blocked: ${check.reason ?? 'unsupported scheme'}`,
      { details: { imageId: image.id } },
    );
    claimClipboardImage(image, ctx, 'The image source was blocked.');
    return;
  }

  if (scheme === 'data') {
    image.src = rawSrc;
    image.resolution = 'resolved';
    image.origin = 'data-uri';
    const mime = /^data:([^;,]+)/i.exec(rawSrc)?.[1];
    if (mime) image.mimeType = mime;
    image.byteLength = rawSrc.length;
    return;
  }

  if (scheme === 'http' || scheme === 'https' || scheme === 'blob') {
    image.src = rawSrc;
    image.resolution = 'external';
    image.origin = scheme === 'blob' ? 'clipboard-blob' : 'http';
    if (scheme !== 'blob') {
      ctx.diagnostics.info(
        DiagnosticCode.SECURITY_EXTERNAL_RESOURCE,
        'An image points at an external URL. Rendering it will make a network request to a third-party host.',
        { details: { imageId: image.id, host: hostOf(rawSrc) }, fidelity: 'EQUIVALENT' },
      );
    }
    return;
  }

  if (scheme === 'file') {
    ctx.diagnostics.warn(
      DiagnosticCode.WORD_LOCAL_FILE_IMAGE,
      'Word referenced an image by local file path. That path only exists on the machine the content was copied from, so the reference cannot be loaded by a browser.',
      { details: { imageId: image.id, source: rawSrc.slice(0, 200) } },
    );
    claimClipboardImage(image, ctx, 'The image was referenced by local file path.');
    return;
  }

  if (scheme === 'cid') {
    ctx.diagnostics.warn(
      DiagnosticCode.WORD_CID_IMAGE,
      'Word referenced an image by content id (cid:). The bytes live in a MIME part that the HTML clipboard flavour does not carry.',
      { details: { imageId: image.id, source: rawSrc.slice(0, 200) } },
    );
    claimClipboardImage(image, ctx, 'The image was referenced by content id.');
    return;
  }

  // A relative path — possible when Word HTML was saved to disk and reopened.
  image.src = rawSrc;
  image.resolution = 'external';
  image.origin = 'http';
}

/**
 * Pair an unresolvable reference with actual bytes from the clipboard.
 *
 * The pairing is positional: clipboard image items arrive in document order,
 * so the Nth unresolvable image takes the Nth unused blob. That is right in
 * practice and is not guaranteed, so every pairing is diagnosed as an
 * assumption rather than presented as a fact.
 */
function claimClipboardImage(image: WordImage, ctx: WordParseContext, reason: string): void {
  if (ctx.options.matchClipboardImages === false) {
    markUnresolved(image, ctx, reason);
    return;
  }
  const cursor = ctx.state.clipboardImageCursor;
  const candidate = ctx.clipboardImages[cursor];
  if (!candidate) {
    markUnresolved(image, ctx, reason);
    return;
  }

  const source = candidate.dataUri ?? candidate.objectUrl;
  if (!source) {
    markUnresolved(image, ctx, reason);
    return;
  }

  ctx.state.clipboardImageCursor = cursor + 1;
  image.src = source;
  image.resolution = 'resolved';
  image.origin = 'clipboard-blob';
  image.mimeType = candidate.mimeType;
  if (candidate.byteLength !== undefined) image.byteLength = candidate.byteLength;

  ctx.diagnostics.info(
    DiagnosticCode.WORD_UNRESOLVED_IMAGE,
    `${reason} Image bytes were recovered from clipboard item ${cursor + 1} by document order; verify the pairing if the document contains several pictures.`,
    { details: { imageId: image.id, clipboardIndex: cursor }, fidelity: 'EQUIVALENT' },
  );
}

function markUnresolved(image: WordImage, ctx: WordParseContext, reason: string): void {
  image.resolution = 'unresolved';
  image.src = '';
  ctx.diagnostics.warn(
    DiagnosticCode.WORD_UNRESOLVED_IMAGE,
    `${reason} No image bytes were present in the clipboard, so a labelled placeholder is rendered instead of a broken image. The original reference is preserved on the model.`,
    { details: { imageId: image.id, source: image.originalSource.slice(0, 200) } },
  );
}

function hostOf(url: string): string {
  const match = /^[a-z][a-z0-9+.-]*:\/\/([^/?#]+)/i.exec(url);
  return match?.[1] ?? '';
}

/**
 * Harvest VML shapes from raw markup.
 *
 * Word puts the VML twin of every picture inside a downlevel-hidden
 * conditional comment, where the DOM never parses it into elements. So the
 * shapes are scanned out of the comment text before anything else happens.
 */
export function collectVmlShapes(markup: string, ctx: WordParseContext): void {
  const pattern = /<v:(shape|rect|oval|group|line|roundrect|polyline|image)\b([^>]*)>([\s\S]*?)<\/v:\1>|<v:(shape|rect|oval|image)\b([^>]*)\/>/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(markup)) !== null) {
    const tagName = `v:${(match[1] ?? match[4] ?? 'shape').toLowerCase()}`;
    const attributes = match[2] ?? match[5] ?? '';
    const body = match[3] ?? '';
    const raw = match[0];

    const id = readAttribute(attributes, 'id') ?? readAttribute(attributes, 'o:spid') ?? '';
    const shape: VmlShape = { id, tagName, raw: raw.slice(0, 4000) };

    const style = readAttribute(attributes, 'style');
    if (style) shape.style = style;

    const imagedata = /<v:imagedata\b([^>]*)\/?>/i.exec(body);
    if (imagedata) {
      const imageAttributes = imagedata[1] ?? '';
      const src = readAttribute(imageAttributes, 'src');
      if (src) shape.imageSrc = decodeAttributeEntities(src);
      const title = readAttribute(imageAttributes, 'o:title');
      if (title) shape.title = decodeAttributeEntities(title);
      const crop = {
        top: readAttribute(imageAttributes, 'croptop'),
        right: readAttribute(imageAttributes, 'cropright'),
        bottom: readAttribute(imageAttributes, 'cropbottom'),
        left: readAttribute(imageAttributes, 'cropleft'),
      };
      if (crop.top || crop.right || crop.bottom || crop.left) {
        shape.crop = {};
        if (crop.top) shape.crop.top = crop.top;
        if (crop.right) shape.crop.right = crop.right;
        if (crop.bottom) shape.crop.bottom = crop.bottom;
        if (crop.left) shape.crop.left = crop.left;
      }
    }

    if (id) ctx.vmlShapes.set(id, shape);
    else ctx.vmlShapes.set(`anon-${ctx.vmlShapes.size}`, shape);
  }
}

function readAttribute(attributes: string, name: string): string | undefined {
  const pattern = new RegExp(`\\b${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`, 'i');
  const match = pattern.exec(attributes);
  if (!match) return undefined;
  return match[1] ?? match[2] ?? match[3];
}

function decodeAttributeEntities(value: string): string {
  return value
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
}

/**
 * Record which shape ids an `<img v:shapes="…">` twin claims.
 *
 * Scanned from the raw payload rather than the DOM, for the same reason the
 * shapes themselves are: whether the VML is live elements or comment text
 * varies by host, and the raw text is the one representation that always has
 * both halves.
 */
export function collectClaimedShapeIds(markup: string, ctx: WordParseContext): void {
  const pattern = /\bv:shapes\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(markup)) !== null) {
    const value = match[1] ?? match[2] ?? match[3] ?? '';
    // The attribute can name several shapes, space separated.
    for (const id of value.split(/\s+/).filter(Boolean)) {
      ctx.claimedShapeIds.add(id);
      ctx.claimedShapeIds.add(id.replace(/_x0020_/g, ' '));
    }
  }
}

/** True when a live VML element is a definition or a sub-part, not a drawing. */
export function isVmlSubPart(tagName: string): boolean {
  return /^v:(shapetype|stroke|fill|formulas|f|path|handles|h|lock|textpath|shadow|imagedata|background)$/i.test(
    tagName,
  );
}

/** Report the VML shapes that no `<img>` claimed, so nothing vanishes silently. */
export function reportUnconsumedShapes(ctx: WordParseContext): VmlShape[] {
  const leftovers: VmlShape[] = [];
  for (const shape of ctx.vmlShapes.values()) {
    if (shape.consumed) continue;
    leftovers.push(shape);
  }
  return leftovers;
}

/** True when the element is a live (non-comment) VML element. */
export function isVmlElement(element: Element): boolean {
  return tagNameOf(element).startsWith('v:');
}

/** Build the raw-metadata bag for an unsupported drawing object. */
export function describeUnsupportedObject(element: Element): Record<string, string> {
  const metadata: Record<string, string> = { tagName: tagNameOf(element) };
  for (const attribute of Array.from(element.attributes)) {
    if (attribute.value.length <= 500) metadata[attribute.name] = attribute.value;
  }
  metadata.excerpt = excerpt(element, 400);
  return metadata;
}
