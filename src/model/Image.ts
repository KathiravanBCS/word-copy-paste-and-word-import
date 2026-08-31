import type { Length } from './Style.js';

/** How an image relates to the text flow. */
export type ImagePlacement = 'inline' | 'floating' | 'anchored';

/** Percentage crop, as Word expresses it on `v:imagedata`. */
export interface ImageCrop {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
}

/**
 * How completely the image survived the clipboard round trip.
 *
 * `resolved`   — a usable, renderable source exists (data URI or object URL).
 * `external`   — an http(s) URL the browser may or may not be able to load.
 * `unresolved` — Word only gave a `file:///` or `cid:` reference; the bytes are
 *                not in the clipboard. Rendered as a placeholder, never as a
 *                silently broken `<img>`.
 */
export type ImageResolution = 'resolved' | 'external' | 'unresolved';

export interface WordImage {
  id: string;
  /** Renderable source, or empty when unresolved. */
  src: string;
  /** The reference exactly as Word wrote it. */
  originalSource: string;
  resolution: ImageResolution;
  /** Where the bytes came from. */
  origin: 'data-uri' | 'clipboard-blob' | 'http' | 'file' | 'cid' | 'none';
  width?: Length;
  height?: Length;
  /** Natural size when Word declared it separately from display size. */
  naturalWidth?: Length;
  naturalHeight?: Length;
  alt?: string;
  title?: string;
  crop?: ImageCrop;
  placement: ImagePlacement;
  /** VML shape id (`v:shapes` / `<v:shape id>`) linking img to its VML twin. */
  shapeId?: string;
  /** Preserved VML markup when the image came from a VML shape. */
  vml?: string;
  /** MIME type when known. */
  mimeType?: string;
  /** Byte length when the bytes are in hand. */
  byteLength?: number;
}
