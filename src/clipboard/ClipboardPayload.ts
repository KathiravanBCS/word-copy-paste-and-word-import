/**
 * The immutable record of what the browser actually handed us.
 *
 * Nothing in the engine is allowed to mutate a payload. Cleanup, sanitisation
 * and normalisation all happen on parsed clones, so a real-world paste can
 * always be re-examined byte for byte after the fact.
 */

/** An image flavour that arrived alongside the HTML (Word supplies these on some platforms). */
export interface ClipboardImageItem {
  /** MIME type as reported by the clipboard, e.g. `image/png`. */
  mimeType: string;
  /** The bytes, when they could be read synchronously from a DataTransferItem. */
  blob?: Blob;
  /** A `data:` URI built from the bytes, when materialised. */
  dataUri?: string;
  /** An object URL, when materialised. Caller owns revocation. */
  objectUrl?: string;
  /** File name when the clipboard reported one (`image.png`, `clip_image001.png`). */
  name?: string;
  byteLength?: number;
}

export interface ClipboardPayload {
  /** The `text/html` flavour, byte for byte, never sanitised. */
  html?: string;
  /** The `text/plain` flavour. */
  text?: string;
  /** Every MIME type the clipboard advertised, in the order reported. */
  types: string[];
  /** Every flavour we were able to read, keyed by MIME type. */
  flavors: Record<string, string>;
  /** Image items recovered from the clipboard, in report order. */
  images: ClipboardImageItem[];
  /** Free-form hint about where the payload came from, e.g. `paste-event`. */
  source?: string;
  /** Epoch milliseconds when the payload was captured. */
  capturedAt: number;
}

/** Build an empty payload. */
export function emptyClipboardPayload(source = 'unknown'): ClipboardPayload {
  return { types: [], flavors: {}, images: [], source, capturedAt: Date.now() };
}

/**
 * Convenience constructor for tests, fixtures and server-side use, where there
 * is no real `DataTransfer` — just an HTML string.
 */
export function clipboardPayloadFromHtml(html: string, text?: string): ClipboardPayload {
  const payload: ClipboardPayload = {
    html,
    types: ['text/html'],
    flavors: { 'text/html': html },
    images: [],
    source: 'string',
    capturedAt: Date.now(),
  };
  if (text !== undefined) {
    payload.text = text;
    payload.types.push('text/plain');
    payload.flavors['text/plain'] = text;
  }
  return payload;
}

/**
 * The raw, undestroyed clipboard document plus the detection verdict.
 *
 * Kept separate from `ClipboardPayload` because detection is a decision *about*
 * a payload, and both need to be inspectable independently in the lab UI.
 */
export interface RawClipboardDocument {
  rawHtml: string;
  rawText?: string;
  detectedAsWord: boolean;
  detectionConfidence: number;
  detectionSignals: string[];
  payload: ClipboardPayload;
}
