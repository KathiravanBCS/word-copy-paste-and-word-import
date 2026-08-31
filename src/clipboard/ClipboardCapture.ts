import {
  emptyClipboardPayload,
  type ClipboardImageItem,
  type ClipboardPayload,
} from './ClipboardPayload.js';

/**
 * Browser clipboard capture.
 *
 * Two rules govern this layer:
 *
 *  1. Read *everything* the clipboard offers, not just `text/html`. Word puts
 *     useful flavours alongside the HTML on some platforms, and the image
 *     bytes are frequently the only way to recover a picture whose HTML `src`
 *     is a dead `file:///` path.
 *  2. Never sanitise here. Capture is a recording device; interpretation
 *     happens downstream where it can be diagnosed and tested.
 */

export interface CaptureOptions {
  /**
   * Additional MIME types to attempt to read as text beyond the defaults.
   * Word/Office sometimes advertise vendor flavours worth keeping for debugging.
   */
  extraTextTypes?: string[];
  /** Read image items into `Blob`s (default true). */
  readImages?: boolean;
  /**
   * Materialise captured image blobs as `data:` URIs. Costs a read per image
   * but makes the payload serialisable and fixture-exportable. Default true.
   */
  materialiseDataUris?: boolean;
  /** Hard cap on a single image, in bytes. Larger images are skipped with a note. Default 16 MB. */
  maxImageBytes?: number;
}

const DEFAULT_TEXT_TYPES = [
  'text/html',
  'text/plain',
  'text/rtf',
  'application/rtf',
  'text/uri-list',
  'text/csv',
  'text/xml',
  'application/xhtml+xml',
];

const DEFAULT_MAX_IMAGE_BYTES = 16 * 1024 * 1024;

/**
 * Capture a payload synchronously from a paste event.
 *
 * Synchronous by design: `DataTransfer` contents are only valid during the
 * event dispatch, so the text flavours must be read before any `await`. Image
 * blobs are grabbed synchronously via `getAsFile()`; turning them into data
 * URIs is asynchronous and lives in {@link materialiseImages}.
 */
export function captureFromPasteEvent(
  event: ClipboardEvent,
  options: CaptureOptions = {},
): ClipboardPayload {
  const payload = emptyClipboardPayload('paste-event');
  const data = event.clipboardData;
  if (!data) return payload;

  payload.types = Array.from(data.types ?? []);

  const textTypes = new Set([...DEFAULT_TEXT_TYPES, ...(options.extraTextTypes ?? [])]);
  for (const type of payload.types) {
    if (!textTypes.has(type) && !type.startsWith('text/')) continue;
    let value = '';
    try {
      value = data.getData(type);
    } catch {
      // Some browsers refuse exotic flavours; a failed read is not fatal.
      continue;
    }
    if (value) payload.flavors[type] = value;
  }

  payload.html = payload.flavors['text/html'];
  payload.text = payload.flavors['text/plain'];

  if (options.readImages !== false) {
    payload.images = readImageItems(data, options.maxImageBytes ?? DEFAULT_MAX_IMAGE_BYTES);
  }

  return payload;
}

/** Pull image files out of a DataTransfer without consuming the text flavours. */
function readImageItems(data: DataTransfer, maxBytes: number): ClipboardImageItem[] {
  const images: ClipboardImageItem[] = [];
  const items = data.items;
  if (items) {
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (!item || item.kind !== 'file') continue;
      if (!item.type.startsWith('image/')) continue;
      const file = item.getAsFile();
      if (!file) continue;
      if (file.size > maxBytes) continue;
      const entry: ClipboardImageItem = {
        mimeType: file.type || item.type,
        blob: file,
        byteLength: file.size,
      };
      const name = (file as File).name;
      if (name) entry.name = name;
      images.push(entry);
    }
  }
  if (images.length === 0 && data.files) {
    for (let i = 0; i < data.files.length; i++) {
      const file = data.files[i];
      if (!file || !file.type.startsWith('image/')) continue;
      if (file.size > maxBytes) continue;
      images.push({
        mimeType: file.type,
        blob: file,
        byteLength: file.size,
        name: file.name,
      });
    }
  }
  return images;
}

/**
 * Turn captured blobs into data URIs (and optionally object URLs).
 *
 * Returns a *new* payload; the input is left untouched, per the immutability
 * rule for raw payloads.
 */
export async function materialiseImages(
  payload: ClipboardPayload,
  options: { objectUrls?: boolean } = {},
): Promise<ClipboardPayload> {
  if (payload.images.length === 0) return payload;
  const images = await Promise.all(
    payload.images.map(async (image) => {
      if (!image.blob || image.dataUri) return image;
      const next: ClipboardImageItem = { ...image };
      try {
        next.dataUri = await blobToDataUri(image.blob);
      } catch {
        // A blob we cannot read simply stays unmaterialised; the image parser
        // will report it as unresolved rather than fabricate a source.
      }
      if (options.objectUrls && typeof URL !== 'undefined' && URL.createObjectURL) {
        next.objectUrl = URL.createObjectURL(image.blob);
      }
      return next;
    }),
  );
  return { ...payload, images };
}

function blobToDataUri(blob: Blob): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error ?? new Error('FileReader failed'));
    reader.onload = () => {
      const result = reader.result;
      if (typeof result === 'string') resolve(result);
      else reject(new Error('Unexpected FileReader result'));
    };
    reader.readAsDataURL(blob);
  });
}

/**
 * Capture from the async Clipboard API (`navigator.clipboard.read()`).
 *
 * Useful for a "read clipboard" button in the lab UI where there is no paste
 * event to hook. Requires a user gesture and clipboard-read permission.
 */
export async function captureFromNavigatorClipboard(
  options: CaptureOptions = {},
): Promise<ClipboardPayload> {
  const payload = emptyClipboardPayload('navigator.clipboard');
  const clipboard = typeof navigator !== 'undefined' ? navigator.clipboard : undefined;
  if (!clipboard || typeof clipboard.read !== 'function') {
    throw new Error('navigator.clipboard.read() is not available in this browser context');
  }
  const items = await clipboard.read();
  const maxBytes = options.maxImageBytes ?? DEFAULT_MAX_IMAGE_BYTES;
  for (const item of items) {
    for (const type of item.types) {
      if (!payload.types.includes(type)) payload.types.push(type);
      if (type.startsWith('text/')) {
        const blob = await item.getType(type);
        payload.flavors[type] = await blob.text();
      } else if (type.startsWith('image/') && options.readImages !== false) {
        const blob = await item.getType(type);
        if (blob.size > maxBytes) continue;
        payload.images.push({ mimeType: type, blob, byteLength: blob.size });
      }
    }
  }
  payload.html = payload.flavors['text/html'];
  payload.text = payload.flavors['text/plain'];
  if (options.materialiseDataUris !== false) {
    return materialiseImages(payload);
  }
  return payload;
}

/**
 * Install a paste listener that hands over the captured payload.
 *
 * The listener calls `preventDefault()` so the browser never inserts the raw
 * Word HTML into a contenteditable behind our back — that would both bypass
 * the engine and inject unsanitised markup.
 */
export function installPasteCapture(
  target: EventTarget,
  handler: (payload: ClipboardPayload, event: ClipboardEvent) => void,
  options: CaptureOptions & { preventDefault?: boolean } = {},
): () => void {
  const listener = (event: Event): void => {
    const clipboardEvent = event as ClipboardEvent;
    if (options.preventDefault !== false) clipboardEvent.preventDefault();
    handler(captureFromPasteEvent(clipboardEvent, options), clipboardEvent);
  };
  target.addEventListener('paste', listener);
  return () => target.removeEventListener('paste', listener);
}
