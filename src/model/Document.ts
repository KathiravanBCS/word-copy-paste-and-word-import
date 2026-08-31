import type { WordBlock } from './Block.js';
import type { WordListDefinition } from './List.js';
import type { WordStyleSheet } from './Style.js';
import type { WordImage } from './Image.js';
import type { WordHyperlink, WordBookmark } from './Hyperlink.js';
import type { WordDocumentMetadata } from './WordMetadata.js';
import type { WordDiagnostic } from '../diagnostics/UnsupportedFeature.js';
import type { WordDetectionResult } from '../detection/WordDetector.js';

/**
 * The canonical intermediate model — the source of truth for the whole engine.
 *
 * Everything downstream (renderers, editor adapters, exporters) reads this and
 * only this. Nothing downstream may re-interpret Word semantics; if a fact is
 * not in here, it was not in the clipboard.
 */
export interface WordDocument {
  /** Top-level block flow. */
  blocks: WordBlock[];
  /** Parsed clipboard stylesheet (styles, fonts, list definitions, @page). */
  styles: WordStyleSheet;
  /** List definitions, also reachable via `styles.lists`, in declaration order. */
  lists: WordListDefinition[];
  /** Images keyed by id, referenced from runs and image blocks. */
  images: Record<string, WordImage>;
  /** Hyperlinks keyed by id, referenced from runs. */
  hyperlinks: Record<string, WordHyperlink>;
  /** Bookmarks keyed by id. */
  bookmarks: Record<string, WordBookmark>;
  /** Everything the engine could not represent exactly. */
  diagnostics: WordDiagnostic[];
  metadata: WordDocumentMetadata;
  /** Word detection outcome that produced this document. */
  detection: WordDetectionResult;
  /**
   * The untouched clipboard HTML. Never mutated, never rendered directly —
   * kept so a real payload can always be re-examined and re-parsed.
   */
  rawHtml: string;
  /** `text/plain` flavour when the clipboard provided one. */
  rawText?: string;
}
