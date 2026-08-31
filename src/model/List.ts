import type { Length, RunFormatting } from './Style.js';
import type { WordListLevel, WordNumberFormat, ListLevelType } from './ListLevel.js';

/**
 * A Word list definition (`@list l0 { … }` plus its `@list l0:levelN` blocks).
 *
 * This is authoritative. The renderer reads it; it never rewrites it.
 */
export interface WordListDefinition {
  /** The `@list` identifier as it appears in the CSS, e.g. `l0`. */
  listId: string;
  /** `mso-list-id` — the numbering-definition id inside the document. */
  msoListId?: string;
  /** `mso-list-type`: `simple`, `hybrid`, `multilevel`. */
  listType?: string;
  /** `mso-list-template-ids`. */
  templateIds?: string;
  /** Levels in ascending order; sparse input is filled to a dense array. */
  levels: WordListLevel[];
  declarations: Record<string, string>;
}

/**
 * The `lfo` (list format override) reference from `mso-list:l0 level1 lfo1`.
 *
 * Word uses the lfo to distinguish two paragraphs that share a numbering
 * definition but restart independently. Two paragraphs are in the same list
 * run only when listId *and* lfo agree.
 */
export interface WordListReference {
  listId: string;
  /** Zero-based. */
  level: number;
  lfo?: string;
  raw: string;
}

/**
 * The marker Word actually rendered next to a list item, harvested from the
 * `<span style='mso-list:Ignore'>` block inside `<![if !supportLists]>`.
 *
 * `text` is the literal Word displayed ("I.", "1.1", "a)"), so golden output
 * can be compared character for character against what the user saw in Word.
 */
export interface ListItemMarker {
  type: ListLevelType;
  /** Bullet glyph in Unicode after symbol-font mapping (bullets only). */
  glyph?: string;
  /** The raw glyph byte Word emitted, pre-mapping. */
  rawGlyph?: string;
  /** Font required to render `rawGlyph`, e.g. `Symbol`. */
  font?: string;
  /** True when the glyph came through a symbol-font code page mapping. */
  fontMapped?: boolean;
  /** Literal marker text Word rendered, e.g. `1.1`, `I.`, `a)`. */
  text?: string;
  numberFormat?: WordNumberFormat;
  /** `mso-level-text` pattern from the list definition, e.g. `%1.%2`. */
  levelText?: string;
  startAt?: number;
  /** Character formatting of the marker span itself. */
  formatting?: RunFormatting;
  /** Where this marker came from, for diagnostics. */
  source: 'mso-list-ignore' | 'list-definition' | 'html-list' | 'text-heuristic';
}

/**
 * Attached to a paragraph that participates in a list.
 */
export interface ListItemInfo {
  /** Word list id, e.g. `l0`. */
  listId: string;
  /** Zero-based nesting level. */
  level: number;
  lfo?: string;
  /**
   * Identity of the *contiguous list run* this item belongs to, assigned by
   * NormalizeLists. Two items with the same `listInstanceId` are the same
   * visual list; a restart or an interruption yields a new instance.
   */
  listInstanceId?: string;
  marker: ListItemMarker;
  /** The level definition resolved from the list definition, if present. */
  levelDefinition?: WordListLevel;
  /** True when this item starts a new numbering run (restart). */
  restart?: boolean;
  /** Effective start value for a restarting item. */
  startAt?: number;
  /** Indentation Word declared on the paragraph itself. */
  marginLeft?: Length;
  textIndent?: Length;
  /** True when the item came from a real `<ol>`/`<ul>` rather than mso-list. */
  fromHtmlList?: boolean;
}
