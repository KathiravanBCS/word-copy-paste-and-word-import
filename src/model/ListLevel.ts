import type { Length, RunFormatting } from './Style.js';

/**
 * Word's `mso-level-number-format` values, normalised.
 *
 * These are Word's own names — the engine never re-interprets them into an
 * application numbering scheme. `custom` means Word declared a format string
 * the engine has no CSS counter-style equivalent for; `levelText` is then the
 * authority.
 */
export type WordNumberFormat =
  | 'decimal'
  | 'decimal-leading-zero'
  | 'lower-alpha'
  | 'upper-alpha'
  | 'lower-roman'
  | 'upper-roman'
  | 'ordinal'
  | 'ordinal-text'
  | 'cardinal-text'
  | 'chicago'
  | 'hebrew-1'
  | 'arabic-alpha'
  | 'chosung'
  | 'ganada'
  | 'japanese-counting'
  | 'bullet'
  | 'image'
  | 'none'
  | 'custom';

export type ListLevelType = 'number' | 'bullet' | 'image' | 'none';

export type LevelJustification = 'left' | 'center' | 'right';

/** What follows the number/bullet before the paragraph text begins. */
export type LevelSuffix = 'tab' | 'space' | 'nothing';

/**
 * One level of a Word list definition, straight out of `@list lN:levelM`.
 *
 * Nothing in here is inferred from the rendered marker text; it is Word's
 * declaration. The rendered marker Word actually emitted is carried separately
 * on each list item (`ListItemMarker.text`) so the two can be cross-checked.
 */
export interface WordListLevel {
  /** Zero-based level index (Word's `level1` is level 0 here). */
  level: number;
  type: ListLevelType;
  numberFormat: WordNumberFormat;
  /** Raw `mso-level-number-format` literal, e.g. `roman-upper`. */
  numberFormatRaw?: string;
  /**
   * `mso-level-text` with `%1`..`%9` placeholders intact, e.g. `%1.%2`.
   * For bullets this is the (already decoded) glyph.
   */
  levelText?: string;
  /** `mso-level-text` exactly as written, e.g. `"\F0B7"`. */
  levelTextRaw?: string;
  /** Resolved bullet glyph in Unicode (Symbol/Wingdings mapped through). */
  bulletGlyph?: string;
  /** The glyph byte Word wrote, before symbol-font mapping. */
  bulletGlyphRaw?: string;
  /** Font the glyph must be rendered in, e.g. `Symbol`, `Wingdings`. */
  bulletFont?: string;
  /** True when `bulletGlyph` required a symbol-font code page mapping. */
  bulletFontMapped?: boolean;
  /** `mso-level-start-at`. Defaults to 1 for numbered levels. */
  startAt?: number;
  /** `mso-level-tab-stop`. */
  tabStop?: Length;
  /** `mso-level-number-position`. */
  justification?: LevelJustification;
  /** `text-indent` declared on the level (negative == hanging). */
  textIndent?: Length;
  /** `margin-left` declared on the level. */
  marginLeft?: Length;
  /** `mso-level-indent`. */
  indent?: Length;
  /** What separates the marker from the text. */
  suffix?: LevelSuffix;
  /** Character formatting Word declared for the marker itself. */
  markerFormatting?: RunFormatting;
  /** `mso-level-legacy*` compatibility flags. */
  legacy?: boolean;
  /** All raw declarations for this level. */
  declarations: Record<string, string>;
}
