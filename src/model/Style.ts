import type { WordListDefinition } from './List.js';

/**
 * Formatting primitives shared by runs, paragraphs, tables and cells.
 *
 * Everything here is *canonical*: lengths have already been passed through
 * `parseWordLength()` and are expressed in the engine's logical unit (px,
 * see `src/word/WordLengthParser.ts`). The original Word literal is kept
 * alongside so that nothing is lost and diagnostics can show provenance.
 */

/** A length that remembers where it came from. */
export interface Length {
  /** Canonical value in CSS pixels (96 px == 1 in). */
  px: number;
  /** Value in twips (1/1440 in) — Word's own native unit. */
  twips: number;
  /** Original literal exactly as Word wrote it, e.g. `-.25in`. */
  raw: string;
}

/** A colour that remembers where it came from. */
export interface Color {
  /** Normalised `#rrggbb`, or `transparent`. */
  hex: string;
  /** Original literal, e.g. `windowtext`, `yellow`, `#1F497D`. */
  raw: string;
}

export type UnderlineStyle =
  | 'single'
  | 'double'
  | 'thick'
  | 'dotted'
  | 'dashed'
  | 'wave'
  | 'none';

export type VerticalAlignRun = 'baseline' | 'super' | 'sub';

export type TextDirection = 'ltr' | 'rtl';

/** Character-level formatting. Every field is optional: absent means "inherit". */
export interface RunFormatting {
  fontFamily?: string;
  /** Font stack exactly as Word declared it, before de-quoting. */
  fontFamilyRaw?: string;
  fontSize?: Length;
  bold?: boolean;
  italic?: boolean;
  underline?: UnderlineStyle;
  underlineColor?: Color;
  strike?: boolean;
  doubleStrike?: boolean;
  smallCaps?: boolean;
  allCaps?: boolean;
  color?: Color;
  /** `background-color` / `mso-highlight`. */
  highlight?: Color;
  verticalAlign?: VerticalAlignRun;
  /** `letter-spacing` / `mso-character-spacing`. */
  letterSpacing?: Length;
  /** Horizontal scale percentage (`mso-font-kerning` sibling `font-stretch`). */
  characterScale?: number;
  /** `lang` attribute or `mso-ansi-language`. */
  language?: string;
  direction?: TextDirection;
  /** True when Word marked the run hidden (`mso-hide:all`, `display:none`). */
  hidden?: boolean;
  /** Word style name applied at character level, e.g. `Emphasis`. */
  characterStyle?: string;
}

export type ParagraphAlignment = 'left' | 'right' | 'center' | 'justify';

export type LineSpacingRule = 'auto' | 'exact' | 'atLeast' | 'multiple';

export interface LineSpacing {
  rule: LineSpacingRule;
  /** For `multiple`: the multiplier (1.5, 2 …). For the rest: a length. */
  value: number;
  length?: Length;
  raw: string;
}

export type BorderStyle =
  | 'none'
  | 'solid'
  | 'double'
  | 'dotted'
  | 'dashed'
  | 'groove'
  | 'ridge'
  | 'inset'
  | 'outset';

export interface Border {
  style: BorderStyle;
  width?: Length;
  color?: Color;
  raw: string;
}

export interface Borders {
  top?: Border;
  right?: Border;
  bottom?: Border;
  left?: Border;
  /** Word `mso-border-*-alt` values kept for diagnostics. */
  raw?: Record<string, string>;
}

export interface Shading {
  fill?: Color;
  /** `mso-shading` pattern colour. */
  pattern?: string;
  raw: string;
}

export interface TabStop {
  position: Length;
  alignment: 'left' | 'right' | 'center' | 'decimal' | 'bar';
  leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'middot';
}

/** Paragraph-level formatting. */
export interface ParagraphFormatting {
  alignment?: ParagraphAlignment;
  marginLeft?: Length;
  marginRight?: Length;
  /** Positive = first line indent, negative = hanging indent. */
  textIndent?: Length;
  spaceBefore?: Length;
  spaceAfter?: Length;
  lineSpacing?: LineSpacing;
  keepWithNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  /** `widows` / `orphans` when Word declared them. */
  widowControl?: boolean;
  tabStops?: TabStop[];
  borders?: Borders;
  shading?: Shading;
  direction?: TextDirection;
  /** `mso-outline-level` — Word's own heading outline level (1..9). */
  outlineLevel?: number;
  /** Background colour applied to the paragraph box. */
  backgroundColor?: Color;
}

/**
 * A Word style definition harvested from the clipboard stylesheet
 * (`p.MsoHeading1`, `span.EmphasisChar`, …).
 */
export interface WordStyleDefinition {
  /** Normalised identifier, e.g. `msoheading1`. */
  id: string;
  /** Human name recovered from `mso-style-name`, else the class name. */
  name: string;
  /** Selectors this definition was declared under. */
  selectors: string[];
  type: 'paragraph' | 'character' | 'table' | 'unknown';
  /** `mso-style-parent`. */
  parent?: string;
  /** `mso-style-link` (paired paragraph/character style). */
  link?: string;
  /** Raw declarations exactly as parsed, lower-cased property names. */
  declarations: Record<string, string>;
  run: RunFormatting;
  paragraph: ParagraphFormatting;
  /** `mso-outline-level` promoted from declarations. */
  outlineLevel?: number;
}

/** `@font-face` block from the Word clipboard stylesheet. */
export interface WordFontDefinition {
  family: string;
  panose1?: string;
  charset?: string;
  /** `mso-font-alt` — the substitute Word suggests. */
  alt?: string;
  /** `mso-generic-font-family`, e.g. `roman`, `swiss`, `decorative`. */
  genericFamily?: string;
  /** True for Symbol/Wingdings-class fonts where glyph bytes are not Unicode. */
  isSymbolFont: boolean;
  declarations: Record<string, string>;
}

/** Everything recovered from the clipboard `<style>` blocks. */
export interface WordStyleSheet {
  /** Keyed by normalised style id. */
  styles: Record<string, WordStyleDefinition>;
  /** Keyed by lower-cased font family. */
  fonts: Record<string, WordFontDefinition>;
  /** Keyed by list id (`l0`, `l1`, …) — see `model/List.ts`. */
  lists: Record<string, WordListDefinition>;
  /** `@page` declarations, e.g. section margins. */
  pages: Record<string, Record<string, string>>;
  /** Every rule, in source order, for debugging. */
  rules: WordCssRule[];
  /** The concatenated raw CSS text the stylesheet was built from. */
  rawCss: string;
}

export interface WordCssRule {
  kind: 'style' | 'at-list' | 'at-font-face' | 'at-page' | 'at-other';
  selector: string;
  declarations: Record<string, string>;
  raw: string;
}
