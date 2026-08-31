import type {
  Border,
  Borders,
  Color,
  Length,
  LineSpacing,
  ParagraphAlignment,
  ParagraphFormatting,
  RunFormatting,
  Shading,
  TabStop,
  UnderlineStyle,
} from '../model/Style.js';
import { parseBoxShorthand, parseWordLength, roundTo } from './WordLengthParser.js';
import { parseHighlight, parseWordColor } from './WordColorParser.js';
import { unquote } from './WordCssTokenizer.js';

/**
 * Turns a CSS declaration map into model formatting.
 *
 * Shared by four callers — the stylesheet parser, the list level parser, the
 * run parser and the paragraph parser — so that `font-size:11.0pt` means
 * exactly one thing everywhere in the engine. Word writes the same property in
 * several spellings (`font-weight:bold` and `<b>`, `text-decoration:underline`
 * and `<u>`), and reconciling them in one place is what keeps the rest of the
 * parser from turning into a pile of special cases.
 */

/** Font-size context so `em`/`%` lengths inside a declaration resolve sanely. */
export interface FormattingContext {
  fontSizePx?: number;
}

const UNDERLINE_STYLE_MAP: Record<string, UnderlineStyle> = {
  single: 'single',
  double: 'double',
  thick: 'thick',
  dotted: 'dotted',
  dash: 'dashed',
  dashed: 'dashed',
  wave: 'wave',
  wavy: 'wave',
  none: 'none',
  words: 'single',
  'dot-dash': 'dashed',
  'dot-dot-dash': 'dashed',
};

export function parseRunFormattingFromCss(
  css: Record<string, string>,
  context: FormattingContext = {},
): RunFormatting {
  const formatting: RunFormatting = {};
  const fontSizeOptions = {
    defaultUnit: 'pt' as const,
    fontSizePx: context.fontSizePx ?? 16,
  };

  // `font` shorthand — Word uses it in bullet spacer spans: font:7.0pt "Times New Roman"
  const shorthand = css['font'];
  if (shorthand) {
    applyFontShorthand(shorthand, formatting, fontSizeOptions);
  }

  const fontFamily = css['font-family'] ?? css['mso-ascii-font-family'];
  if (fontFamily) {
    formatting.fontFamilyRaw = fontFamily;
    formatting.fontFamily = primaryFontFamily(fontFamily);
  }

  const fontSize = css['font-size'] ?? css['mso-ansi-font-size'];
  const parsedSize = parseWordLength(fontSize, fontSizeOptions);
  if (parsedSize) formatting.fontSize = parsedSize;

  // `mso-bidi-font-weight` and `mso-bidi-font-style` are deliberately NOT
  // consulted here. They describe the complex-script run, not the Latin one,
  // and Word writes bold text as
  //     <b style='mso-bidi-font-weight:normal'>Bold</b>
  // on essentially every bold run in every document. Reading that `normal` as
  // the weight cancels the <b> and silently un-bolds the whole paste.
  const weight = css['font-weight'];
  if (weight !== undefined) {
    const bold = isBoldWeight(weight);
    if (bold !== undefined) formatting.bold = bold;
  }

  const style = css['font-style'];
  if (style !== undefined) {
    const lower = style.trim().toLowerCase();
    if (lower === 'italic' || lower === 'oblique') formatting.italic = true;
    else if (lower === 'normal') formatting.italic = false;
  }

  applyTextDecoration(css, formatting);

  const color = parseWordColor(css['color']);
  if (color) formatting.color = color;

  // Highlight: `background`, `background-color` on a span, or `mso-highlight`.
  const highlight =
    parseHighlight(css['mso-highlight']) ??
    parseHighlight(css['background-color']) ??
    parseHighlight(css['background']);
  if (highlight && highlight.hex !== 'transparent') formatting.highlight = highlight;

  const verticalAlign = css['vertical-align'];
  if (verticalAlign) {
    const lower = verticalAlign.trim().toLowerCase();
    if (lower === 'super') formatting.verticalAlign = 'super';
    else if (lower === 'sub') formatting.verticalAlign = 'sub';
    else if (lower === 'baseline') formatting.verticalAlign = 'baseline';
  }

  const spacing = parseWordLength(css['letter-spacing'] ?? css['mso-character-spacing'], {
    defaultUnit: 'pt',
  });
  if (spacing) formatting.letterSpacing = spacing;

  const scale = css['mso-char-scale'] ?? css['font-stretch'];
  if (scale) {
    const percent = Number.parseFloat(scale);
    if (Number.isFinite(percent)) formatting.characterScale = percent;
  }

  const variant = css['font-variant'];
  if (variant && /small-caps/i.test(variant)) formatting.smallCaps = true;
  if (css['mso-style-textoutline-type'] === undefined) {
    const transform = css['text-transform'];
    if (transform && /uppercase/i.test(transform)) formatting.allCaps = true;
  }

  const language = css['mso-ansi-language'] ?? css['lang'];
  if (language) formatting.language = language.trim();

  const direction = css['direction'];
  if (direction) {
    const lower = direction.trim().toLowerCase();
    if (lower === 'rtl' || lower === 'ltr') formatting.direction = lower;
  }

  if (
    css['mso-hide'] === 'all' ||
    /^\s*none\s*$/i.test(css['display'] ?? '') ||
    /^\s*hidden\s*$/i.test(css['visibility'] ?? '')
  ) {
    formatting.hidden = true;
  }

  return formatting;
}

function applyFontShorthand(
  shorthand: string,
  formatting: RunFormatting,
  options: { defaultUnit: 'pt'; fontSizePx: number },
): void {
  // `font: [style] [variant] [weight] size[/line-height] family`
  const value = shorthand.trim();
  const sizeMatch = /(^|\s)((?:\d+\.?\d*|\.\d+)(?:pt|px|in|cm|mm|pc|em|%))(?:\s*\/\s*[^\s]+)?\s+/.exec(
    value,
  );
  if (sizeMatch) {
    const size = parseWordLength(sizeMatch[2], options);
    if (size) formatting.fontSize = size;
    const family = value.slice(sizeMatch.index + sizeMatch[0].length).trim();
    if (family) {
      formatting.fontFamilyRaw = family;
      formatting.fontFamily = primaryFontFamily(family);
    }
    const prefix = value.slice(0, sizeMatch.index).toLowerCase();
    if (/\bitalic\b|\boblique\b/.test(prefix)) formatting.italic = true;
    if (/\bbold\b|\b[6-9]00\b/.test(prefix)) formatting.bold = true;
    if (/\bsmall-caps\b/.test(prefix)) formatting.smallCaps = true;
  }
}

function applyTextDecoration(css: Record<string, string>, formatting: RunFormatting): void {
  const decoration = [
    css['text-decoration'] ?? '',
    css['text-decoration-line'] ?? '',
    css['mso-text-decoration'] ?? '',
  ]
    .join(' ')
    .toLowerCase();

  if (/\bunderline\b/.test(decoration)) {
    formatting.underline = 'single';
  }
  if (/\bline-through\b/.test(decoration)) {
    formatting.strike = true;
  }
  if (/\bnone\b/.test(decoration) && !/\bunderline\b|\bline-through\b/.test(decoration)) {
    // Word writes `text-decoration:none` to cancel an inherited hyperlink
    // underline. Record the cancellation rather than leaving it ambiguous.
    formatting.underline = 'none';
  }

  const underlineStyle = css['text-underline-style'] ?? css['mso-text-underline-style'];
  if (underlineStyle) {
    const mapped = UNDERLINE_STYLE_MAP[underlineStyle.trim().toLowerCase()];
    if (mapped) formatting.underline = mapped;
  }
  const underlineDouble = css['text-underline'];
  if (underlineDouble) {
    const mapped = UNDERLINE_STYLE_MAP[underlineDouble.trim().toLowerCase()];
    if (mapped) formatting.underline = mapped;
  }
  const underlineColor = parseWordColor(
    css['text-underline-color'] ?? css['text-decoration-color'],
  );
  if (underlineColor) formatting.underlineColor = underlineColor;

  const strikeStyle = css['mso-text-strike'] ?? css['text-line-through-style'];
  if (strikeStyle && /double/i.test(strikeStyle)) {
    formatting.strike = true;
    formatting.doubleStrike = true;
  }
}

export function parseParagraphFormattingFromCss(
  css: Record<string, string>,
  context: FormattingContext = {},
): ParagraphFormatting {
  const formatting: ParagraphFormatting = {};
  const lengthOptions = { defaultUnit: 'pt' as const, fontSizePx: context.fontSizePx ?? 16 };

  const alignment = parseAlignment(css['text-align'] ?? css['mso-alignment']);
  if (alignment) formatting.alignment = alignment;

  const marginShorthand = parseBoxShorthand(css['margin'], lengthOptions);
  const marginTop =
    parseWordLength(css['margin-top'], lengthOptions) ??
    parseWordLength(css['mso-margin-top-alt'], lengthOptions) ??
    marginShorthand?.top;
  const marginBottom =
    parseWordLength(css['margin-bottom'], lengthOptions) ??
    parseWordLength(css['mso-margin-bottom-alt'], lengthOptions) ??
    marginShorthand?.bottom;
  const marginLeft =
    parseWordLength(css['margin-left'], lengthOptions) ?? marginShorthand?.left;
  const marginRight =
    parseWordLength(css['margin-right'], lengthOptions) ?? marginShorthand?.right;

  if (marginLeft) formatting.marginLeft = marginLeft;
  if (marginRight) formatting.marginRight = marginRight;
  if (marginTop) formatting.spaceBefore = marginTop;
  if (marginBottom) formatting.spaceAfter = marginBottom;

  const textIndent = parseWordLength(css['text-indent'], lengthOptions);
  if (textIndent) formatting.textIndent = textIndent;

  const lineSpacing = parseLineSpacing(css, lengthOptions);
  if (lineSpacing) formatting.lineSpacing = lineSpacing;

  // Word expresses "keep with next" as page-break-after:avoid, and
  // "keep lines together" as page-break-inside:avoid.
  if (/avoid/i.test(css['page-break-after'] ?? '')) formatting.keepWithNext = true;
  if (/avoid/i.test(css['page-break-inside'] ?? '')) formatting.keepLines = true;
  if (/always|left|right/i.test(css['page-break-before'] ?? '')) {
    formatting.pageBreakBefore = true;
  }
  if (/avoid/i.test(css['mso-pagination'] ?? '')) formatting.keepLines = true;
  const widows = css['widows'] ?? css['orphans'];
  if (widows !== undefined) {
    const n = Number.parseInt(widows, 10);
    if (Number.isFinite(n)) formatting.widowControl = n > 0;
  }
  if (/widow-orphan/i.test(css['mso-pagination'] ?? '')) formatting.widowControl = true;

  const outline = css['mso-outline-level'];
  if (outline) {
    const level = Number.parseInt(outline, 10);
    if (Number.isFinite(level) && level >= 1 && level <= 9) formatting.outlineLevel = level;
  }

  const borders = parseBorders(css, lengthOptions);
  if (borders) formatting.borders = borders;

  const shading = parseShading(css);
  if (shading) formatting.shading = shading;

  const background = parseWordColor(css['background-color'] ?? css['background']);
  if (background && background.hex !== 'transparent') formatting.backgroundColor = background;

  const direction = css['direction'];
  if (direction) {
    const lower = direction.trim().toLowerCase();
    if (lower === 'rtl' || lower === 'ltr') formatting.direction = lower;
  }

  const tabStops = parseTabStops(css['tab-stops'] ?? css['mso-tab-stops'], lengthOptions);
  if (tabStops && tabStops.length) formatting.tabStops = tabStops;

  return formatting;
}

export function parseAlignment(value: string | undefined): ParagraphAlignment | undefined {
  if (!value) return undefined;
  const lower = value.trim().toLowerCase();
  switch (lower) {
    case 'left':
    case 'start':
      return 'left';
    case 'right':
    case 'end':
      return 'right';
    case 'center':
    case 'centre':
      return 'center';
    case 'justify':
    case 'justify-all':
      return 'justify';
    default:
      return undefined;
  }
}

function parseLineSpacing(
  css: Record<string, string>,
  options: { defaultUnit: 'pt'; fontSizePx: number },
): LineSpacing | undefined {
  const raw = css['line-height'] ?? css['mso-line-height-alt'];
  if (!raw) return undefined;
  const value = raw.trim();
  const lower = value.toLowerCase();
  if (lower === 'normal') return undefined;

  const rule = css['mso-line-height-rule'];
  const exact = rule ? /exact/i.test(rule) : false;
  const atLeast = rule ? /at-?least/i.test(rule) : false;

  if (value.endsWith('%')) {
    const percent = Number.parseFloat(value);
    if (!Number.isFinite(percent)) return undefined;
    return { rule: 'multiple', value: roundTo(percent / 100, 4), raw };
  }

  const length = parseWordLength(value, options);
  if (length) {
    return {
      rule: exact ? 'exact' : atLeast ? 'atLeast' : 'exact',
      value: length.px,
      length,
      raw,
    };
  }

  const multiplier = Number.parseFloat(value);
  if (Number.isFinite(multiplier)) {
    return { rule: 'multiple', value: multiplier, raw };
  }
  return undefined;
}

const BORDER_SIDES = ['top', 'right', 'bottom', 'left'] as const;

export function parseBorders(
  css: Record<string, string>,
  options: { defaultUnit: 'pt'; fontSizePx: number },
): Borders | undefined {
  const shorthand = css['border'];
  const result: Borders = {};
  let found = false;
  const rawExtras: Record<string, string> = {};

  const base = shorthand ? parseBorderValue(shorthand, options) : undefined;
  if (base) {
    for (const side of BORDER_SIDES) result[side] = base;
    found = true;
  }

  for (const side of BORDER_SIDES) {
    const value = css[`border-${side}`];
    if (value) {
      const border = parseBorderValue(value, options);
      if (border) {
        result[side] = border;
        found = true;
      } else if (/^\s*none\s*$/i.test(value)) {
        result[side] = { style: 'none', raw: value };
        found = true;
      }
    }
    const alt = css[`mso-border-${side}-alt`];
    if (alt) rawExtras[`mso-border-${side}-alt`] = alt;
  }

  const altAll = css['mso-border-alt'];
  if (altAll) rawExtras['mso-border-alt'] = altAll;
  const insideH = css['mso-border-insideh'];
  if (insideH) rawExtras['mso-border-insideh'] = insideH;
  const insideV = css['mso-border-insidev'];
  if (insideV) rawExtras['mso-border-insidev'] = insideV;

  if (Object.keys(rawExtras).length > 0) {
    result.raw = rawExtras;
    found = true;
  }
  return found ? result : undefined;
}

const BORDER_STYLE_WORDS = new Set([
  'none', 'hidden', 'solid', 'double', 'dotted', 'dashed', 'groove', 'ridge', 'inset', 'outset',
]);

export function parseBorderValue(
  value: string,
  options: { defaultUnit: 'pt'; fontSizePx: number },
): Border | undefined {
  const raw = value.trim();
  if (!raw || /^none$/i.test(raw)) {
    return raw ? { style: 'none', raw } : undefined;
  }
  const parts = raw.split(/\s+/);
  const border: Border = { style: 'solid', raw };
  let sawStyle = false;
  for (const part of parts) {
    const lower = part.toLowerCase();
    if (BORDER_STYLE_WORDS.has(lower)) {
      border.style = lower === 'hidden' ? 'none' : (lower as Border['style']);
      sawStyle = true;
      continue;
    }
    const width = parseWordLength(part, options);
    if (width) {
      border.width = width;
      continue;
    }
    const color = parseWordColor(part);
    if (color) border.color = color;
  }
  if (!sawStyle && !border.width && !border.color) return undefined;
  return border;
}

export function parseShading(css: Record<string, string>): Shading | undefined {
  const raw = css['mso-shading'] ?? css['background'] ?? css['background-color'];
  if (!raw) return undefined;
  const fill = parseWordColor(css['background-color'] ?? css['background'] ?? css['mso-shading']);
  const pattern = css['mso-pattern'];
  if (!fill && !pattern) return undefined;
  const shading: Shading = { raw };
  if (fill && fill.hex !== 'transparent') shading.fill = fill;
  if (pattern) shading.pattern = pattern;
  return shading.fill || shading.pattern ? shading : undefined;
}

/**
 * Word writes tab stops as `tab-stops:list .5in` or
 * `tab-stops:center 207.65pt right 415.3pt` or `tab-stops:45.8pt`.
 */
export function parseTabStops(
  value: string | undefined,
  options: { defaultUnit: 'pt'; fontSizePx: number },
): TabStop[] | undefined {
  if (!value) return undefined;
  const tokens = value.trim().split(/\s+/);
  const stops: TabStop[] = [];
  let alignment: TabStop['alignment'] = 'left';
  let leader: TabStop['leader'] | undefined;

  for (const token of tokens) {
    const lower = token.toLowerCase();
    if (lower === 'left' || lower === 'right' || lower === 'center' || lower === 'decimal' || lower === 'bar') {
      alignment = lower;
      continue;
    }
    if (lower === 'list') {
      // `tab-stops:list .5in` — the stop belonging to a list level.
      alignment = 'left';
      continue;
    }
    if (lower === 'dotted' || lower === 'dot') {
      leader = 'dot';
      continue;
    }
    if (lower === 'hyphen') {
      leader = 'hyphen';
      continue;
    }
    if (lower === 'underscore' || lower === 'thick') {
      leader = 'underscore';
      continue;
    }
    const position = parseWordLength(token, options);
    if (position) {
      const stop: TabStop = { position, alignment };
      if (leader) stop.leader = leader;
      stops.push(stop);
      alignment = 'left';
      leader = undefined;
    }
  }
  return stops.length ? stops : undefined;
}

/** Take the first family from a font stack and strip Word's quoting. */
export function primaryFontFamily(fontFamily: string): string {
  const first = fontFamily.split(',')[0] ?? fontFamily;
  return unquote(first).trim();
}

function isBoldWeight(value: string): boolean | undefined {
  const lower = value.trim().toLowerCase();
  if (lower === 'bold' || lower === 'bolder') return true;
  if (lower === 'normal' || lower === 'lighter') return false;
  const numeric = Number.parseInt(lower, 10);
  if (Number.isFinite(numeric)) return numeric >= 600;
  return undefined;
}

/** Merge two run formattings, with `child` winning where it declares a value. */
export function mergeRunFormatting(parent: RunFormatting, child: RunFormatting): RunFormatting {
  return { ...parent, ...stripUndefined(child) };
}

/** Merge two paragraph formattings, with `child` winning. */
export function mergeParagraphFormatting(
  parent: ParagraphFormatting,
  child: ParagraphFormatting,
): ParagraphFormatting {
  const merged: ParagraphFormatting = { ...parent, ...stripUndefined(child) };
  if (parent.borders && child.borders) {
    merged.borders = { ...parent.borders, ...stripUndefined(child.borders) };
  }
  return merged;
}

function stripUndefined<T extends object>(value: T): Partial<T> {
  const out: Record<string, unknown> = {};
  for (const [key, v] of Object.entries(value)) {
    if (v !== undefined) out[key] = v;
  }
  return out as Partial<T>;
}

/** True when a run formatting object carries no information at all. */
export function isEmptyFormatting(formatting: RunFormatting): boolean {
  return Object.keys(formatting).length === 0;
}

/** Structural equality for run formatting, used to decide whether runs merge. */
export function runFormattingEquals(a: RunFormatting, b: RunFormatting): boolean {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  for (const key of keys) {
    const av = a[key as keyof RunFormatting];
    const bv = b[key as keyof RunFormatting];
    if (av === bv) continue;
    if (av === undefined || bv === undefined) return false;
    if (isLength(av) && isLength(bv)) {
      if (av.px !== bv.px) return false;
      continue;
    }
    if (isColor(av) && isColor(bv)) {
      if (av.hex !== bv.hex) return false;
      continue;
    }
    return false;
  }
  return true;
}

function isLength(value: unknown): value is Length {
  return typeof value === 'object' && value !== null && 'px' in value && 'twips' in value;
}

function isColor(value: unknown): value is Color {
  return typeof value === 'object' && value !== null && 'hex' in value;
}
