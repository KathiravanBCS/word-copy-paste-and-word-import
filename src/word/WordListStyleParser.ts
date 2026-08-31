import type { WordListDefinition, WordListReference } from '../model/List.js';
import type { WordListLevel, WordNumberFormat, ListLevelType } from '../model/ListLevel.js';
import type { CssRuleNode } from './WordCssTokenizer.js';
import { declarationsToMap } from './WordCssTokenizer.js';
import { parseWordLength } from './WordLengthParser.js';
import { parseRunFormattingFromCss } from './WordFormattingParser.js';
import { decodeMsoLevelText, resolveSymbolGlyph } from './WordSymbolFonts.js';
import type { DiagnosticCollector } from '../diagnostics/DiagnosticCollector.js';
import { DiagnosticCode } from '../diagnostics/UnsupportedFeature.js';

/**
 * Parses Word's `@list` at-rules into list definitions.
 *
 * A Word numbering definition reaches the clipboard as a family of at-rules:
 *
 *     @list l0
 *       {mso-list-id:1587389017; mso-list-type:hybrid;}
 *     @list l0:level1
 *       {mso-level-number-format:roman-upper; mso-level-text:"%1\.";
 *        mso-level-tab-stop:.5in; text-indent:-.5in;}
 *     @list l0:level2
 *       {mso-level-text:"%1\.%2"; mso-level-start-at:1; text-indent:-.5in;}
 *
 * This is the authority for how the list numbers. The engine records it as
 * Word wrote it — number format, level text with its `%N` placeholders, start
 * value, indentation — and never substitutes an application numbering scheme
 * for it. `%1.%2` stays `%1.%2`; it does not become "Section 1.01".
 */

/** `mso-level-number-format` literals -> canonical format names. */
const NUMBER_FORMAT_MAP: Record<string, WordNumberFormat> = {
  bullet: 'bullet',
  arabic: 'decimal',
  decimal: 'decimal',
  'arabic-leading-zero': 'decimal-leading-zero',
  'alpha-lower': 'lower-alpha',
  'alpha-upper': 'upper-alpha',
  'roman-lower': 'lower-roman',
  'roman-upper': 'upper-roman',
  ordinal: 'ordinal',
  'ordinal-text': 'ordinal-text',
  'cardinal-text': 'cardinal-text',
  chicago: 'chicago',
  'hebrew-1': 'hebrew-1',
  'hebrew-2': 'hebrew-1',
  'arabic-alpha': 'arabic-alpha',
  chosung: 'chosung',
  ganada: 'ganada',
  'japanese-counting': 'japanese-counting',
  image: 'image',
  none: 'none',
};

const LIST_SELECTOR = /^l(\d+)(?::level(\d+))?$/i;

export interface ParsedListRules {
  definitions: Record<string, WordListDefinition>;
  order: string[];
}

/**
 * Build list definitions from the `@list` rules in a tokenized stylesheet.
 *
 * Levels are stored densely (index 0 == Word's `level1`) with gaps filled by
 * synthesised placeholder levels, so downstream code can index by depth
 * without a null check on every access.
 */
export function parseListRules(
  rules: CssRuleNode[],
  diagnostics: DiagnosticCollector,
): ParsedListRules {
  const definitions: Record<string, WordListDefinition> = {};
  const order: string[] = [];
  const levelsById = new Map<string, Map<number, WordListLevel>>();

  for (const rule of rules) {
    if (rule.kind !== 'at-rule' || rule.atName !== 'list') continue;
    const prelude = (rule.prelude ?? '').trim();
    const match = LIST_SELECTOR.exec(prelude);
    if (!match) {
      diagnostics.info(
        DiagnosticCode.WORD_CSS_PARSE_WARNING,
        `Unrecognised @list selector "${prelude}"; the rule was kept as raw metadata.`,
        { details: { selector: prelude } },
      );
      continue;
    }

    const listId = `l${match[1]}`;
    const declarations = declarationsToMap(rule.declarations);

    if (!definitions[listId]) {
      definitions[listId] = { listId, levels: [], declarations: {} };
      levelsById.set(listId, new Map());
      order.push(listId);
    }
    const definition = definitions[listId]!;

    if (match[2] === undefined) {
      // `@list l0` — the list header.
      definition.declarations = { ...definition.declarations, ...declarations };
      const msoListId = declarations['mso-list-id'];
      if (msoListId) definition.msoListId = msoListId.trim();
      const listType = declarations['mso-list-type'];
      if (listType) definition.listType = listType.trim();
      const templateIds = declarations['mso-list-template-ids'];
      if (templateIds) definition.templateIds = templateIds.trim();
      continue;
    }

    // `@list l0:levelN` — one level. Word's levels are 1-based; ours are 0-based.
    const wordLevel = Number.parseInt(match[2], 10);
    if (!Number.isFinite(wordLevel) || wordLevel < 1) continue;
    const level = wordLevel - 1;
    const parsed = parseListLevel(level, declarations, diagnostics, listId);
    levelsById.get(listId)!.set(level, parsed);
  }

  // Densify.
  for (const listId of order) {
    const levelMap = levelsById.get(listId);
    if (!levelMap || levelMap.size === 0) continue;
    const maxLevel = Math.max(...levelMap.keys());
    const levels: WordListLevel[] = [];
    for (let i = 0; i <= maxLevel; i++) {
      const level = levelMap.get(i);
      levels.push(level ?? synthesiseMissingLevel(i, diagnostics, listId));
    }
    definitions[listId]!.levels = levels;
  }

  return { definitions, order };
}

function synthesiseMissingLevel(
  level: number,
  diagnostics: DiagnosticCollector,
  listId: string,
): WordListLevel {
  diagnostics.warn(
    DiagnosticCode.WORD_LIST_LEVEL_MISSING,
    `List ${listId} declared no @list ${listId}:level${level + 1}; a decimal placeholder level was synthesised so nesting stays addressable.`,
    { details: { listId, level: level + 1 } },
  );
  return {
    level,
    type: 'number',
    numberFormat: 'decimal',
    levelText: `%${level + 1}.`,
    startAt: 1,
    declarations: {},
  };
}

/** Parse one `@list lN:levelM` declaration block. */
export function parseListLevel(
  level: number,
  declarations: Record<string, string>,
  diagnostics: DiagnosticCollector,
  listId: string,
): WordListLevel {
  const formatRaw = (declarations['mso-level-number-format'] ?? '').trim().toLowerCase();
  // Word omits the property entirely for plain arabic numbering, so the
  // *absence* of a format means decimal, not "unknown".
  const numberFormat: WordNumberFormat = formatRaw
    ? (NUMBER_FORMAT_MAP[formatRaw] ?? 'custom')
    : 'decimal';

  if (formatRaw && !NUMBER_FORMAT_MAP[formatRaw]) {
    diagnostics.warn(
      DiagnosticCode.WORD_LIST_NUMBER_FORMAT_APPROXIMATED,
      `Unknown mso-level-number-format "${formatRaw}" on ${listId}:level${level + 1}; the literal level text Word emitted is used instead.`,
      { details: { listId, level: level + 1, format: formatRaw } },
    );
  }

  const type: ListLevelType =
    numberFormat === 'bullet'
      ? 'bullet'
      : numberFormat === 'image'
        ? 'image'
        : numberFormat === 'none'
          ? 'none'
          : 'number';

  const result: WordListLevel = {
    level,
    type,
    numberFormat,
    declarations,
  };
  if (formatRaw) result.numberFormatRaw = formatRaw;

  const levelTextRaw = declarations['mso-level-text'];
  if (levelTextRaw !== undefined) {
    result.levelTextRaw = levelTextRaw;
    result.levelText = decodeMsoLevelText(levelTextRaw);
  }

  const markerFormatting = parseRunFormattingFromCss(declarations);
  if (Object.keys(markerFormatting).length > 0) result.markerFormatting = markerFormatting;

  if (type === 'bullet' && result.levelText) {
    const font = markerFormatting.fontFamily ?? declarations['font-family'];
    const resolved = resolveSymbolGlyph(result.levelText, font);
    result.bulletGlyph = resolved.glyph;
    result.bulletGlyphRaw = resolved.rawGlyph;
    if (resolved.font) result.bulletFont = resolved.font;
    if (resolved.mapped) {
      result.bulletFontMapped = true;
      diagnostics.info(
        DiagnosticCode.WORD_SYMBOL_FONT_MAPPED,
        `Bullet glyph for ${listId}:level${level + 1} was mapped from ${resolved.font ?? 'a symbol font'} to "${resolved.glyph}".`,
        {
          details: {
            listId,
            level: level + 1,
            font: resolved.font ?? '',
            glyph: resolved.glyph,
            codePoint: resolved.codePoint ?? -1,
          },
        },
      );
    } else if (resolved.unmapped) {
      diagnostics.warn(
        DiagnosticCode.WORD_SYMBOL_FONT_UNMAPPED,
        `Bullet glyph for ${listId}:level${level + 1} is a ${resolved.font ?? 'symbol'} font byte with no known Unicode equivalent; it is rendered in its original font.`,
        {
          details: {
            listId,
            level: level + 1,
            font: resolved.font ?? '',
            codePoint: resolved.codePoint ?? -1,
          },
        },
      );
    }
  }

  const startAt = declarations['mso-level-start-at'];
  if (startAt !== undefined) {
    const value = Number.parseInt(startAt, 10);
    if (Number.isFinite(value)) result.startAt = value;
  } else if (type === 'number') {
    result.startAt = 1;
  }

  const lengthOptions = { defaultUnit: 'pt' as const };
  const tabStop = parseWordLength(declarations['mso-level-tab-stop'], lengthOptions);
  if (tabStop) result.tabStop = tabStop;
  const textIndent = parseWordLength(declarations['text-indent'], lengthOptions);
  if (textIndent) result.textIndent = textIndent;
  const marginLeft = parseWordLength(declarations['margin-left'], lengthOptions);
  if (marginLeft) result.marginLeft = marginLeft;
  const indent = parseWordLength(
    declarations['mso-level-indent'] ?? declarations['mso-level-legacy-indent'],
    lengthOptions,
  );
  if (indent) result.indent = indent;

  const position = declarations['mso-level-number-position'];
  if (position) {
    const lower = position.trim().toLowerCase();
    if (lower === 'left' || lower === 'right' || lower === 'center') result.justification = lower;
  }

  const suffix = declarations['mso-level-suffix'] ?? declarations['mso-level-text-suffix'];
  if (suffix) {
    const lower = suffix.trim().toLowerCase();
    if (lower === 'tab' || lower === 'space' || lower === 'nothing') result.suffix = lower;
  } else if (declarations['mso-level-tab-stop'] === 'none') {
    result.suffix = 'space';
  }

  if (/^\s*yes\s*$/i.test(declarations['mso-level-legacy'] ?? '')) result.legacy = true;

  return result;
}

/**
 * Parse the `mso-list` declaration found on a list paragraph.
 *
 * The value is `l0 level1 lfo1` — list id, one-based level, and the list
 * *format override* reference. The lfo matters: two paragraphs can share a
 * numbering definition while belonging to independently numbered lists, and
 * only the lfo distinguishes them.
 *
 * The value `Ignore` appears on the marker span itself and is not a reference.
 */
export function parseListReference(value: string | undefined | null): WordListReference | null {
  if (!value) return null;
  const raw = value.trim();
  if (!raw || /^ignore$/i.test(raw)) return null;

  const listMatch = /\bl(\d+)\b/i.exec(raw);
  const levelMatch = /\blevel(\d+)\b/i.exec(raw);
  const lfoMatch = /\blfo(\d+)\b/i.exec(raw);
  if (!listMatch) return null;

  const wordLevel = levelMatch ? Number.parseInt(levelMatch[1]!, 10) : 1;
  const reference: WordListReference = {
    listId: `l${listMatch[1]}`,
    level: Number.isFinite(wordLevel) && wordLevel > 0 ? wordLevel - 1 : 0,
    raw,
  };
  if (lfoMatch) reference.lfo = `lfo${lfoMatch[1]}`;
  return reference;
}

/**
 * Expand a `mso-level-text` pattern against a stack of counter values.
 *
 * `%1.%2` with counters `[2, 3]` becomes `2.3`. Used when Word gave us the
 * definition but no rendered marker (Word Online omits the marker spans), and
 * by the renderer when it needs the literal text for a level it cannot express
 * as a native CSS counter style.
 */
export function expandLevelText(levelText: string, counters: number[], formats: WordNumberFormat[]): string {
  return levelText.replace(/%(\d)/g, (_match, digit: string) => {
    const index = Number.parseInt(digit, 10) - 1;
    const value = counters[index];
    if (value === undefined) return '';
    return formatNumber(value, formats[index] ?? 'decimal');
  });
}

/** Render a counter value in one of Word's number formats. */
export function formatNumber(value: number, format: WordNumberFormat): string {
  switch (format) {
    case 'decimal-leading-zero':
      return value < 10 ? `0${value}` : String(value);
    case 'lower-alpha':
      return toAlpha(value).toLowerCase();
    case 'upper-alpha':
      return toAlpha(value);
    case 'lower-roman':
      return toRoman(value).toLowerCase();
    case 'upper-roman':
      return toRoman(value);
    case 'ordinal':
      return toOrdinal(value);
    case 'ordinal-text':
      return toOrdinalText(value);
    case 'cardinal-text':
      return toCardinalText(value);
    case 'none':
      return '';
    case 'decimal':
    default:
      return String(value);
  }
}

/** 1 -> A, 26 -> Z, 27 -> AA (Word's bijective base-26). */
export function toAlpha(value: number): string {
  if (value < 1) return String(value);
  let n = value;
  let out = '';
  while (n > 0) {
    const remainder = (n - 1) % 26;
    out = String.fromCharCode(65 + remainder) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

const ROMAN_TABLE: Array<[number, string]> = [
  [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'],
  [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'],
  [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
];

export function toRoman(value: number): string {
  if (value < 1 || value > 3999) return String(value);
  let n = value;
  let out = '';
  for (const [amount, numeral] of ROMAN_TABLE) {
    while (n >= amount) {
      out += numeral;
      n -= amount;
    }
  }
  return out;
}

export function toOrdinal(value: number): string {
  const mod100 = value % 100;
  if (mod100 >= 11 && mod100 <= 13) return `${value}th`;
  switch (value % 10) {
    case 1:
      return `${value}st`;
    case 2:
      return `${value}nd`;
    case 3:
      return `${value}rd`;
    default:
      return `${value}th`;
  }
}

const ONES = [
  '', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine',
  'ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen',
  'seventeen', 'eighteen', 'nineteen',
];
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];
const ORDINAL_ONES = [
  '', 'first', 'second', 'third', 'fourth', 'fifth', 'sixth', 'seventh', 'eighth', 'ninth',
  'tenth', 'eleventh', 'twelfth', 'thirteenth', 'fourteenth', 'fifteenth', 'sixteenth',
  'seventeenth', 'eighteenth', 'nineteenth',
];
const ORDINAL_TENS = [
  '', '', 'twentieth', 'thirtieth', 'fortieth', 'fiftieth', 'sixtieth', 'seventieth',
  'eightieth', 'ninetieth',
];

export function toCardinalText(value: number): string {
  if (value < 0 || value > 9999) return String(value);
  if (value === 0) return 'zero';
  if (value < 20) return ONES[value]!;
  if (value < 100) {
    const tens = Math.floor(value / 10);
    const ones = value % 10;
    return ones === 0 ? TENS[tens]! : `${TENS[tens]}-${ONES[ones]}`;
  }
  if (value < 1000) {
    const hundreds = Math.floor(value / 100);
    const rest = value % 100;
    return rest === 0
      ? `${ONES[hundreds]} hundred`
      : `${ONES[hundreds]} hundred ${toCardinalText(rest)}`;
  }
  const thousands = Math.floor(value / 1000);
  const rest = value % 1000;
  return rest === 0
    ? `${toCardinalText(thousands)} thousand`
    : `${toCardinalText(thousands)} thousand ${toCardinalText(rest)}`;
}

export function toOrdinalText(value: number): string {
  if (value < 0 || value > 9999) return toOrdinal(value);
  if (value < 20) return ORDINAL_ONES[value] ?? toOrdinal(value);
  if (value < 100) {
    const tens = Math.floor(value / 10);
    const ones = value % 10;
    return ones === 0 ? ORDINAL_TENS[tens]! : `${TENS[tens]}-${ORDINAL_ONES[ones]}`;
  }
  const cardinal = toCardinalText(value);
  const lastSpace = cardinal.lastIndexOf(' ');
  if (lastSpace === -1) return `${cardinal}th`;
  const head = cardinal.slice(0, lastSpace);
  const tail = cardinal.slice(lastSpace + 1);
  const tailIndex = ONES.indexOf(tail);
  if (tailIndex > 0) return `${head} ${ORDINAL_ONES[tailIndex]}`;
  const tensIndex = TENS.indexOf(tail);
  if (tensIndex > 0) return `${head} ${ORDINAL_TENS[tensIndex]}`;
  return `${cardinal}th`;
}
