import { describe, expect, it } from 'vitest';
import { tokenizeCss, parseInlineStyle, declarationsToMap } from '../../word/WordCssTokenizer.js';
import { parseWordStyleSheet } from '../../word/WordStyleParser.js';
import { parseListReference, expandLevelText, formatNumber, toRoman, toAlpha, toOrdinalText } from '../../word/WordListStyleParser.js';
import { parseWordLength, parseBoxShorthand, lengthToCss } from '../../word/WordLengthParser.js';
import { parseWordColor } from '../../word/WordColorParser.js';
import { decodeMsoLevelText, resolveSymbolGlyph } from '../../word/WordSymbolFonts.js';
import { DiagnosticCollector } from '../../diagnostics/DiagnosticCollector.js';

describe('CSS tokenizer', () => {
  it('keeps at-rules a conforming CSS parser would discard', () => {
    const rules = tokenizeCss(`
      @list l0 {mso-list-id:123;}
      @list l0:level1 {mso-level-number-format:bullet; mso-level-text:\\F0B7;}
      p.MsoNormal {margin:0in;}
    `);
    expect(rules).toHaveLength(3);
    expect(rules[0]!.atName).toBe('list');
    expect(rules[0]!.prelude).toBe('l0');
    expect(rules[1]!.prelude).toBe('l0:level1');
    expect(declarationsToMap(rules[1]!.declarations)['mso-level-text']).toBe('\\F0B7');
  });

  it('strips the HTML comment wrapper Word puts inside <style>', () => {
    const rules = tokenizeCss('<!--\n p.MsoNormal {margin:0in;}\n-->');
    expect(rules).toHaveLength(1);
    expect(rules[0]!.selector).toBe('p.MsoNormal');
  });

  it('does not split a value on a semicolon inside a string', () => {
    const rules = tokenizeCss(`p {font:7.0pt "Times; New Roman"; color:red;}`);
    const map = declarationsToMap(rules[0]!.declarations);
    expect(map['font']).toBe('7.0pt "Times; New Roman"');
    expect(map['color']).toBe('red');
  });

  it('does not split a declaration on a colon inside a value', () => {
    const map = parseInlineStyle(`background:url(data:image/png;base64,AAA);color:blue`);
    expect(map['color']).toBe('blue');
    expect(map['background']).toContain('data:image/png');
  });

  it('survives a truncated stylesheet without throwing', () => {
    expect(() => tokenizeCss('p.MsoNormal {margin:0in; font-fam')).not.toThrow();
    expect(() => tokenizeCss('@list l0 {')).not.toThrow();
    expect(() => tokenizeCss('}}}{{{')).not.toThrow();
  });

  it('flattens conditional groups', () => {
    const sheet = parseWordStyleSheet(
      ['@media print { p.MsoNormal {margin:1in;} }'],
      new DiagnosticCollector(),
    );
    expect(sheet.styles['msonormal']).toBeDefined();
  });
});

describe('Word stylesheet', () => {
  const sheet = parseWordStyleSheet(
    [
      `@font-face {font-family:Wingdings; mso-font-charset:2; mso-generic-font-family:decorative;}
       @font-face {font-family:Calibri; mso-font-charset:0; mso-generic-font-family:swiss;}
       p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
         {mso-style-name:"List Paragraph"; margin-left:.5in; font-size:11.0pt;}
       span.EmphasisChar {mso-style-name:"Emphasis Char"; font-style:italic;}
       p.MsoHeading1 {mso-style-name:"Heading 1"; mso-outline-level:1; font-size:16.0pt;}
       @page WordSection1 {size:8.5in 11.0in; margin:1.0in;}
       @list l0 {mso-list-id:99;}
       @list l0:level1 {mso-level-number-format:roman-upper; mso-level-text:"%1\\."; mso-level-start-at:3;}`,
    ],
    new DiagnosticCollector(),
  );

  it('recovers the human style name from mso-style-name', () => {
    expect(sheet.styles['msolistparagraph']!.name).toBe('List Paragraph');
    expect(sheet.styles['msoheading1']!.name).toBe('Heading 1');
  });

  it('registers a style under every selector Word declared it with', () => {
    expect(sheet.styles['msolistparagraph']!.selectors).toEqual([
      'p.MsoListParagraph',
      'li.MsoListParagraph',
      'div.MsoListParagraph',
    ]);
  });

  it('classifies span-only rules as character styles', () => {
    expect(sheet.styles['emphasischar']!.type).toBe('character');
    expect(sheet.styles['emphasischar']!.run.italic).toBe(true);
  });

  it('marks symbol fonts from mso-font-charset:2', () => {
    expect(sheet.fonts['wingdings']!.isSymbolFont).toBe(true);
    expect(sheet.fonts['calibri']!.isSymbolFont).toBe(false);
  });

  it('parses @page blocks', () => {
    expect(sheet.pages['WordSection1']!['size']).toBe('8.5in 11.0in');
  });

  it('parses list levels including start-at', () => {
    const level = sheet.lists['l0']!.levels[0]!;
    expect(level.numberFormat).toBe('upper-roman');
    expect(level.levelText).toBe('%1.');
    expect(level.startAt).toBe(3);
  });

  it('treats an absent mso-level-number-format as decimal, not unknown', () => {
    const plain = parseWordStyleSheet(
      ['@list l1:level1 {mso-level-tab-stop:.5in;}'],
      new DiagnosticCollector(),
    );
    expect(plain.lists['l1']!.levels[0]!.numberFormat).toBe('decimal');
    expect(plain.lists['l1']!.levels[0]!.type).toBe('number');
  });
});

describe('mso-list references', () => {
  it('parses list id, level and lfo', () => {
    expect(parseListReference('l0 level3 lfo2')).toEqual({
      listId: 'l0',
      level: 2,
      lfo: 'lfo2',
      raw: 'l0 level3 lfo2',
    });
  });

  it('treats Ignore as not a reference', () => {
    expect(parseListReference('Ignore')).toBeNull();
    expect(parseListReference('')).toBeNull();
    expect(parseListReference(undefined)).toBeNull();
  });

  it('defaults to level 1 when the level is omitted', () => {
    expect(parseListReference('l5 lfo1')?.level).toBe(0);
  });
});

describe('level text expansion', () => {
  it('expands a hierarchical pattern from counter values', () => {
    expect(expandLevelText('%1.%2', [1, 4], ['decimal', 'decimal'])).toBe('1.4');
    expect(expandLevelText('%1.', [7], ['upper-roman'])).toBe('VII.');
    expect(expandLevelText('(%1)', [3], ['lower-alpha'])).toBe('(c)');
    expect(expandLevelText('%1.%2.%3', [2, 1, 5], ['decimal', 'decimal', 'decimal'])).toBe('2.1.5');
  });

  it('does not invent a scheme of its own', () => {
    // The exact failure the project exists to prevent: `%1.%2` must never
    // become "Section 1.01" or "Article I".
    const result = expandLevelText('%1.%2', [1, 1], ['decimal', 'decimal']);
    expect(result).toBe('1.1');
    expect(result).not.toMatch(/Section|Article/);
  });
});

describe('number formatting', () => {
  it('formats every Word number format the engine claims to support', () => {
    expect(formatNumber(4, 'decimal')).toBe('4');
    expect(formatNumber(4, 'decimal-leading-zero')).toBe('04');
    expect(formatNumber(4, 'lower-alpha')).toBe('d');
    expect(formatNumber(27, 'upper-alpha')).toBe('AA');
    expect(formatNumber(9, 'lower-roman')).toBe('ix');
    expect(formatNumber(1990, 'upper-roman')).toBe('MCMXC');
    expect(formatNumber(22, 'ordinal')).toBe('22nd');
    expect(formatNumber(3, 'ordinal-text')).toBe('third');
    expect(formatNumber(42, 'cardinal-text')).toBe('forty-two');
  });

  it('uses bijective base-26 for alpha, as Word does', () => {
    expect(toAlpha(1)).toBe('A');
    expect(toAlpha(26)).toBe('Z');
    expect(toAlpha(27)).toBe('AA');
    expect(toAlpha(52)).toBe('AZ');
  });

  it('formats roman numerals subtractively', () => {
    expect(toRoman(4)).toBe('IV');
    expect(toRoman(40)).toBe('XL');
    expect(toRoman(3999)).toBe('MMMCMXCIX');
  });

  it('formats ordinal text', () => {
    expect(toOrdinalText(1)).toBe('first');
    expect(toOrdinalText(21)).toBe('twenty-first');
    expect(toOrdinalText(30)).toBe('thirtieth');
  });
});

describe('length parsing', () => {
  it('converts every unit Word emits to the same canonical value', () => {
    expect(parseWordLength('1in')!.px).toBe(96);
    expect(parseWordLength('72pt')!.px).toBe(96);
    expect(parseWordLength('2.54cm')!.px).toBeCloseTo(96, 2);
    expect(parseWordLength('25.4mm')!.px).toBeCloseTo(96, 2);
    expect(parseWordLength('96px')!.px).toBe(96);
    expect(parseWordLength('1in')!.twips).toBe(1440);
  });

  it('parses the abbreviated decimals Word writes', () => {
    expect(parseWordLength('.25in')!.px).toBe(24);
    expect(parseWordLength('-.5in')!.px).toBe(-48);
    expect(parseWordLength('11.0pt')!.px).toBeCloseTo(14.6667, 3);
  });

  it('keeps the original literal', () => {
    expect(parseWordLength('-.25in')!.raw).toBe('-.25in');
  });

  it('returns undefined rather than zero for a non-length', () => {
    expect(parseWordLength('auto')).toBeUndefined();
    expect(parseWordLength('')).toBeUndefined();
    expect(parseWordLength(undefined)).toBeUndefined();
    expect(parseWordLength('inherit')).toBeUndefined();
    expect(parseWordLength('0')!.px).toBe(0);
  });

  it('parses box shorthands', () => {
    const box = parseBoxShorthand('0in 5.4pt 0in 5.4pt')!;
    expect(box.top!.px).toBe(0);
    expect(box.right!.px).toBeCloseTo(7.2, 2);
    expect(box.left!.px).toBeCloseTo(7.2, 2);
  });

  it('renders back to points by default', () => {
    expect(lengthToCss(parseWordLength('.5in')!)).toBe('36pt');
  });
});

describe('colour parsing', () => {
  it('resolves Word system colours', () => {
    expect(parseWordColor('windowtext')!.hex).toBe('#000000');
    expect(parseWordColor('window')!.hex).toBe('#ffffff');
  });

  it('resolves named and hex colours', () => {
    expect(parseWordColor('red')!.hex).toBe('#ff0000');
    expect(parseWordColor('#1F497D')!.hex).toBe('#1f497d');
    expect(parseWordColor('#abc')!.hex).toBe('#aabbcc');
    expect(parseWordColor('rgb(31, 73, 125)')!.hex).toBe('#1f497d');
  });

  it('keeps the original literal', () => {
    expect(parseWordColor('windowtext')!.raw).toBe('windowtext');
  });

  it('treats auto as absent, not as black', () => {
    expect(parseWordColor('auto')).toBeUndefined();
  });
});

describe('symbol fonts', () => {
  it('decodes mso-level-text escapes', () => {
    expect(decodeMsoLevelText('\\F0B7')).toBe('\uf0b7');
    expect(decodeMsoLevelText('"%1\\."')).toBe('%1.');
    expect(decodeMsoLevelText('"%1\\.%2"')).toBe('%1.%2');
    expect(decodeMsoLevelText('o')).toBe('o');
  });

  it('maps the three Word default bullets to Unicode', () => {
    expect(resolveSymbolGlyph('\uf0b7', 'Symbol').glyph).toBe('\u2022');
    expect(resolveSymbolGlyph('\u00b7', 'Symbol').glyph).toBe('\u2022');
    expect(resolveSymbolGlyph('o', 'Courier New').glyph).toBe('o');
    expect(resolveSymbolGlyph('\uf0a7', 'Wingdings').glyph).toBe('\u25aa');
    expect(resolveSymbolGlyph('\u00a7', 'Wingdings').glyph).toBe('\u25aa');
  });

  it('preserves the raw byte and the font alongside the mapping', () => {
    const result = resolveSymbolGlyph('\uf0b7', 'Symbol');
    expect(result.rawGlyph).toBe('\uf0b7');
    expect(result.font).toBe('Symbol');
    expect(result.mapped).toBe(true);
    expect(result.codePoint).toBe(0xb7);
  });

  it('flags an unmapped symbol byte instead of guessing', () => {
    const result = resolveSymbolGlyph('\uf001', 'Wingdings');
    expect(result.mapped).toBe(false);
    expect(result.unmapped).toBe(true);
    expect(result.glyph).toBe('\uf001');
    expect(result.font).toBe('Wingdings');
  });

  it('leaves an ordinary character alone', () => {
    const result = resolveSymbolGlyph('\u2022', 'Calibri');
    expect(result.glyph).toBe('\u2022');
    expect(result.mapped).toBe(false);
    expect(result.unmapped).toBe(false);
  });
});
