/**
 * Symbol-font code page mapping.
 *
 * Word's default bullets are not Unicode bullets. They are byte values in a
 * legacy symbol font:
 *
 *   level 1  Symbol       0xB7  which *draws* as a bullet   -> U+2022 •
 *   level 2  Courier New  0x6F  which really is the letter o -> o
 *   level 3  Wingdings    0xA7  which *draws* as a small square -> U+25AA ▪
 *
 * In the clipboard those bytes reach us two ways:
 *
 *   1. In `mso-level-text` as an escape: `mso-level-text:"\F0B7"` — the byte
 *      lifted into the Unicode private use area (0xF000 + byte).
 *   2. In the rendered marker span as the literal character the byte maps to
 *      under the document's ANSI code page: `<span style='mso-list:Ignore'>·`
 *      (U+00B7 MIDDLE DOT) — which is *not* what the reader saw.
 *
 * Taking either at face value is how a paste ends up showing `·` or `§`
 * instead of `•` and `▪`. So the engine maps the byte through the font's own
 * code page, and keeps the raw byte and the font name next to the result so no
 * information is destroyed and the renderer can still fall back to rendering
 * the original glyph in the original font.
 */

/** Fonts whose byte values are glyph indices rather than Unicode code points. */
const SYMBOL_FONT_NAMES = new Set([
  'symbol',
  'wingdings',
  'wingdings 2',
  'wingdings 3',
  'webdings',
  'monotype sorts',
  'zapfdingbats',
  'zapf dingbats',
  'marlett',
  'ms outlook',
  'mt extra',
]);

export function isSymbolFont(fontFamily: string | undefined): boolean {
  if (!fontFamily) return false;
  return SYMBOL_FONT_NAMES.has(normaliseFontName(fontFamily));
}

/** Strip quotes and take the first family from a font stack. */
export function normaliseFontName(fontFamily: string): string {
  const first = fontFamily.split(',')[0] ?? fontFamily;
  return first.trim().replace(/^["']|["']$/g, '').toLowerCase();
}

/**
 * Adobe Symbol encoding, high range (0xA0–0xFF) plus the Greek letters.
 * Only the values Word can actually emit as a list bullet or inline symbol are
 * needed in practice, but the full table costs nothing and prevents a future
 * "why did this one turn into a box" bug.
 */
const SYMBOL_MAP: Record<number, string> = {
  0x22: '∀', 0x24: '∃', 0x27: '∋', 0x2a: '∗', 0x2d: '−',
  0x40: '≅',
  0x41: 'Α', 0x42: 'Β', 0x43: 'Χ', 0x44: 'Δ', 0x45: 'Ε',
  0x46: 'Φ', 0x47: 'Γ', 0x48: 'Η', 0x49: 'Ι', 0x4a: 'ϑ',
  0x4b: 'Κ', 0x4c: 'Λ', 0x4d: 'Μ', 0x4e: 'Ν', 0x4f: 'Ο',
  0x50: 'Π', 0x51: 'Θ', 0x52: 'Ρ', 0x53: 'Σ', 0x54: 'Τ',
  0x55: 'Υ', 0x56: 'ς', 0x57: 'Ω', 0x58: 'Ξ', 0x59: 'Ψ',
  0x5a: 'Ζ', 0x5c: '∴', 0x5e: '⊥', 0x60: '‾',
  0x61: 'α', 0x62: 'β', 0x63: 'χ', 0x64: 'δ', 0x65: 'ε',
  0x66: 'φ', 0x67: 'γ', 0x68: 'η', 0x69: 'ι', 0x6a: 'ϕ',
  0x6b: 'κ', 0x6c: 'λ', 0x6d: 'μ', 0x6e: 'ν', 0x6f: 'ο',
  0x70: 'π', 0x71: 'θ', 0x72: 'ρ', 0x73: 'σ', 0x74: 'τ',
  0x75: 'υ', 0x76: 'ϖ', 0x77: 'ω', 0x78: 'ξ', 0x79: 'ψ',
  0x7a: 'ζ', 0x7e: '∼',
  0xa0: '€', 0xa1: 'ϒ', 0xa2: '′', 0xa3: '≤', 0xa4: '⁄',
  0xa5: '∞', 0xa6: 'ƒ', 0xa7: '♣', 0xa8: '♦', 0xa9: '♥',
  0xaa: '♠', 0xab: '↔', 0xac: '←', 0xad: '↑', 0xae: '→',
  0xaf: '↓',
  0xb0: '°', 0xb1: '±', 0xb2: '″', 0xb3: '≥', 0xb4: '×',
  0xb5: '∝', 0xb6: '∂',
  // The Word default bullet.
  0xb7: '•',
  0xb8: '÷', 0xb9: '≠', 0xba: '≡', 0xbb: '≈', 0xbc: '…',
  0xbd: '⏐', 0xbe: '⎯', 0xbf: '↵',
  0xc0: 'ℵ', 0xc1: 'ℑ', 0xc2: 'ℜ', 0xc3: '℘', 0xc4: '⊗',
  0xc5: '⊕', 0xc6: '∅', 0xc7: '∩', 0xc8: '∪', 0xc9: '⊃',
  0xca: '⊇', 0xcb: '⊄', 0xcc: '⊂', 0xcd: '⊆', 0xce: '∈',
  0xcf: '∉',
  0xd0: '∠', 0xd1: '∇', 0xd2: '®', 0xd3: '©', 0xd4: '™',
  0xd5: '∏', 0xd6: '√', 0xd7: '⋅', 0xd8: '¬', 0xd9: '∧',
  0xda: '∨', 0xdb: '⇔', 0xdc: '⇐', 0xdd: '⇑', 0xde: '⇒',
  0xdf: '⇓',
  0xe0: '◊', 0xe1: '〈', 0xe2: '®', 0xe3: '©', 0xe4: '™',
  0xe5: '∑', 0xf1: '〉', 0xf2: '∫', 0xf3: '⌠', 0xf5: '⌡',
};

/**
 * Wingdings. The full font is 224 pictographs; mapped here are the values Word
 * uses for list bullets and the handful of dingbats that survive a paste with
 * a sensible Unicode equivalent. Anything not listed keeps its raw glyph and
 * its font, and raises WORD_SYMBOL_FONT_UNMAPPED.
 */
const WINGDINGS_MAP: Record<number, string> = {
  0x28: '☎', // telephone
  0x2a: '✉', // envelope
  0x3c: '⌛', // hourglass
  0x40: '⌨', // keyboard
  0x46: '✎', // pencil
  0x4a: '☺', // smiling face
  0x4b: '☼', // neutral face
  0x4c: '☹', // frowning face
  0x4e: '☠', // skull
  0x4f: '⚐', // flag
  0x54: '✈', // airplane
  0x57: '☀', // sun
  0x58: '☁', // cloud
  0x5a: '❄', // snowflake
  0x5b: '✝', // latin cross
  0x6c: '●', // BLACK CIRCLE  — a common Word bullet
  0x6d: '❍', // shadowed white circle
  0x6e: '■', // BLACK SQUARE  — a common Word bullet
  0x6f: '□', // white square
  0x70: '❑', // lower-right shadowed white square
  0x71: '❒', // upper-right shadowed white square
  0x72: '☐', // ballot box
  0x73: '◆', // black diamond
  0x74: '❖', // black diamond minus white X
  0x75: '◇', // white diamond
  0x76: '♦', // black diamond suit
  0x77: '◈', // white diamond containing black
  0x78: '⌧', // X in a rectangle
  0x9f: '•', // bullet
  0xa4: '●', // black circle
  0xa5: '○', // white circle
  // The Word level-3 default bullet.
  0xa7: '▪', // BLACK SMALL SQUARE
  0xa8: '▫', // white small square
  0xa9: '❑',
  0xab: '✦', // black four-pointed star
  0xac: '★', // black star
  0xd8: '➔', // heavy wide-headed rightwards arrow
  0xfc: '✔', // heavy check mark
  0xfd: '✘', // heavy ballot X
  0xfe: '☑', // ballot box with check
  0xff: '☒', // ballot box with X
};

/** Wingdings 2 — only the bullet-capable values. */
const WINGDINGS2_MAP: Record<number, string> = {
  0x97: '●',
  0x98: '○',
  0x9f: '▪',
  0xa1: '■',
  0xa2: '□',
  0xb7: '•',
  0xd8: '◆',
};

/** Wingdings 3 — arrows only; used for custom arrow bullets. */
const WINGDINGS3_MAP: Record<number, string> = {
  0x75: '▲',
  0x76: '▼',
  0x77: '◄',
  0x78: '►',
  0x70: '↑',
  0x71: '↓',
  0x74: '→',
};

/** Webdings — only what is plausibly used as a bullet. */
const WEBDINGS_MAP: Record<number, string> = {
  0x6e: '■',
  0x6f: '□',
};

/** Monotype Sorts / ZapfDingbats — arrow and check bullets. */
const DINGBATS_MAP: Record<number, string> = {
  0x6c: '●',
  0x6e: '■',
  0x75: '❖',
  0x76: '◆',
  0xa4: '•',
  0xd8: '➜',
  0xfc: '✔',
};

const FONT_MAPS: Record<string, Record<number, string>> = {
  symbol: SYMBOL_MAP,
  wingdings: WINGDINGS_MAP,
  'wingdings 2': WINGDINGS2_MAP,
  'wingdings 3': WINGDINGS3_MAP,
  webdings: WEBDINGS_MAP,
  'monotype sorts': DINGBATS_MAP,
  zapfdingbats: DINGBATS_MAP,
  'zapf dingbats': DINGBATS_MAP,
};

export interface SymbolMappingResult {
  /** The best Unicode representation of the glyph. */
  glyph: string;
  /** The glyph exactly as it arrived, before any mapping. */
  rawGlyph: string;
  /** The font required to render `rawGlyph` correctly, when it is a symbol font. */
  font?: string;
  /** True when a code page mapping was applied. */
  mapped: boolean;
  /** True when the font is a symbol font but the byte had no known mapping. */
  unmapped: boolean;
  /** The byte value that was looked up, when one was derived. */
  codePoint?: number;
}

/**
 * Resolve a bullet glyph to its Unicode equivalent.
 *
 * `glyph` may be a literal character from the marker span, a private-use
 * character from an `mso-level-text` escape, or a plain letter (`o`) that needs
 * no mapping at all.
 */
export function resolveSymbolGlyph(glyph: string, fontFamily?: string): SymbolMappingResult {
  const result: SymbolMappingResult = { glyph, rawGlyph: glyph, mapped: false, unmapped: false };
  if (fontFamily) result.font = fontFamily;
  if (glyph.length === 0) return result;

  const codePoint = glyph.codePointAt(0);
  if (codePoint === undefined) return result;

  // Private use area: Word lifted a font byte into U+F000..U+F0FF. This is
  // unambiguous evidence of a symbol font even when no font name reached us.
  let byte: number | undefined;
  if (codePoint >= 0xf000 && codePoint <= 0xf0ff) {
    byte = codePoint - 0xf000;
  } else if (codePoint <= 0xff) {
    byte = codePoint;
  }
  result.codePoint = byte;

  const fontKey = fontFamily ? normaliseFontName(fontFamily) : undefined;
  const map = fontKey ? FONT_MAPS[fontKey] : undefined;

  if (!map) {
    if (byte !== undefined && codePoint >= 0xf000) {
      // Private use with no font name: Symbol is by far the most likely source
      // and is the only font Word uses for the default bullet, but say so.
      const fallback = SYMBOL_MAP[byte];
      if (fallback) {
        result.glyph = fallback;
        result.mapped = true;
        if (!result.font) result.font = 'Symbol';
        return result;
      }
      result.unmapped = true;
    }
    return result;
  }

  if (byte === undefined) return result;
  const mapped = map[byte];
  if (mapped) {
    result.glyph = mapped;
    result.mapped = true;
    return result;
  }
  // A symbol font byte we have no table entry for. Keep the raw glyph and the
  // font so it still draws correctly, and let the caller diagnose it.
  result.unmapped = true;
  return result;
}

/**
 * Decode the escape syntax Word uses inside `mso-level-text`.
 *
 * Word writes CSS string escapes: `"\F0B7"`, `"\F0A7"`, `"%1\."`. The escapes
 * are hex code points (private use for symbol fonts), and a backslash before a
 * literal is just an escape of that literal.
 */
export function decodeMsoLevelText(raw: string): string {
  let value = raw.trim();
  // Strip the surrounding quotes Word always writes.
  if (
    (value.startsWith('"') && value.endsWith('"')) ||
    (value.startsWith("'") && value.endsWith("'"))
  ) {
    value = value.slice(1, -1);
  }
  let out = '';
  for (let i = 0; i < value.length; i++) {
    const ch = value[i];
    if (ch !== '\\') {
      out += ch;
      continue;
    }
    // CSS escape: up to six hex digits, optionally followed by one whitespace.
    const rest = value.slice(i + 1);
    const hex = /^([0-9a-fA-F]{1,6})\s?/.exec(rest);
    if (hex) {
      const code = Number.parseInt(hex[1]!, 16);
      if (Number.isFinite(code) && code > 0 && code <= 0x10ffff) {
        out += String.fromCodePoint(code);
      }
      i += hex[0]!.length;
      continue;
    }
    // Escaped literal, e.g. `\.` -> `.`
    const next = value[i + 1];
    if (next !== undefined) {
      out += next;
      i += 1;
    }
  }
  return out;
}
