import type { Length } from '../model/Style.js';

/**
 * The one place in the engine where physical units are converted.
 *
 * Word writes lengths in whatever unit the user's Word is configured for:
 * `.25in`, `18.0pt`, `1.27cm`, `36.0px`, and occasionally bare numbers. Every
 * one of those has to end up comparable, so the model stores a canonical
 * `px` value *and* the exact literal Word wrote. Nothing downstream may
 * re-derive a conversion; it calls this.
 *
 * Canonical unit: CSS pixel (96 px == 1 in), which is also the unit the
 * renderer emits. Twips (1/1440 in) are carried alongside because that is
 * Word's own internal unit and round-tripping through it loses nothing.
 */

/** CSS reference pixels per inch. */
export const PX_PER_INCH = 96;
/** Points per inch. */
export const PT_PER_INCH = 72;
/** Twentieths of a point per inch — Word's native unit. */
export const TWIPS_PER_INCH = 1440;
/** Centimetres per inch. */
export const CM_PER_INCH = 2.54;

const PX_PER_UNIT: Record<string, number> = {
  px: 1,
  pt: PX_PER_INCH / PT_PER_INCH, // 1.3333…
  pc: PX_PER_INCH / 6, // 1 pica == 12 pt
  in: PX_PER_INCH,
  cm: PX_PER_INCH / CM_PER_INCH,
  mm: PX_PER_INCH / (CM_PER_INCH * 10),
  q: PX_PER_INCH / (CM_PER_INCH * 40),
  twip: PX_PER_INCH / TWIPS_PER_INCH,
};

const LENGTH_PATTERN = /^\s*([+-]?(?:\d+\.?\d*|\.\d+))\s*([a-z%]*)\s*$/i;

export interface ParseLengthOptions {
  /**
   * Unit assumed when the literal has no unit suffix. Word occasionally emits
   * bare numbers in HTML attributes (`width=100`), which are pixels; inside
   * CSS a bare non-zero number is invalid but Word still produces it for
   * `mso-` properties, where points are the intended reading.
   */
  defaultUnit?: keyof typeof PX_PER_UNIT;
  /** Font size in px used to resolve `em`/`ex`/`rem` values. */
  fontSizePx?: number;
  /** Reference length in px used to resolve percentages. */
  percentBasisPx?: number;
}

/**
 * Parse a Word length literal into a canonical {@link Length}.
 *
 * Returns `undefined` for anything that is not a length — `auto`, `inherit`,
 * empty strings, keyword values — so callers can distinguish "absent" from
 * "zero" without a sentinel.
 */
export function parseWordLength(
  input: string | number | null | undefined,
  options: ParseLengthOptions = {},
): Length | undefined {
  if (input === null || input === undefined) return undefined;

  if (typeof input === 'number') {
    if (!Number.isFinite(input)) return undefined;
    const unit = options.defaultUnit ?? 'px';
    return makeLength(input * (PX_PER_UNIT[unit] ?? 1), String(input));
  }

  const raw = input.trim();
  if (raw.length === 0) return undefined;

  const lowered = raw.toLowerCase();
  if (lowered === 'auto' || lowered === 'inherit' || lowered === 'initial' || lowered === 'none') {
    return undefined;
  }

  const match = LENGTH_PATTERN.exec(raw);
  if (!match) return undefined;

  const value = Number.parseFloat(match[1] ?? '');
  if (!Number.isFinite(value)) return undefined;
  const unit = (match[2] ?? '').toLowerCase();

  if (unit === '') {
    // A bare `0` is zero in any unit. Anything else takes the caller's default.
    if (value === 0) return makeLength(0, raw);
    const defaultUnit = options.defaultUnit ?? 'px';
    return makeLength(value * (PX_PER_UNIT[defaultUnit] ?? 1), raw);
  }

  if (unit === '%') {
    if (options.percentBasisPx === undefined) return undefined;
    return makeLength((value / 100) * options.percentBasisPx, raw);
  }

  if (unit === 'em' || unit === 'rem') {
    const basis = options.fontSizePx ?? 16;
    return makeLength(value * basis, raw);
  }
  if (unit === 'ex') {
    const basis = options.fontSizePx ?? 16;
    return makeLength(value * basis * 0.5, raw);
  }

  const factor = PX_PER_UNIT[unit];
  if (factor === undefined) return undefined;
  return makeLength(value * factor, raw);
}

/** Parse a percentage literal (`75%`, `100.0%`) into a number, or undefined. */
export function parsePercent(input: string | null | undefined): number | undefined {
  if (!input) return undefined;
  const match = /^\s*([+-]?(?:\d+\.?\d*|\.\d+))\s*%\s*$/.exec(input);
  if (!match) return undefined;
  const value = Number.parseFloat(match[1] ?? '');
  return Number.isFinite(value) ? value : undefined;
}

function makeLength(px: number, raw: string): Length {
  const rounded = roundTo(px, 4);
  return {
    px: rounded,
    twips: Math.round((rounded / PX_PER_INCH) * TWIPS_PER_INCH),
    raw,
  };
}

/** Build a Length from a px number (used when the engine synthesises a value). */
export function lengthFromPx(px: number, raw?: string): Length {
  return makeLength(px, raw ?? `${roundTo(px, 4)}px`);
}

/** Build a Length from points. */
export function lengthFromPt(pt: number): Length {
  return makeLength(pt * PX_PER_UNIT.pt!, `${roundTo(pt, 4)}pt`);
}

/** Build a Length from twips. */
export function lengthFromTwips(twips: number): Length {
  return makeLength(twips * PX_PER_UNIT.twip!, `${twips}twip`);
}

/** Render a Length back to CSS, preferring `pt` because that is what Word means. */
export function lengthToCss(length: Length, unit: 'px' | 'pt' | 'in' = 'pt'): string {
  switch (unit) {
    case 'px':
      return `${roundTo(length.px, 2)}px`;
    case 'in':
      return `${roundTo(length.px / PX_PER_INCH, 4)}in`;
    case 'pt':
    default:
      return `${roundTo((length.px / PX_PER_INCH) * PT_PER_INCH, 2)}pt`;
  }
}

export function pxToPt(px: number): number {
  return roundTo((px / PX_PER_INCH) * PT_PER_INCH, 4);
}

export function ptToPx(pt: number): number {
  return roundTo((pt / PT_PER_INCH) * PX_PER_INCH, 4);
}

export function roundTo(value: number, digits: number): number {
  const factor = 10 ** digits;
  return Math.round(value * factor + Number.EPSILON) / factor;
}

/**
 * Parse the four-value CSS shorthand Word uses for `margin`, `padding` and
 * `mso-padding-alt`, returning the sides in canonical units.
 */
export function parseBoxShorthand(
  value: string | null | undefined,
  options: ParseLengthOptions = {},
): { top?: Length; right?: Length; bottom?: Length; left?: Length } | undefined {
  if (!value) return undefined;
  const parts = value.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return undefined;
  const lengths = parts.map((p) => parseWordLength(p, options));
  const [a, b, c, d] = lengths;
  const box =
    lengths.length === 1
      ? { top: a, right: a, bottom: a, left: a }
      : lengths.length === 2
        ? { top: a, right: b, bottom: a, left: b }
        : lengths.length === 3
          ? { top: a, right: b, bottom: c, left: b }
          : { top: a, right: b, bottom: c, left: d };
  const result: { top?: Length; right?: Length; bottom?: Length; left?: Length } = {};
  if (box.top) result.top = box.top;
  if (box.right) result.right = box.right;
  if (box.bottom) result.bottom = box.bottom;
  if (box.left) result.left = box.left;
  return Object.keys(result).length > 0 ? result : undefined;
}
