import type {
  Border,
  Borders,
  Color,
  Length,
  ParagraphFormatting,
  RunFormatting,
  Shading,
} from '../model/Style.js';
import { lengthToCss, roundTo } from '../word/WordLengthParser.js';

/**
 * Model formatting -> clean browser CSS.
 *
 * This is the only place Word semantics become CSS, and it is deliberately a
 * *translation* rather than a pass-through. None of Word's `mso-` properties
 * reach the output; every declaration emitted is standard CSS that a browser,
 * an editor, and a downloaded HTML file all understand the same way.
 *
 * Lengths come out in points. That is what Word means, it round-trips back
 * into Word cleanly, and it keeps the output stable across devices in a way
 * pixels do not.
 */

export interface StyleRenderOptions {
  /** Unit for emitted lengths. Default `pt`. */
  unit?: 'pt' | 'px' | 'in';
  /** Emit `data-word-*` attributes carrying the original Word values. */
  includeWordMetadata?: boolean;
}

/** An ordered list of `property: value` pairs. */
export class StyleBuilder {
  private readonly declarations: string[] = [];
  private readonly seen = new Set<string>();

  set(property: string, value: string | undefined | null): this {
    if (value === undefined || value === null || value === '') return this;
    if (this.seen.has(property)) {
      const index = this.declarations.findIndex((d) => d.startsWith(`${property}:`));
      if (index >= 0) this.declarations[index] = `${property}:${value}`;
      return this;
    }
    this.seen.add(property);
    this.declarations.push(`${property}:${value}`);
    return this;
  }

  has(property: string): boolean {
    return this.seen.has(property);
  }

  get isEmpty(): boolean {
    return this.declarations.length === 0;
  }

  toString(): string {
    return this.declarations.join(';');
  }
}

const DEFAULT_UNIT: StyleRenderOptions['unit'] = 'pt';

export function renderLength(length: Length, options: StyleRenderOptions = {}): string {
  return lengthToCss(length, options.unit ?? DEFAULT_UNIT);
}

export function renderColor(color: Color): string {
  return color.hex;
}

/** Character formatting as inline CSS. */
export function renderRunStyle(
  formatting: RunFormatting,
  options: StyleRenderOptions = {},
): string {
  const style = new StyleBuilder();

  if (formatting.fontFamily) style.set('font-family', quoteFontStack(formatting.fontFamilyRaw ?? formatting.fontFamily));
  if (formatting.fontSize) style.set('font-size', renderLength(formatting.fontSize, options));
  if (formatting.bold) style.set('font-weight', 'bold');
  if (formatting.italic) style.set('font-style', 'italic');

  const decorations: string[] = [];
  if (formatting.underline && formatting.underline !== 'none') decorations.push('underline');
  if (formatting.strike) decorations.push('line-through');
  if (decorations.length > 0) {
    style.set('text-decoration', decorations.join(' '));
    if (formatting.underline && formatting.underline !== 'single' && formatting.underline !== 'none') {
      style.set('text-decoration-style', underlineToCssStyle(formatting.underline));
    }
    if (formatting.underlineColor) {
      style.set('text-decoration-color', renderColor(formatting.underlineColor));
    }
  }

  if (formatting.color) style.set('color', renderColor(formatting.color));
  if (formatting.highlight) style.set('background-color', renderColor(formatting.highlight));

  if (formatting.verticalAlign === 'super' || formatting.verticalAlign === 'sub') {
    style.set('vertical-align', formatting.verticalAlign);
    // Browsers only shrink text inside <sup>/<sub>; a styled span needs it said.
    if (!formatting.fontSize) style.set('font-size', 'smaller');
  }

  if (formatting.letterSpacing) style.set('letter-spacing', renderLength(formatting.letterSpacing, options));
  if (formatting.characterScale !== undefined && formatting.characterScale !== 100) {
    style.set('transform', `scaleX(${roundTo(formatting.characterScale / 100, 3)})`);
    style.set('display', 'inline-block');
  }
  if (formatting.smallCaps) style.set('font-variant', 'small-caps');
  if (formatting.allCaps) style.set('text-transform', 'uppercase');
  if (formatting.direction === 'rtl') style.set('direction', 'rtl');
  if (formatting.hidden) style.set('display', 'none');

  return style.toString();
}

function underlineToCssStyle(underline: NonNullable<RunFormatting['underline']>): string {
  switch (underline) {
    case 'double':
      return 'double';
    case 'dotted':
      return 'dotted';
    case 'dashed':
      return 'dashed';
    case 'wave':
      return 'wavy';
    case 'thick':
      return 'solid';
    default:
      return 'solid';
  }
}

/** Paragraph formatting as inline CSS. */
export function renderParagraphStyle(
  formatting: ParagraphFormatting,
  options: StyleRenderOptions = {},
): string {
  const style = new StyleBuilder();

  if (formatting.alignment) style.set('text-align', formatting.alignment);
  if (formatting.marginLeft) style.set('margin-left', renderLength(formatting.marginLeft, options));
  if (formatting.marginRight) style.set('margin-right', renderLength(formatting.marginRight, options));

  // Word's spaceBefore/spaceAfter are rendered as padding, never as margin.
  // Adjacent block-level margins collapse in CSS — the browser keeps only the
  // larger of two touching margins, not their sum — so a 12pt space-after
  // meeting a 12pt space-before renders as a 12pt gap instead of 24pt. Word's
  // box model has no such collapsing: the gap between two paragraphs really is
  // the sum of both. Padding never collapses, so it is the only property that
  // reproduces the numbers Word actually declared. See ARCHITECTURE.md.
  if (formatting.spaceBefore) style.set('padding-top', renderLength(formatting.spaceBefore, options));
  if (formatting.spaceAfter) style.set('padding-bottom', renderLength(formatting.spaceAfter, options));

  if (formatting.textIndent) style.set('text-indent', renderLength(formatting.textIndent, options));

  const lineSpacing = formatting.lineSpacing;
  if (lineSpacing) {
    if (lineSpacing.rule === 'multiple') {
      style.set('line-height', String(roundTo(lineSpacing.value, 3)));
    } else if (lineSpacing.length) {
      style.set('line-height', renderLength(lineSpacing.length, options));
    }
  }

  if (formatting.keepWithNext) style.set('break-after', 'avoid');
  if (formatting.keepLines) style.set('break-inside', 'avoid');
  if (formatting.pageBreakBefore) style.set('break-before', 'page');
  if (formatting.widowControl) {
    style.set('widows', '2');
    style.set('orphans', '2');
  }
  if (formatting.direction === 'rtl') style.set('direction', 'rtl');

  const borders = renderBorders(formatting.borders, options);
  for (const [property, value] of borders) style.set(property, value);

  const background = formatting.shading?.fill ?? formatting.backgroundColor;
  if (background) style.set('background-color', renderColor(background));

  if (formatting.borders || formatting.shading) {
    // Word puts padding inside a bordered/shaded paragraph; without it the
    // border sits flush against the glyphs. Horizontal padding is only ever
    // needed for the box itself, so it is set unconditionally; vertical
    // padding defers to spaceBefore/spaceAfter above when either is present,
    // and only supplies a small fallback when neither is.
    if (!style.has('padding-top')) style.set('padding-top', '1pt');
    if (!style.has('padding-bottom')) style.set('padding-bottom', '1pt');
    style.set('padding-left', '4pt');
    style.set('padding-right', '4pt');
  }

  return style.toString();
}

/** Border declarations as `property -> value` pairs. */
export function renderBorders(
  borders: Borders | undefined,
  options: StyleRenderOptions = {},
): Array<[string, string]> {
  if (!borders) return [];
  const sides = ['top', 'right', 'bottom', 'left'] as const;
  const rendered = sides
    .map((side) => [side, borders[side]] as const)
    .filter((entry): entry is readonly [(typeof sides)[number], Border] => entry[1] !== undefined)
    .map(([side, border]) => [side, renderBorderValue(border, options)] as const)
    .filter(([, value]) => value !== '');

  if (rendered.length === 4) {
    const [first] = rendered;
    if (first && rendered.every(([, value]) => value === first[1])) {
      return [['border', first[1]]];
    }
  }
  return rendered.map(([side, value]) => [`border-${side}`, value]);
}

export function renderBorderValue(border: Border, options: StyleRenderOptions = {}): string {
  if (border.style === 'none') return 'none';
  const parts: string[] = [];
  parts.push(border.width ? renderLength(border.width, options) : '1pt');
  parts.push(border.style);
  if (border.color) parts.push(renderColor(border.color));
  return parts.join(' ');
}

export function renderShading(shading: Shading | undefined): string | undefined {
  if (!shading?.fill) return undefined;
  return renderColor(shading.fill);
}

/**
 * Quote a font stack so a family with spaces survives.
 *
 * Word writes `"Times New Roman",serif` already quoted, but also writes bare
 * `Courier New` in `@list` rules, which is invalid CSS and drops the font.
 */
export function quoteFontStack(stack: string): string {
  return stack
    .split(',')
    .map((family) => {
      const trimmed = family.trim().replace(/^["']|["']$/g, '');
      if (!trimmed) return '';
      if (/^(serif|sans-serif|monospace|cursive|fantasy|system-ui|ui-\w+|inherit|initial)$/i.test(trimmed)) {
        return trimmed.toLowerCase();
      }
      return /[\s'"]/.test(trimmed) ? `"${trimmed.replace(/"/g, '')}"` : trimmed;
    })
    .filter(Boolean)
    .join(', ');
}

/** Combine several style strings, dropping empties. */
export function joinStyles(...styles: Array<string | undefined>): string {
  return styles.filter((s): s is string => Boolean(s && s.length > 0)).join(';');
}
