import type { ListItemInfo } from '../model/List.js';
import type { WordNumberFormat } from '../model/ListLevel.js';
import type { ListTreeNode } from '../normalization/NormalizeLists.js';
import { computeListIndentation, type ListIndentation } from '../normalization/NormalizeUnits.js';
import { escapeHtmlAttribute, escapeHtmlText } from '../util/dom.js';
import { quoteFontStack } from './HtmlStyleRenderer.js';
import { roundTo } from '../word/WordLengthParser.js';

/**
 * List rendering.
 *
 * The requirement this file exists to satisfy: **a list marker is never text.**
 * `1.` and `•` are markers; `Item` is content; and the two are separate in the
 * output, so the list can be re-indented, re-numbered, extended or restyled by
 * whatever receives it.
 *
 * Two rendering modes, both of which keep that separation:
 *
 *   `native`
 *       Real `<ul>`/`<ol>` with real browser markers. Word's numbering
 *       definition is compiled into a generated `@counter-style` rule, so a
 *       Word upper-roman level with a `.` suffix becomes an actual
 *       `list-style-type` and the browser draws it. Numbers carry `value` so
 *       the count is exact even across continuations, and because it is a
 *       real `<ol>`, a live editor's own list-continuation logic can extend
 *       it — press Enter at the end of an item and the next number is the
 *       browser's, not this engine's.
 *
 *       Level texts CSS counter styles cannot express — `%1.%2` and friends —
 *       fall back to `element` rendering for that level (below), not to
 *       `::marker { content: attr(data-marker) }`. That was tried and
 *       measured against a real editor's content model (RoosterJS): the
 *       `data-marker` attribute it depends on does not survive the round trip
 *       — dropped the same way an external stylesheet's classes are — so the
 *       marker silently vanished. There was never a way to auto-continue a
 *       composite counter like "1.1" with a single CSS counter-style anyway,
 *       so falling back costs nothing `native` mode could actually deliver
 *       for that case.
 *
 *   `element`  (default)
 *       `list-style: none` plus an explicit `<span class="wce-marker">`. Less
 *       native, but reproduces Word's hanging-indent geometry exactly, and
 *       depends on nothing but the inline `style` this engine already puts on
 *       the element — which is what survives environments (rich-text editors
 *       especially) that strip `::marker` styling, CSS classes, and
 *       non-standard attributes alike.
 *
 * What neither mode ever does is re-interpret the numbering. If the model says
 * upper-roman with level text `%1.`, the output says `I.`; if it says `%1.%2`,
 * the output says `1.1`. The renderer has no opinion of its own.
 */

export type ListMarkerMode = 'native' | 'element';

export interface ListRenderOptions {
  markerMode?: ListMarkerMode;
  /** CSS class prefix for generated classes. Default `wce`. */
  classPrefix?: string;
  /** Emit `data-word-*` attributes describing the Word list definition. */
  includeWordMetadata?: boolean;
}

/** Collects the CSS a render pass generates. */
export class CssRegistry {
  private readonly rules: string[] = [];
  private readonly names = new Map<string, string>();
  private sequence = 0;

  constructor(private readonly prefix: string) {}

  /** Return an existing generated name for `signature`, or mint a new one. */
  intern(signature: string, build: (name: string) => string[]): string {
    const existing = this.names.get(signature);
    if (existing) return existing;
    const name = `${this.prefix}-${++this.sequence}`;
    this.names.set(signature, name);
    this.rules.push(...build(name));
    return name;
  }

  add(rule: string): void {
    if (!this.rules.includes(rule)) this.rules.push(rule);
  }

  toCss(): string {
    return this.rules.join('\n');
  }

  get size(): number {
    return this.rules.length;
  }
}

/** Base counter styles a Word number format maps onto directly. */
const CSS_COUNTER_SYSTEM: Partial<Record<WordNumberFormat, string>> = {
  decimal: 'decimal',
  'decimal-leading-zero': 'decimal-leading-zero',
  'lower-alpha': 'lower-alpha',
  'upper-alpha': 'upper-alpha',
  'lower-roman': 'lower-roman',
  'upper-roman': 'upper-roman',
  'cardinal-text': 'decimal',
  'ordinal-text': 'decimal',
  ordinal: 'decimal',
  hebrew: 'hebrew',
} as Partial<Record<WordNumberFormat, string>>;

/** How a level's marker will be produced. */
type MarkerStrategy = { kind: 'counter-style'; name: string } | { kind: 'element' };

export interface ListLevelStyle {
  /** CSS class applied to the `<ol>`/`<ul>`. */
  listClass: string;
  strategy: MarkerStrategy;
  indentation: ListIndentation;
}

/**
 * Compile one list level into CSS and decide how its marker is produced.
 */
export function compileLevelStyle(
  node: ListTreeNode,
  item: ListItemInfo,
  indentation: ListIndentation,
  parentIndentation: ListIndentation | null,
  css: CssRegistry,
  options: ListRenderOptions,
): ListLevelStyle {
  const prefix = options.classPrefix ?? 'wce';
  const mode = options.markerMode ?? 'element';
  const definition = item.levelDefinition;
  const marker = item.marker;

  const relativeIndentPx = parentIndentation
    ? roundTo(Math.max(0, indentation.textOffsetPx - parentIndentation.textOffsetPx), 2)
    : roundTo(indentation.textOffsetPx, 2);
  const hangingPx = roundTo(Math.max(0, indentation.hangingPx), 2);

  const markerFont = marker.font ?? definition?.bulletFont;
  const needsFont = Boolean(markerFont) && marker.fontMapped !== true;

  let strategy: MarkerStrategy;
  if (mode === 'element') {
    strategy = { kind: 'element' };
  } else if (marker.type === 'bullet') {
    const glyph = marker.glyph ?? definition?.bulletGlyph ?? '•';
    const signature = `bullet|${glyph}|${markerFont ?? ''}`;
    const name = css.intern(signature, (generated) => [
      `@counter-style ${generated} {`,
      `  system: cyclic;`,
      `  symbols: "${cssStringEscape(glyph)}";`,
      `  suffix: " ";`,
      `}`,
    ]);
    strategy = { kind: 'counter-style', name };
  } else {
    const simple = asSimpleLevelText(
      definition?.levelText ?? marker.levelText,
      node.level,
      definition?.numberFormat ?? marker.numberFormat ?? 'decimal',
    );
    if (simple) {
      const signature = `number|${simple.system}|${simple.prefix}|${simple.suffix}`;
      const name = css.intern(signature, (generated) => [
        `@counter-style ${generated} {`,
        `  system: extends ${simple.system};`,
        ...(simple.prefix ? [`  prefix: "${cssStringEscape(simple.prefix)}";`] : []),
        `  suffix: "${cssStringEscape(simple.suffix)} ";`,
        `}`,
      ]);
      strategy = { kind: 'counter-style', name };
    } else {
      // A multi-level level text such as `%1.%2`, or a Word number format
      // with no CSS counter-style equivalent. `::marker { content:
      // attr(data-marker) }` can draw this without ever putting the marker
      // in the text flow — but only when whatever receives the HTML keeps
      // that `data-marker` attribute on the `<li>`. Verified directly against
      // a real editor's content model (RoosterJS): it does not — attributes
      // it does not itself recognise are dropped when pasted content is
      // converted to the editor's own model, same as it drops CSS classes
      // (see the `element`-mode block above). A marker that depends on an
      // attribute surviving an arbitrary consumer is not durable enough to
      // call "the native option", so this falls back to the same `element`
      // rendering non-native mode uses: a real, protectable span that only
      // needs its inline `style` preserved, which is far more commonly kept
      // intact than an arbitrary `data-*` attribute. It costs the
      // auto-continuation `native` mode exists for, but only for the
      // composite levels that could never really have it anyway — a
      // "1.1"-shaped counter cannot be a single CSS counter-style, so there
      // was never a real browser-native mechanism to auto-continue it with.
      strategy = { kind: 'element' };
    }
  }

  const listClassSignature = [
    strategy.kind,
    strategy.kind === 'counter-style' ? strategy.name : '',
    relativeIndentPx,
    hangingPx,
    needsFont ? (markerFont ?? '') : '',
    marker.type,
  ].join('|');

  const listClass = css.intern(`list|${listClassSignature}`, (generated) => {
    const rules: string[] = [];
    const listSelector = `.${generated}`;
    const itemSelector = `.${generated} > li`;

    rules.push(
      `${listSelector} {`,
      `  margin: 0;`,
      `  padding-left: ${relativeIndentPx}px;`,
      ...(strategy.kind === 'counter-style' ? [`  list-style-type: ${strategy.name};`] : []),
      ...(strategy.kind !== 'counter-style' ? [`  list-style-type: none;`] : []),
      `}`,
    );

    if (strategy.kind === 'element') {
      // Word's geometry exactly: the first line starts at the marker offset,
      // the marker occupies the hanging width, so text lines up at the text
      // offset on every line, wrapped or not.
      rules.push(
        `${itemSelector} {`,
        `  padding-left: ${hangingPx}px;`,
        `  text-indent: -${hangingPx}px;`,
        `}`,
        `${itemSelector} > .${prefix}-marker {`,
        `  display: inline-block;`,
        `  min-width: ${hangingPx}px;`,
        `  text-indent: 0;`,
        ...(needsFont && markerFont ? [`  font-family: ${quoteFontStack(markerFont)};`] : []),
        `}`,
      );
    } else if (needsFont && markerFont) {
      rules.push(`${itemSelector}::marker { font-family: ${quoteFontStack(markerFont)}; }`);
    }

    return rules;
  });

  return { listClass, strategy, indentation };
}

interface SimpleLevelText {
  system: string;
  prefix: string;
  suffix: string;
}

/**
 * Decide whether a Word level text can become a CSS counter style.
 *
 * `%1.` at level 0 can: it is this level's own counter with a `.` after it.
 * `%1.%2` cannot: CSS counter styles format one number, not a hierarchy.
 * `%1.` at level 2 cannot either — it references an ancestor's counter.
 */
export function asSimpleLevelText(
  levelText: string | undefined,
  level: number,
  format: WordNumberFormat,
): SimpleLevelText | null {
  const system = CSS_COUNTER_SYSTEM[format];
  if (!system) return null;
  // Word formats with no CSS equivalent must not be silently rendered as
  // decimal; they take the literal-marker path instead.
  if (format === 'ordinal' || format === 'ordinal-text' || format === 'cardinal-text') return null;

  if (!levelText) return { system, prefix: '', suffix: '' };

  const match = /^([^%]*)%(\d)([^%]*)$/.exec(levelText);
  if (!match) return null;
  const placeholder = Number.parseInt(match[2]!, 10);
  if (placeholder !== level + 1) return null;

  return { system, prefix: match[1] ?? '', suffix: match[3] ?? '' };
}

/** Attributes for one `<li>`, given the marker strategy in effect. */
export function renderListItemAttributes(
  item: ListItemInfo,
  strategy: MarkerStrategy,
  options: ListRenderOptions,
): string {
  const parts: string[] = [];

  if (item.marker.type === 'number' && strategy.kind === 'counter-style') {
    const value = numericValue(item);
    if (value !== undefined) parts.push(` value="${value}"`);
  }
  if (options.includeWordMetadata) {
    parts.push(` data-word-list="${escapeHtmlAttribute(item.listId)}"`);
    parts.push(` data-word-level="${item.level + 1}"`);
    if (item.marker.levelText) {
      parts.push(` data-word-level-text="${escapeHtmlAttribute(readableGlyph(item.marker.levelText))}"`);
    }
    if (item.marker.numberFormat) {
      parts.push(` data-word-number-format="${escapeHtmlAttribute(item.marker.numberFormat)}"`);
    }
    if (item.marker.rawGlyph && item.marker.rawGlyph !== item.marker.glyph) {
      parts.push(` data-word-raw-glyph="${escapeHtmlAttribute(readableGlyph(item.marker.rawGlyph))}"`);
    }
    if (item.marker.font) {
      parts.push(` data-word-marker-font="${escapeHtmlAttribute(item.marker.font)}"`);
    }
  }
  return parts.join('');
}

/**
 * Render a glyph for a `data-` attribute.
 *
 * A symbol-font byte lifted into the Unicode private use area (U+F0B7 and
 * friends) has no printable form: putting it in an attribute verbatim writes an
 * invisible character into the output that shows as a blank or a tofu box.
 * `U+F0B7` says the same thing and can be read.
 */
function readableGlyph(value: string): string {
  let out = '';
  for (const char of value) {
    const code = char.codePointAt(0)!;
    out += code >= 0xe000 && code <= 0xf8ff
      ? `U+${code.toString(16).toUpperCase().padStart(4, '0')}`
      : char;
  }
  return out;
}

/**
 * The counter value for a numbered item, read back out of the marker Word drew.
 *
 * Using Word's own text as the source means a list copied from the middle of a
 * document numbers from where it actually started, not from 1.
 */
function numericValue(item: ListItemInfo): number | undefined {
  const text = item.marker.text;
  if (!text) return item.startAt;
  const cleaned = text.trim().replace(/^[([{]+/, '').replace(/[.)\]}]+$/, '');
  const last = cleaned.split('.').pop() ?? cleaned;
  const format = item.marker.numberFormat ?? item.levelDefinition?.numberFormat ?? 'decimal';

  if (format === 'lower-roman' || format === 'upper-roman') {
    const value = romanToInt(last.toUpperCase());
    return value > 0 ? value : item.startAt;
  }
  if (format === 'lower-alpha' || format === 'upper-alpha') {
    const value = alphaToInt(last.toUpperCase());
    return value > 0 ? value : item.startAt;
  }
  const value = Number.parseInt(last, 10);
  return Number.isFinite(value) ? value : item.startAt;
}

const ROMAN: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };

function romanToInt(text: string): number {
  if (!/^[IVXLCDM]+$/.test(text)) return 0;
  let total = 0;
  for (let i = 0; i < text.length; i++) {
    const value = ROMAN[text[i]!] ?? 0;
    const next = ROMAN[text[i + 1] ?? ''] ?? 0;
    total += value < next ? -value : value;
  }
  return total;
}

function alphaToInt(text: string): number {
  if (!/^[A-Z]+$/.test(text)) return 0;
  let total = 0;
  for (const char of text) total = total * 26 + (char.charCodeAt(0) - 64);
  return total;
}

/** The explicit marker element used by `element` mode. */
export function renderMarkerElement(item: ListItemInfo, options: ListRenderOptions): string {
  const prefix = options.classPrefix ?? 'wce';
  const text =
    item.marker.type === 'bullet'
      ? (item.marker.glyph ?? item.marker.rawGlyph ?? '•')
      : (item.marker.text ?? '');
  if (!text) return '';
  // `aria-hidden` because the list element already conveys "this is a list
  // item" to assistive technology; the glyph would be read out twice.
  //
  // `contenteditable="false"` because this HTML is not only ever displayed
  // read-only — it is routinely inserted into a live editable surface (a
  // rich-text editor's paste handler). Without it, a cursor placed inside or
  // before the marker lets ordinary typing, backspace and delete edit the
  // glyph directly: "•" becomes "-" because someone typed a dash next to it,
  // or a stray keystroke lands inside the span and the marker is now
  // literally text again — the exact failure this whole engine exists to
  // prevent, just reintroduced one editing session after the paste. Marking
  // the span as an atomic, non-editable island is the standard technique
  // every mainstream rich-text editor uses for this (mentions, emoji, and
  // other generated "chips" inside editable text). It is inert HTML when the
  // page itself is not editable, so it costs nothing in the static/display
  // case this attribute doesn't apply to.
  return `<span class="${prefix}-marker" contenteditable="false" aria-hidden="true">${escapeHtmlText(text)}</span>`;
}

/** Compute a list item's indentation, honouring the parent list's. */
export function indentationFor(
  item: ListItemInfo,
  formatting: Parameters<typeof computeListIndentation>[1],
): ListIndentation {
  return computeListIndentation(item, formatting);
}

/** Escape a string for use inside a CSS string literal. */
export function cssStringEscape(value: string): string {
  let out = '';
  for (const char of value) {
    const code = char.codePointAt(0)!;
    if (char === '"' || char === '\\') out += `\\${char}`;
    else if (code < 0x20 || code === 0x7f) out += `\\${code.toString(16)} `;
    else out += char;
  }
  return out;
}

/** The element name for a list node. */
export function listTagFor(node: ListTreeNode): 'ol' | 'ul' {
  return node.ordered ? 'ol' : 'ul';
}
