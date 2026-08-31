import type {
  WordCssRule,
  WordFontDefinition,
  WordStyleDefinition,
  WordStyleSheet,
} from '../model/Style.js';
import type { CssRuleNode } from './WordCssTokenizer.js';
import { declarationsToMap, tokenizeCss, unquote } from './WordCssTokenizer.js';
import {
  parseParagraphFormattingFromCss,
  parseRunFormattingFromCss,
} from './WordFormattingParser.js';
import { parseListRules } from './WordListStyleParser.js';
import { isSymbolFont } from './WordSymbolFonts.js';
import type { DiagnosticCollector } from '../diagnostics/DiagnosticCollector.js';

/**
 * Builds the {@link WordStyleSheet} from the clipboard's `<style>` blocks.
 *
 * Word's stylesheet carries facts that exist nowhere else in the payload:
 *
 *   - `@list` numbering definitions (see WordListStyleParser).
 *   - `@font-face` blocks that say whether a font is a symbol font, which is
 *     what makes a Wingdings bullet decodable.
 *   - `mso-style-name` on each style rule, which is how a class called
 *     `MsoListParagraphCxSpFirst` is known to be "List Paragraph".
 *   - Paragraph and character defaults that inline styles then override.
 *
 * All of it is gone the moment the browser's CSS parser touches it, so the
 * stylesheet is read from raw text before the DOM is consulted.
 */

export interface ParseStyleSheetOptions {
  /** Cap on total CSS characters parsed, to bound a hostile payload. */
  maxCssLength?: number;
}

const DEFAULT_MAX_CSS_LENGTH = 4 * 1024 * 1024;

export function createEmptyStyleSheet(): WordStyleSheet {
  return { styles: {}, fonts: {}, lists: {}, pages: {}, rules: [], rawCss: '' };
}

/**
 * Parse one or more CSS texts (the contents of every `<style>` element found,
 * including those inside Word conditional comments) into a WordStyleSheet.
 */
export function parseWordStyleSheet(
  cssTexts: string[],
  diagnostics: DiagnosticCollector,
  options: ParseStyleSheetOptions = {},
): WordStyleSheet {
  const sheet = createEmptyStyleSheet();
  const maxLength = options.maxCssLength ?? DEFAULT_MAX_CSS_LENGTH;

  let combined = cssTexts.join('\n');
  if (combined.length > maxLength) {
    diagnostics.warn(
      'LIMIT_DOCUMENT_TRUNCATED',
      `Clipboard stylesheet was ${combined.length} characters; truncated to ${maxLength} to bound parsing.`,
      { details: { length: combined.length, limit: maxLength } },
    );
    combined = combined.slice(0, maxLength);
  }
  sheet.rawCss = combined;
  if (!combined.trim()) return sheet;

  const rules = tokenizeCss(combined);
  const flattened = flattenRules(rules);

  for (const rule of flattened) {
    sheet.rules.push(toWordCssRule(rule));
  }

  // Lists first: a style rule may reference a list, and list levels carry
  // formatting the style rules do not.
  const lists = parseListRules(flattened, diagnostics);
  sheet.lists = lists.definitions;

  for (const rule of flattened) {
    if (rule.kind === 'at-rule') {
      if (rule.atName === 'font-face') {
        const font = parseFontFace(rule);
        if (font) sheet.fonts[font.family.toLowerCase()] = font;
      } else if (rule.atName === 'page') {
        const name = (rule.prelude ?? 'default').trim() || 'default';
        sheet.pages[name] = declarationsToMap(rule.declarations);
      }
      continue;
    }
    mergeStyleRule(sheet, rule);
  }

  return sheet;
}

/** Flatten nested conditional groups (`@media print { … }`) into one list. */
function flattenRules(rules: CssRuleNode[]): CssRuleNode[] {
  const out: CssRuleNode[] = [];
  for (const rule of rules) {
    if (rule.children.length > 0) {
      out.push(...flattenRules(rule.children));
    } else {
      out.push(rule);
    }
  }
  return out;
}

function toWordCssRule(rule: CssRuleNode): WordCssRule {
  const kind: WordCssRule['kind'] =
    rule.kind === 'rule'
      ? 'style'
      : rule.atName === 'list'
        ? 'at-list'
        : rule.atName === 'font-face'
          ? 'at-font-face'
          : rule.atName === 'page'
            ? 'at-page'
            : 'at-other';
  return {
    kind,
    selector: rule.selector,
    declarations: declarationsToMap(rule.declarations),
    raw: rule.raw,
  };
}

function parseFontFace(rule: CssRuleNode): WordFontDefinition | null {
  const declarations = declarationsToMap(rule.declarations);
  const familyRaw = declarations['font-family'];
  if (!familyRaw) return null;
  const family = unquote(familyRaw.split(',')[0] ?? familyRaw).trim();
  if (!family) return null;

  const definition: WordFontDefinition = {
    family,
    declarations,
    // Word marks symbol fonts with charset 2. The name check is the backstop
    // for payloads where the @font-face block did not survive.
    isSymbolFont:
      declarations['mso-font-charset'] === '2' ||
      /decorative/i.test(declarations['mso-generic-font-family'] ?? '') ||
      isSymbolFont(family),
  };
  const panose = declarations['panose-1'];
  if (panose) definition.panose1 = panose;
  const charset = declarations['mso-font-charset'];
  if (charset) definition.charset = charset;
  const alt = declarations['mso-font-alt'];
  if (alt) definition.alt = unquote(alt);
  const generic = declarations['mso-generic-font-family'];
  if (generic) definition.genericFamily = generic;
  return definition;
}

/**
 * Word declares one style across several selectors:
 *
 *     p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
 *       {mso-style-name:"List Paragraph"; margin-left:.5in;}
 *
 * Each selector is registered under its own normalised id so a lookup by class
 * name works, and they all share one definition object.
 */
function mergeStyleRule(sheet: WordStyleSheet, rule: CssRuleNode): void {
  const declarations = declarationsToMap(rule.declarations);
  if (Object.keys(declarations).length === 0) return;

  const selectors = rule.selector.split(',').map((s) => s.trim()).filter(Boolean);
  if (selectors.length === 0) return;

  const styleName = declarations['mso-style-name']
    ? unquote(declarations['mso-style-name'])
    : undefined;

  for (const selector of selectors) {
    const target = styleIdFromSelector(selector);
    if (!target) continue;

    const existing = sheet.styles[target.id];
    const definition: WordStyleDefinition = existing ?? {
      id: target.id,
      name: styleName ?? humaniseClassName(target.name),
      selectors: [],
      type: target.type,
      declarations: {},
      run: {},
      paragraph: {},
    };

    if (styleName) definition.name = styleName;
    if (!definition.selectors.includes(selector)) definition.selectors.push(selector);
    // A rule matching several element types (p/li/div) is a paragraph style;
    // `span.X` alone is a character style.
    if (definition.type === 'unknown') definition.type = target.type;
    else if (definition.type !== target.type && target.type !== 'unknown') {
      definition.type = definition.type === 'character' ? target.type : definition.type;
    }

    definition.declarations = { ...definition.declarations, ...declarations };
    definition.run = parseRunFormattingFromCss(definition.declarations);
    definition.paragraph = parseParagraphFormattingFromCss(definition.declarations);

    const parent = declarations['mso-style-parent'];
    if (parent) definition.parent = normaliseStyleId(unquote(parent));
    const link = declarations['mso-style-link'];
    if (link) definition.link = normaliseStyleId(unquote(link));
    const outline = definition.paragraph.outlineLevel;
    if (outline !== undefined) definition.outlineLevel = outline;

    sheet.styles[target.id] = definition;
  }
}

interface SelectorTarget {
  id: string;
  name: string;
  type: WordStyleDefinition['type'];
}

/**
 * Reduce a selector to the style it names.
 *
 * `p.MsoHeading1` -> paragraph style `msoheading1`
 * `span.EmphasisChar` -> character style `emphasischar`
 * `table.MsoTableGrid` -> table style `msotablegrid`
 * `p.MsoNormal, li.MsoNormal, div.MsoNormal` -> handled per selector
 * A bare element selector (`h1`, `p`) registers under `element:h1`.
 */
export function styleIdFromSelector(selector: string): SelectorTarget | null {
  const trimmed = selector.trim();
  if (!trimmed || trimmed.startsWith('@')) return null;

  // Ignore pseudo-elements and attribute noise; keep the last simple selector.
  const last = trimmed.split(/\s+|>/).filter(Boolean).pop() ?? trimmed;
  const clean = last.replace(/::?[a-z-]+(\([^)]*\))?/gi, '');

  const classMatch = /^([a-zA-Z][\w-]*)?\.([\w-]+)/.exec(clean);
  if (classMatch) {
    const element = (classMatch[1] ?? '').toLowerCase();
    const className = classMatch[2]!;
    return {
      id: normaliseStyleId(className),
      name: className,
      type:
        element === 'span'
          ? 'character'
          : element === 'table' || element === 'td' || element === 'tr' || element === 'th'
            ? 'table'
            : element === ''
              ? 'unknown'
              : 'paragraph',
    };
  }

  const elementMatch = /^([a-zA-Z][\w-]*)$/.exec(clean);
  if (elementMatch) {
    const element = elementMatch[1]!.toLowerCase();
    return { id: `element:${element}`, name: element, type: elementTypeOf(element) };
  }

  return null;
}

function elementTypeOf(element: string): WordStyleDefinition['type'] {
  if (element === 'span' || element === 'a' || element === 'b' || element === 'i') {
    return 'character';
  }
  if (element === 'table' || element === 'td' || element === 'th' || element === 'tr') {
    return 'table';
  }
  return 'paragraph';
}

/**
 * `MsoListParagraphCxSpFirst` -> `List Paragraph`.
 *
 * Word omits `mso-style-name` when the style's formatting matched the
 * default, so the class name is sometimes the only name available. The
 * `CxSpFirst`/`CxSpMiddle`/`CxSpLast` suffixes describe where a paragraph sits
 * in a run of same-styled paragraphs, not a different style, so they go.
 */
export function humaniseClassName(className: string): string {
  let name = className;
  name = name.replace(/CxSp(First|Middle|Last)$/i, '');
  name = name.replace(/^Mso/, '');
  name = name.replace(/([a-z0-9])([A-Z])/g, '$1 $2');
  return name.trim() || className;
}

/** Normalise a style name or class into a lookup key. */
export function normaliseStyleId(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, '');
}

/**
 * Resolve a style plus its `mso-style-parent` chain into a flat declaration map.
 *
 * Word only writes the differences from the parent style, so a Heading 2 rule
 * may say nothing about the font and inherit it from Heading 1's linked style.
 */
export function resolveStyleChain(
  sheet: WordStyleSheet,
  styleId: string,
  seen = new Set<string>(),
): Record<string, string> {
  const definition = sheet.styles[styleId];
  if (!definition || seen.has(styleId)) return {};
  seen.add(styleId);
  const parentDeclarations = definition.parent
    ? resolveStyleChain(sheet, definition.parent, seen)
    : {};
  return { ...parentDeclarations, ...definition.declarations };
}

/**
 * Look a style up by any of the names Word might use for it: the class name
 * (`MsoHeading1`), the display name (`Heading 1`), or the normalised id.
 */
export function findStyle(
  sheet: WordStyleSheet,
  name: string | undefined,
): WordStyleDefinition | undefined {
  if (!name) return undefined;
  const id = normaliseStyleId(name);
  const direct = sheet.styles[id];
  if (direct) return direct;
  for (const style of Object.values(sheet.styles)) {
    if (normaliseStyleId(style.name) === id) return style;
  }
  return undefined;
}

/** Convenience: parse a single CSS string. */
export function parseWordCss(
  css: string,
  diagnostics: DiagnosticCollector,
  options?: ParseStyleSheetOptions,
): WordStyleSheet {
  return parseWordStyleSheet([css], diagnostics, options);
}
