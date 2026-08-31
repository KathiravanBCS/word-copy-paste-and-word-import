/**
 * A small, dependency-free CSS tokenizer sized for Word clipboard stylesheets.
 *
 * Why not `CSSStyleSheet` / the browser's own parser? Because it throws away
 * exactly the information this engine needs:
 *
 *   - Unknown at-rules. `@list l0:level1 { … }` is not valid CSS, so browsers
 *     drop the whole rule. That rule *is* Word's numbering definition.
 *   - Unknown properties. Every `mso-*` declaration is discarded by a
 *     conforming parser, taking `mso-level-text`, `mso-list`, `mso-style-name`
 *     and the rest with it.
 *
 * So the stylesheet is parsed here, from the raw text, before the DOM ever
 * sees it. The tokenizer is tolerant: Word emits stylesheets wrapped in HTML
 * comments, with unbalanced constructs after a truncated copy, and it must
 * never throw on real input.
 */

export interface CssDeclaration {
  /** Lower-cased property name. */
  property: string;
  /** Value with surrounding whitespace trimmed; quoting and escapes intact. */
  value: string;
  /** True when the declaration carried `!important`. */
  important: boolean;
}

export interface CssRuleNode {
  kind: 'rule' | 'at-rule';
  /** For a style rule: the full selector text. For an at-rule: `@name prelude`. */
  selector: string;
  /** At-rule name without the `@`, lower-cased. */
  atName?: string;
  /** At-rule prelude, e.g. `l0:level1` for `@list l0:level1`. */
  prelude?: string;
  declarations: CssDeclaration[];
  /** Nested rules, for conditional groups such as `@media`. */
  children: CssRuleNode[];
  /** The rule's source text, for diagnostics. */
  raw: string;
}

/** At-rules whose bodies contain nested rules rather than declarations. */
const NESTING_AT_RULES = new Set(['media', 'supports', 'document', '-moz-document', 'layer']);
/** At-rules that have no block at all. */
const STATEMENT_AT_RULES = new Set(['import', 'charset', 'namespace']);

export interface TokenizeOptions {
  /** Hard cap on rules parsed, to bound a hostile or enormous payload. */
  maxRules?: number;
}

const DEFAULT_MAX_RULES = 20000;

/**
 * Parse a stylesheet into a flat-ish rule tree. Never throws.
 */
export function tokenizeCss(input: string, options: TokenizeOptions = {}): CssRuleNode[] {
  const css = stripStyleCommentWrapper(input);
  const state = { pos: 0, rules: 0, max: options.maxRules ?? DEFAULT_MAX_RULES };
  return parseRuleList(css, state, css.length);
}

/**
 * Word wraps stylesheet contents in an HTML comment so that ancient browsers
 * do not render the CSS as text:
 *
 *     <style><!--
 *      p.MsoNormal { … }
 *     --></style>
 *
 * The DOM hands us that wrapper as part of `textContent`, so it is stripped
 * here rather than in every caller.
 */
export function stripStyleCommentWrapper(css: string): string {
  let out = css.trim();
  if (out.startsWith('<!--')) out = out.slice(4);
  if (out.endsWith('-->')) out = out.slice(0, -3);
  // Word also emits stray `<!--` / `-->` between blocks in some payloads.
  return out.replace(/^\s*<!--/gm, '').replace(/-->\s*$/gm, '');
}

interface ParseState {
  pos: number;
  rules: number;
  max: number;
}

function parseRuleList(css: string, state: ParseState, end: number): CssRuleNode[] {
  const rules: CssRuleNode[] = [];
  while (state.pos < end) {
    skipWhitespaceAndComments(css, state, end);
    if (state.pos >= end) break;
    if (css[state.pos] === '}') {
      state.pos++;
      continue;
    }
    if (state.rules >= state.max) break;
    const rule = parseRule(css, state, end);
    if (rule) {
      rules.push(rule);
      state.rules++;
    } else {
      break;
    }
  }
  return rules;
}

function parseRule(css: string, state: ParseState, end: number): CssRuleNode | null {
  const start = state.pos;
  const prelude = readPrelude(css, state, end);
  if (prelude === null) return null;

  const selector = collapseWhitespace(prelude.text);
  const isAtRule = selector.startsWith('@');
  let atName: string | undefined;
  let atPrelude: string | undefined;
  if (isAtRule) {
    const match = /^@([a-zA-Z-]+)\s*([\s\S]*)$/.exec(selector);
    atName = (match?.[1] ?? '').toLowerCase();
    atPrelude = (match?.[2] ?? '').trim();
  }

  if (prelude.terminator === ';' || prelude.terminator === 'eof') {
    // A statement at-rule (`@import url(x);`) or trailing junk after a
    // truncated copy. Either way there is no block to read.
    if (!isAtRule) return null;
    const node: CssRuleNode = {
      kind: 'at-rule',
      selector,
      declarations: [],
      children: [],
      raw: css.slice(start, state.pos),
    };
    if (atName) node.atName = atName;
    if (atPrelude) node.prelude = atPrelude;
    if (atName && !STATEMENT_AT_RULES.has(atName)) return node;
    return node;
  }

  // We are just past the opening `{`.
  const blockStart = state.pos;
  const blockEnd = findBlockEnd(css, blockStart, end);
  const body = css.slice(blockStart, blockEnd);

  const node: CssRuleNode = {
    kind: isAtRule ? 'at-rule' : 'rule',
    selector,
    declarations: [],
    children: [],
    raw: css.slice(start, Math.min(blockEnd + 1, end)),
  };
  if (atName) node.atName = atName;
  if (atPrelude !== undefined) node.prelude = atPrelude;

  if (atName && NESTING_AT_RULES.has(atName)) {
    const nested: ParseState = { pos: 0, rules: state.rules, max: state.max };
    node.children = parseRuleList(body, nested, body.length);
    state.rules = nested.rules;
  } else {
    node.declarations = parseDeclarations(body);
  }

  state.pos = Math.min(blockEnd + 1, end);
  return node;
}

interface Prelude {
  text: string;
  terminator: '{' | ';' | 'eof';
}

function readPrelude(css: string, state: ParseState, end: number): Prelude | null {
  let text = '';
  let quote: string | null = null;
  let parens = 0;
  while (state.pos < end) {
    const ch = css[state.pos]!;
    if (quote) {
      text += ch;
      if (ch === '\\' && state.pos + 1 < end) {
        text += css[state.pos + 1];
        state.pos += 2;
        continue;
      }
      if (ch === quote) quote = null;
      state.pos++;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      text += ch;
      state.pos++;
      continue;
    }
    if (ch === '/' && css[state.pos + 1] === '*') {
      skipComment(css, state, end);
      continue;
    }
    if (ch === '(') parens++;
    if (ch === ')' && parens > 0) parens--;
    if (parens === 0) {
      if (ch === '{') {
        state.pos++;
        return { text: text.trim(), terminator: '{' };
      }
      if (ch === ';') {
        state.pos++;
        return { text: text.trim(), terminator: ';' };
      }
      if (ch === '}') {
        // Stray close brace: swallow it and treat what we have as junk.
        state.pos++;
        return text.trim() ? { text: text.trim(), terminator: ';' } : null;
      }
    }
    text += ch;
    state.pos++;
  }
  return text.trim() ? { text: text.trim(), terminator: 'eof' } : null;
}

/** Find the index of the `}` closing the block that starts at `from`. */
function findBlockEnd(css: string, from: number, end: number): number {
  let depth = 1;
  let quote: string | null = null;
  let i = from;
  while (i < end) {
    const ch = css[i]!;
    if (quote) {
      if (ch === '\\') {
        i += 2;
        continue;
      }
      if (ch === quote) quote = null;
      i++;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      i++;
      continue;
    }
    if (ch === '/' && css[i + 1] === '*') {
      const close = css.indexOf('*/', i + 2);
      i = close === -1 ? end : close + 2;
      continue;
    }
    if (ch === '{') depth++;
    else if (ch === '}') {
      depth--;
      if (depth === 0) return i;
    }
    i++;
  }
  return end;
}

/** Split a declaration block into property/value pairs. Tolerates junk. */
export function parseDeclarations(body: string): CssDeclaration[] {
  const declarations: CssDeclaration[] = [];
  let buffer = '';
  let quote: string | null = null;
  let parens = 0;

  const flush = (): void => {
    const text = buffer.trim();
    buffer = '';
    if (!text) return;
    const declaration = parseDeclaration(text);
    if (declaration) declarations.push(declaration);
  };

  for (let i = 0; i < body.length; i++) {
    const ch = body[i]!;
    if (quote) {
      buffer += ch;
      if (ch === '\\' && i + 1 < body.length) {
        buffer += body[i + 1];
        i++;
        continue;
      }
      if (ch === quote) quote = null;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      buffer += ch;
      continue;
    }
    if (ch === '/' && body[i + 1] === '*') {
      const close = body.indexOf('*/', i + 2);
      i = close === -1 ? body.length : close + 1;
      continue;
    }
    if (ch === '(') parens++;
    else if (ch === ')' && parens > 0) parens--;
    if (ch === ';' && parens === 0) {
      flush();
      continue;
    }
    buffer += ch;
  }
  flush();
  return declarations;
}

function parseDeclaration(text: string): CssDeclaration | null {
  // Find the first `:` that is not inside a string or parentheses. Word writes
  // values such as `mso-level-text:"%1\."` and `font:7.0pt "Times New Roman"`.
  let quote: string | null = null;
  let parens = 0;
  let colon = -1;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    if (quote) {
      if (ch === '\\') {
        i++;
        continue;
      }
      if (ch === quote) quote = null;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === '(') parens++;
    else if (ch === ')' && parens > 0) parens--;
    else if (ch === ':' && parens === 0) {
      colon = i;
      break;
    }
  }
  if (colon <= 0) return null;
  const property = text.slice(0, colon).trim().toLowerCase();
  let value = text.slice(colon + 1).trim();
  if (!property) return null;

  let important = false;
  const importantMatch = /!\s*important\s*$/i.exec(value);
  if (importantMatch) {
    important = true;
    value = value.slice(0, importantMatch.index).trim();
  }
  return { property, value, important };
}

/** Parse an inline `style="…"` attribute into a property map. */
export function parseInlineStyle(style: string | null | undefined): Record<string, string> {
  if (!style) return {};
  const map: Record<string, string> = {};
  for (const declaration of parseDeclarations(style)) {
    map[declaration.property] = declaration.value;
  }
  return map;
}

/** Turn a declaration list into a lower-cased property map (last wins). */
export function declarationsToMap(declarations: CssDeclaration[]): Record<string, string> {
  const map: Record<string, string> = {};
  for (const declaration of declarations) {
    map[declaration.property] = declaration.value;
  }
  return map;
}

function skipWhitespaceAndComments(css: string, state: ParseState, end: number): void {
  for (;;) {
    while (state.pos < end && /\s/.test(css[state.pos]!)) state.pos++;
    if (css[state.pos] === '/' && css[state.pos + 1] === '*') {
      skipComment(css, state, end);
      continue;
    }
    return;
  }
}

function skipComment(css: string, state: ParseState, end: number): void {
  const close = css.indexOf('*/', state.pos + 2);
  state.pos = close === -1 ? end : close + 2;
}

function collapseWhitespace(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

/** Strip surrounding quotes from a CSS string value. */
export function unquote(value: string): string {
  const trimmed = value.trim();
  if (trimmed.length >= 2) {
    const first = trimmed[0];
    const last = trimmed[trimmed.length - 1];
    if ((first === '"' && last === '"') || (first === "'" && last === "'")) {
      return trimmed.slice(1, -1);
    }
  }
  return trimmed;
}
