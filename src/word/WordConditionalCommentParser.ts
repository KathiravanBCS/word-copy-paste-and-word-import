/**
 * Word conditional comments.
 *
 * Word emits two syntactically different things that both arrive as comment
 * nodes, and they must not be treated the same way:
 *
 *  1. **Downlevel-hidden**  `<!--[if gte mso 9]> … <![endif]-->`
 *     A real HTML comment. Its *content* is markup that no browser ever
 *     renders — office metadata (`<o:DocumentProperties>`), Word settings
 *     (`<w:WordDocument>`), and the VML twin of every image. Deleting it
 *     without looking is how image data and document metadata get lost.
 *
 *  2. **Downlevel-revealed**  `<![if !supportLists]> … <![endif]>`
 *     Not valid HTML at all. The parser produces a *bogus comment* for the
 *     opening and closing tokens, and the content between them is live DOM.
 *     This is where Word puts the rendered list marker — the single most
 *     important piece of markup in the whole payload for list fidelity.
 *
 * The classifier below tells them apart and says what each block contains, so
 * the extraction stage can mine each one before the cleanup stage removes it.
 */

export type ConditionalKind =
  | 'downlevel-hidden-open'
  | 'downlevel-revealed-open'
  | 'endif'
  | 'fragment-start'
  | 'fragment-end'
  | 'plain';

export type ConditionalPayload =
  | 'list-marker'
  | 'vml'
  | 'office-metadata'
  | 'word-settings'
  | 'style'
  | 'unsupported-office'
  | 'compatibility'
  | 'unknown';

export interface ConditionalCommentInfo {
  kind: ConditionalKind;
  /** The condition expression, e.g. `gte mso 9`, `!supportLists`, `!vml`. */
  condition?: string;
  /** What the block's content is, so callers know whether to mine it. */
  payload: ConditionalPayload;
  /** Inner markup for downlevel-hidden comments (empty for revealed ones). */
  content?: string;
  /** True when the block should be mined before removal. */
  extractable: boolean;
}

const OPEN_HIDDEN = /^\s*\[if\s+([^\]]*)\]>([\s\S]*?)<!\[endif\]\s*$/i;
const OPEN_HIDDEN_UNCLOSED = /^\s*\[if\s+([^\]]*)\]>([\s\S]*)$/i;
const OPEN_REVEALED = /^\s*\[if\s+([^\]]*)\]\s*$/i;
const CLOSE = /^\s*\[endif\]\s*$/i;

/**
 * Classify a comment node's data.
 *
 * `data` is what the DOM gives for `<!-- data -->`. For the revealed form
 * `<![if !supportLists]>` browsers produce a bogus comment whose data is
 * `[if !supportLists]`, which is why both forms land here.
 */
export function classifyComment(data: string): ConditionalCommentInfo {
  const text = data ?? '';

  if (/^\s*StartFragment\s*$/i.test(text)) {
    return { kind: 'fragment-start', payload: 'compatibility', extractable: false };
  }
  if (/^\s*EndFragment\s*$/i.test(text)) {
    return { kind: 'fragment-end', payload: 'compatibility', extractable: false };
  }
  if (CLOSE.test(text)) {
    return { kind: 'endif', payload: 'compatibility', extractable: false };
  }

  const hidden = OPEN_HIDDEN.exec(text) ?? OPEN_HIDDEN_UNCLOSED.exec(text);
  if (hidden) {
    const condition = (hidden[1] ?? '').trim();
    const content = hidden[2] ?? '';
    const payload = classifyPayload(condition, content);
    return {
      kind: 'downlevel-hidden-open',
      condition,
      payload,
      content,
      extractable: payload !== 'compatibility' && payload !== 'unknown',
    };
  }

  const revealed = OPEN_REVEALED.exec(text);
  if (revealed) {
    const condition = (revealed[1] ?? '').trim();
    return {
      kind: 'downlevel-revealed-open',
      condition,
      payload: classifyPayload(condition, ''),
      extractable: /supportlists/i.test(condition),
    };
  }

  return { kind: 'plain', payload: 'unknown', extractable: false };
}

function classifyPayload(condition: string, content: string): ConditionalPayload {
  const cond = condition.toLowerCase();
  if (cond.includes('supportlists')) return 'list-marker';
  if (cond.includes('vml')) return 'vml';

  if (/<o:documentproperties|<o:officedocumentsettings|<o:shapedefaults/i.test(content)) {
    return 'office-metadata';
  }
  if (/<w:worddocument|<w:latentstyles|<w:compat/i.test(content)) return 'word-settings';
  if (/<style[\s>]/i.test(content)) return 'style';
  if (/<v:(shape|shapetype|imagedata|rect|oval|group|line|roundrect|polyline|background)/i.test(content)) {
    return 'vml';
  }
  if (/<o:oleobject|<o:smarttagtype|<object[\s>]|<m:omath/i.test(content)) {
    return 'unsupported-office';
  }
  if (/^\s*$/.test(content)) return 'compatibility';
  return 'unknown';
}

/**
 * Extract every downlevel-hidden conditional block from a raw HTML string,
 * returning the inner markup of each.
 *
 * Done on raw text rather than the DOM because the content of a hidden
 * conditional comment is *not* parsed into elements — the DOM only has the
 * comment's text.
 */
export function extractHiddenConditionalBlocks(html: string): Array<{
  condition: string;
  content: string;
  payload: ConditionalPayload;
}> {
  const blocks: Array<{ condition: string; content: string; payload: ConditionalPayload }> = [];
  const pattern = /<!--\s*\[if\s+([^\]]*)\]>([\s\S]*?)<!\[endif\]\s*-->/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(html)) !== null) {
    const condition = (match[1] ?? '').trim();
    const content = match[2] ?? '';
    blocks.push({ condition, content, payload: classifyPayload(condition, content) });
  }
  return blocks;
}

/**
 * Pull `<style>` blocks out of raw HTML, including those hidden inside
 * conditional comments (Word puts the print stylesheet there).
 */
export function extractStyleBlocks(html: string): string[] {
  const styles: string[] = [];
  const pattern = /<style\b[^>]*>([\s\S]*?)<\/style\s*>/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(html)) !== null) {
    const body = match[1];
    if (body && body.trim()) styles.push(body);
  }
  return styles;
}

/**
 * `<![if !supportLists]>` / `<![endif]>` bracket the rendered list marker.
 * Returns true for a comment node that opens such a block.
 */
export function isListMarkerOpen(data: string): boolean {
  return /^\s*\[if\s+!supportlists\]\s*$/i.test(data ?? '');
}

/** Returns true for a comment node that closes a downlevel-revealed block. */
export function isEndIf(data: string): boolean {
  return CLOSE.test(data ?? '');
}
