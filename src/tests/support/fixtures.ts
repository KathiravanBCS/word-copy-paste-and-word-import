import { readdirSync, readFileSync, statSync, writeFileSync, existsSync } from 'node:fs';
import { join, relative, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * The fixture library.
 *
 * A fixture is a directory holding three files:
 *
 *     input.html            the Word clipboard payload, byte for byte
 *     expected-model.json   the projected canonical model (see model-projection)
 *     expected.html         the rendered output
 *
 * The rule from the spec that this system exists to enforce: every bug found
 * in real Word content becomes a fixture. `docs/TESTING.md` describes how to
 * capture one from a real paste using the clipboard lab.
 *
 * Set `UPDATE_FIXTURES=1` to write the expected files from the current
 * behaviour. That is how a *new* fixture is blessed; for an existing one,
 * read the diff first — an unexpected change there is the test doing its job.
 */

const HERE = dirname(fileURLToPath(import.meta.url));
export const FIXTURE_ROOT = join(HERE, '..', '..', 'fixtures', 'word');

export interface Fixture {
  /** `lists/multilevel-numbering` */
  id: string;
  /** `lists` */
  category: string;
  /** `multilevel-numbering` */
  name: string;
  directory: string;
  inputHtml: string;
  expectedModelPath: string;
  expectedHtmlPath: string;
  /** A `notes.md` beside the fixture, when present: what it is testing and why. */
  notes?: string;
}

export function loadFixtures(root: string = FIXTURE_ROOT): Fixture[] {
  const fixtures: Fixture[] = [];
  if (!existsSync(root)) return fixtures;

  const walk = (directory: string): void => {
    const entries = readdirSync(directory);
    if (entries.includes('input.html')) {
      const id = relative(root, directory).split(/[\\/]/).join('/');
      const [category = 'root', ...rest] = id.split('/');
      const fixture: Fixture = {
        id,
        category,
        name: rest.join('/') || category,
        directory,
        inputHtml: readFileSync(join(directory, 'input.html'), 'utf8'),
        expectedModelPath: join(directory, 'expected-model.json'),
        expectedHtmlPath: join(directory, 'expected.html'),
      };
      const notesPath = join(directory, 'notes.md');
      if (existsSync(notesPath)) fixture.notes = readFileSync(notesPath, 'utf8');
      fixtures.push(fixture);
      return;
    }
    for (const entry of entries) {
      const path = join(directory, entry);
      if (statSync(path).isDirectory()) walk(path);
    }
  };

  walk(root);
  return fixtures.sort((a, b) => a.id.localeCompare(b.id));
}

export const shouldUpdateFixtures = (): boolean =>
  process.env.UPDATE_FIXTURES === '1' || process.env.UPDATE_FIXTURES === 'true';

export function readExpected(path: string): string | null {
  return existsSync(path) ? readFileSync(path, 'utf8') : null;
}

export function writeExpected(path: string, content: string): void {
  writeFileSync(path, content.endsWith('\n') ? content : `${content}\n`, 'utf8');
}

export function serialiseJson(value: unknown): string {
  return JSON.stringify(value, null, 2);
}

/**
 * Pretty-print rendered HTML so a golden diff is readable line by line.
 *
 * Purely cosmetic and applied to both sides of every comparison, so it cannot
 * mask a real difference.
 */
export function formatHtmlForGolden(html: string): string {
  const withBreaks = html
    .replace(/></g, '>\n<')
    .replace(/\n(<\/(?:strong|em|u|s|sup|sub|span|a)>)/g, '$1');

  const lines = withBreaks.split('\n');
  const out: string[] = [];
  let depth = 0;
  const blockTags = /^<\/?(div|p|h[1-6]|ol|ul|li|table|thead|tbody|tfoot|tr|td|th|colgroup|blockquote|aside|main|style)\b/;

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const isClose = trimmed.startsWith('</');
    if (isClose && blockTags.test(trimmed)) depth = Math.max(0, depth - 1);
    out.push('  '.repeat(depth) + trimmed);
    const isSelfClosing = /\/>$/.test(trimmed) || /^<(img|br|hr|col|input|meta)\b/.test(trimmed);
    const hasClose = new RegExp(`</${trimmed.match(/^<([a-z0-9-]+)/i)?.[1] ?? '$^'}>`).test(trimmed);
    if (!isClose && !isSelfClosing && !hasClose && blockTags.test(trimmed)) depth++;
  }
  return out.join('\n');
}
