import { describe, expect, it } from 'vitest';
import { parseWordHtmlString, renderWordDocument } from '../../index.js';
import {
  formatHtmlForGolden,
  loadFixtures,
  readExpected,
  serialiseJson,
  shouldUpdateFixtures,
  writeExpected,
} from '../support/fixtures.js';
import { projectDocument } from '../support/model-projection.js';

/**
 * The golden tests.
 *
 * Each fixture is run through the whole pipeline — raw Word HTML, canonical
 * model, rendered HTML — and both intermediate results are compared against
 * checked-in files. That is the chain the spec asks for:
 *
 *     raw input -> model -> output
 *
 * A failure here is either a regression or an intentional change. There is no
 * third possibility, which is the point: `UPDATE_FIXTURES=1` re-blesses them,
 * and the diff is what you review before doing that.
 */

const fixtures = loadFixtures();

describe('word fixtures', () => {
  it('finds fixtures to run', () => {
    expect(fixtures.length).toBeGreaterThan(0);
  });

  for (const fixture of fixtures) {
    describe(fixture.id, () => {
      const document = parseWordHtmlString(fixture.inputHtml);
      const rendered = renderWordDocument(document, {
        includeWordMetadata: true,
        cssMode: 'separate',
      });

      it('matches the expected model', () => {
        const actual = serialiseJson(projectDocument(document));
        const expected = readExpected(fixture.expectedModelPath);

        if (expected === null || shouldUpdateFixtures()) {
          writeExpected(fixture.expectedModelPath, actual);
          expect(expected === null || shouldUpdateFixtures()).toBe(true);
          return;
        }
        expect(actual.trim()).toBe(expected.trim());
      });

      it('matches the expected html', () => {
        const actual = [
          formatHtmlForGolden(rendered.html),
          '',
          '<!-- generated stylesheet -->',
          '<style>',
          rendered.css,
          '</style>',
        ].join('\n');
        const expected = readExpected(fixture.expectedHtmlPath);

        if (expected === null || shouldUpdateFixtures()) {
          writeExpected(fixture.expectedHtmlPath, actual);
          expect(expected === null || shouldUpdateFixtures()).toBe(true);
          return;
        }
        expect(actual.trim()).toBe(expected.trim());
      });

      it('never leaves a list marker inside a text node', () => {
        // The single most important invariant in the engine, checked on every
        // fixture rather than only on the list ones.
        const offenders: string[] = [];
        const glyphs = /[\u2022\u25aa\u25cf\u25e6\u00b7\u00a7\uf000-\uf0ff]/;

        const walk = (blocks: typeof document.blocks): void => {
          for (const block of blocks) {
            if (block.type === 'paragraph') {
              if (!block.listItem) continue;
              for (const run of block.runs) {
                if (run.type !== 'text') continue;
                if (glyphs.test(run.text)) offenders.push(run.text);
                if (block.listItem.marker.text && run.text.startsWith(block.listItem.marker.text)) {
                  offenders.push(run.text);
                }
              }
            } else if (block.type === 'container') {
              walk(block.blocks);
            } else if (block.type === 'table') {
              for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
            }
          }
        };
        walk(document.blocks);
        expect(offenders).toEqual([]);
      });

      it('renders no Word-only markup', () => {
        // `mso-`, Office namespaces and conditional comments must not survive
        // into the output; they were mined for meaning, then removed.
        expect(rendered.html).not.toMatch(/\bmso-[a-z-]+\s*:/i);
        expect(rendered.html).not.toMatch(/<o:|<v:|<w:|<m:/i);
        expect(rendered.html).not.toMatch(/\[if\s|\[endif\]/i);
        expect(rendered.html).not.toMatch(/class="?Mso/i);
      });

      it('emits no script or event handlers', () => {
        expect(rendered.html).not.toMatch(/<script/i);
        expect(rendered.html).not.toMatch(/\son[a-z]+\s*=/i);
        expect(rendered.html).not.toMatch(/javascript:/i);
      });
    });
  }
});
