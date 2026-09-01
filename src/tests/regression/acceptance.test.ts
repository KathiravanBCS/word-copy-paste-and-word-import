import { describe, expect, it } from 'vitest';
import { parseWordHtmlString, renderWordDocument, getFidelityReport } from '../../index.js';
import { loadFixtures } from '../support/fixtures.js';
import { walkBlocks } from '../../model/Block.js';
import type { WordDocument } from '../../model/Document.js';

/**
 * The non-negotiable acceptance criteria from the specification, as tests.
 *
 * Passing the golden fixtures is not on its own evidence that the engine does
 * what it was built to do — a golden file records whatever the engine
 * currently produces. These assertions record what it is *required* to
 * produce, independently of any fixture's blessed output.
 */

const fixtures = loadFixtures();
const byId = (id: string): WordDocument => {
  const fixture = fixtures.find((f) => f.id === id);
  if (!fixture) throw new Error(`Missing fixture: ${id}`);
  return parseWordHtmlString(fixture.inputHtml);
};

function paragraphs(document: WordDocument) {
  return [...walkBlocks(document.blocks)].filter((b) => b.type === 'paragraph');
}

function textOf(document: WordDocument): string[] {
  return paragraphs(document).map((p) =>
    p.type === 'paragraph' ? p.runs.map((r) => (r.type === 'text' ? r.text : '')).join('') : '',
  );
}

function markers(document: WordDocument) {
  return paragraphs(document)
    .map((p) => (p.type === 'paragraph' ? p.listItem : undefined))
    .filter((item): item is NonNullable<typeof item> => item !== undefined);
}

describe('acceptance criteria', () => {
  it('1. Word bullet markers are not text', () => {
    const document = byId('bullets/bullet-default');
    expect(textOf(document)).toEqual(['Parent', 'Child', 'Grandchild']);
    for (const item of markers(document)) {
      expect(item.marker.type).toBe('bullet');
      expect(item.marker.glyph).toBeTruthy();
    }
  });

  it('2. Word numbered markers are not text', () => {
    const document = byId('numbering/multilevel-numbering');
    expect(textOf(document)).toEqual([
      'Background',
      'Orbis India',
      'We understand',
      'VSTN',
      'Given',
      'A sub-point',
      'Scope',
    ]);
    expect(markers(document).map((i) => i.marker.text)).toEqual([
      'I.',
      '1.1',
      '1.2',
      '1.3',
      '1.4',
      'a.',
      'II.',
    ]);
  });

  it('3. Word custom bullet glyphs are preserved', () => {
    const document = byId('bullets/bullet-custom');
    const glyphs = markers(document).map((i) => i.marker.glyph);
    // Wingdings 0xA8, Wingdings 0xD8, Symbol 0xE0 — the glyphs Word drew, not
    // "standard" bullets substituted for them.
    expect(glyphs).toEqual(['▫', '➔', '◊']);
    for (const item of markers(document)) {
      expect(item.marker.rawGlyph).toBeTruthy();
      expect(item.marker.font).toBeTruthy();
    }
  });

  it('3b. the three Word default bullets map to • o ▪', () => {
    const document = byId('bullets/bullet-default');
    expect(markers(document).map((i) => i.marker.glyph)).toEqual(['•', 'o', '▪']);
    expect(markers(document).map((i) => i.marker.font)).toEqual([
      'Symbol',
      'Courier New',
      'Wingdings',
    ]);
  });

  it('4. Word multilevel numbering is preserved', () => {
    const document = byId('mixed/mixed-number-bullet');
    expect(markers(document).map((i) => i.level)).toEqual([0, 1, 2, 3]);
    expect(markers(document).map((i) => i.marker.type)).toEqual([
      'number',
      'number',
      'bullet',
      'number',
    ]);
    expect(textOf(document)).toEqual(['Parent', 'Child', 'Grandchild', 'Great-grandchild']);
  });

  it('5. Word numbering formats are preserved', () => {
    const document = byId('numbering/simple-numbering');
    const formats = markers(document).map((i) => i.levelDefinition?.numberFormat);
    expect(formats).toEqual([
      'decimal',
      'decimal',
      'lower-alpha',
      'lower-alpha',
      'lower-roman',
      'decimal',
    ]);
  });

  it('5b. Word level text is never rewritten into another scheme', () => {
    const document = byId('numbering/multilevel-numbering');
    const levelTexts = document.lists[0]!.levels.map((l) => l.levelText);
    expect(levelTexts[0]).toBe('%1.');
    expect(levelTexts[1]).toBe('%1.%2');
    const html = renderWordDocument(document).html;
    expect(html).not.toMatch(/Article\s+[IVX]/);
    expect(html).not.toMatch(/Section\s+\d+\.\d+/);
    // Default (element) mode: the literal marker Word computed is the visible
    // text of the marker span.
    expect(html).toContain('>1.1</span>');
    // Native mode: %1.%2 has no CSS counter-style equivalent, and falls back
    // to the same element marker rather than a data-marker-driven ::marker
    // (which does not survive a real editor's content model — see
    // HtmlListRenderer.ts). The literal text is still never rewritten.
    const native = renderWordDocument(document, { markerMode: 'native' }).html;
    expect(native).toContain('>1.1</span>');
  });

  it('6. Word list start values are preserved', () => {
    const document = byId('numbering/start-at');
    expect(document.lists[0]!.levels[0]!.startAt).toBe(5);
    expect(markers(document)[0]!.startAt).toBe(5);
    expect(markers(document).map((i) => i.marker.text)).toEqual(['5.', '6.', '7.']);
    // Default (element) mode: the start-at value is baked into the marker's
    // own visible text, so it renders correctly with no <li value> needed.
    expect(renderWordDocument(document).html).toContain('>5.</span>');
    // Native mode: the same value drives a real <li value> and <ol start>.
    const native = renderWordDocument(document, { markerMode: 'native' }).html;
    expect(native).toContain('value="5"');
    expect(native).toContain('start="5"');
  });

  it('6b. a new list restarts and a continuing list does not', () => {
    const restarted = byId('numbering/restarted-numbering');
    const instances = markers(restarted).map((i) => i.listInstanceId);
    expect(new Set(instances).size).toBe(2);

    const continued = byId('mixed/list-continuation');
    const continuedInstances = markers(continued).map((i) => i.listInstanceId);
    expect(new Set(continuedInstances).size).toBe(1);
    expect(markers(continued).map((i) => i.marker.text)).toEqual(['1.', '2.', '3.', '4.']);
  });

  it('7. Word indentation is preserved', () => {
    const document = byId('bullets/bullet-default');
    const items = markers(document);
    // Word: margin-left .5in / 1.0in / 1.5in with a -.25in hanging indent.
    expect(items.map((i) => i.marginLeft?.px ?? 0)).toEqual([48, 96, 144]);
    expect(items.map((i) => i.textIndent?.px)).toEqual([-24, -24, -24]);
  });

  it('8. text formatting is preserved run by run', () => {
    const document = byId('formatting/character-formatting');
    const last = paragraphs(document).at(-1);
    expect(last?.type).toBe('paragraph');
    if (last?.type !== 'paragraph') return;

    const runs = last.runs.filter((r) => r.type === 'text');
    expect(runs.map((r) => (r.type === 'text' ? r.text : ''))).toEqual([
      'Hello ',
      'world',
      ' and ',
      'goodbye',
      '.',
    ]);
    expect(runs[1]!.formatting.bold).toBe(true);
    expect(runs[1]!.formatting.color?.hex).toBe('#c00000');
    expect(runs[3]!.formatting.italic).toBe(true);
    expect(runs[0]!.formatting.bold).toBeUndefined();
  });

  it('8b. every character format reaches the model', () => {
    const document = byId('formatting/character-formatting');
    const all = paragraphs(document).flatMap((p) => (p.type === 'paragraph' ? p.runs : []));
    const formats = all.map((r) => r.formatting);
    expect(formats.some((f) => f.bold)).toBe(true);
    expect(formats.some((f) => f.italic)).toBe(true);
    expect(formats.some((f) => f.underline === 'single')).toBe(true);
    expect(formats.some((f) => f.strike)).toBe(true);
    expect(formats.some((f) => f.color?.hex === '#ff0000')).toBe(true);
    expect(formats.some((f) => f.highlight?.hex === '#ffff00')).toBe(true);
    expect(formats.some((f) => f.fontFamily === 'Courier New')).toBe(true);
    expect(formats.some((f) => f.fontSize && f.fontSize.px > 20)).toBe(true);
    expect(formats.some((f) => f.verticalAlign === 'super')).toBe(true);
    expect(formats.some((f) => f.verticalAlign === 'sub')).toBe(true);
    expect(formats.some((f) => f.letterSpacing)).toBe(true);
  });

  it('9. tables remain structured', () => {
    const merged = byId('tables/merged-cells');
    const table = merged.blocks.find((b) => b.type === 'table');
    expect(table?.type).toBe('table');
    if (table?.type !== 'table') return;
    expect(table.gridColumnCount).toBe(3);
    expect(table.rows[0]!.cells[0]!.colSpan).toBe(3);
    expect(table.rows[1]!.cells[0]!.rowSpan).toBe(2);
    expect(table.rows[2]!.cells[0]!.colSpan).toBe(2);

    const nested = byId('tables/nested-table');
    const outer = nested.blocks.find((b) => b.type === 'table');
    if (outer?.type !== 'table') throw new Error('no outer table');
    const inner = outer.rows[0]!.cells[0]!.blocks.find((b) => b.type === 'table');
    expect(inner?.type).toBe('table');
    if (inner?.type === 'table') expect(inner.depth).toBe(1);
  });

  it('9b. a list inside a table cell stays a list', () => {
    const document = byId('tables/list-in-cell');
    const table = document.blocks.find((b) => b.type === 'table');
    if (table?.type !== 'table') throw new Error('no table');
    const [bulletCell, numberCell] = table.rows[0]!.cells;
    const bulletItems = bulletCell!.blocks.filter(
      (b) => b.type === 'paragraph' && b.listItem,
    );
    expect(bulletItems).toHaveLength(2);
    const numberItems = numberCell!.blocks.filter((b) => b.type === 'paragraph' && b.listItem);
    expect(numberItems).toHaveLength(2);

    const html = renderWordDocument(document).html;
    expect(html).toContain('<ul');
    expect(html).toContain('<ol');
  });

  it('10. images are resolved where the clipboard permits, and only then', () => {
    const resolved = byId('images/inline-image');
    const resolvedImage = Object.values(resolved.images)[0]!;
    expect(resolvedImage.resolution).toBe('resolved');
    expect(resolvedImage.origin).toBe('data-uri');
    expect(renderWordDocument(resolved).html).toContain('<img');

    const unresolved = byId('images/unresolved-image');
    const unresolvedImage = Object.values(unresolved.images)[0]!;
    expect(unresolvedImage.resolution).toBe('unresolved');
    expect(unresolvedImage.originalSource).toContain('clip_image001.png');
    const html = renderWordDocument(unresolved).html;
    // The whole point: no <img> pointing at the author's disk. The reference
    // itself is still carried, on data-word-source, so nothing is lost — it
    // just is not something the browser will try to load.
    expect(html).not.toMatch(/<img\b/);
    expect(html).not.toMatch(/src="file:/);
    expect(html).toContain('image-placeholder');
    expect(html).toContain('data-word-source="file:///');
    expect(Object.keys(unresolved.images)).toHaveLength(1);
  });

  it('11. unsupported content is diagnosed rather than silently deleted', () => {
    const document = byId('images/unresolved-image');
    const codes = document.diagnostics.map((d) => d.code);
    expect(codes).toContain('WORD_LOCAL_FILE_IMAGE');
    expect(codes).toContain('WORD_UNRESOLVED_IMAGE');
    const report = getFidelityReport(document);
    expect(report.approximatedFeatures.length + report.unsupportedFeatures.length).toBeGreaterThan(0);
  });

  it('12. raw Word HTML remains available for debugging', () => {
    const fixture = fixtures.find((f) => f.id === 'complex/engagement-report')!;
    const document = parseWordHtmlString(fixture.inputHtml);
    expect(document.rawHtml).toBe(fixture.inputHtml);
    expect(document.styles.rawCss).toContain('@list');
    expect(document.detection.signals.length).toBeGreaterThan(3);
  });

  it('13. the model carries no editor-specific concepts', () => {
    const document = byId('complex/engagement-report');
    const serialised = JSON.stringify({
      blocks: document.blocks,
      lists: document.lists,
      images: document.images,
    });
    for (const editor of ['rooster', 'ckeditor', 'quill', 'tinymce', 'froala', 'prosemirror', 'lexical']) {
      expect(serialised.toLowerCase()).not.toContain(editor);
    }
    expect(serialised).not.toContain('OUTLINE_SCHEME');
  });

  it('14. the renderer invents no numbering semantics', () => {
    const document = byId('numbering/roman-numbering');
    // Default (element) mode: the markers come from Word's own definition
    // (upper-roman, "." suffix) rendered as the literal computed text —
    // I. / II. / III., never Word's numbers replaced by anything else, and
    // never welded into the paragraph's own text content.
    const { html } = renderWordDocument(document);
    expect(html).toContain('>I.</span>');
    expect(html).toContain('>II.</span>');
    expect(html).toContain('>III.</span>');
    expect(html).not.toMatch(/>I\.\s*Introduction/);

    // Native mode: the same definition instead drives a real @counter-style
    // and <li value>, still without inventing a scheme of its own.
    const native = renderWordDocument(document, { markerMode: 'native' });
    expect(native.css).toContain('system: extends upper-roman');
    expect(native.css).toContain('suffix: ". "');
    expect(native.html).toContain('value="1"');
    expect(native.html).toContain('value="2"');
    expect(native.html).toContain('value="3"');
    expect(native.html).not.toMatch(/>I\.\s*Introduction/);
  });

  it('15. no application numbering scheme replaces Word’s', () => {
    for (const fixture of fixtures) {
      const document = parseWordHtmlString(fixture.inputHtml);
      for (const item of markers(document)) {
        if (!item.levelDefinition?.levelText) continue;
        // The model's level text is byte-identical to what Word declared.
        expect(item.marker.levelText, fixture.id).toBe(item.levelDefinition.levelText);
      }
    }
  });

  it('16. the engine has no runtime dependencies', async () => {
    const packageJson = await import('../../../package.json', { with: { type: 'json' } });
    // Tolerant of both spellings: npm strips an empty `dependencies` key on
    // uninstall, and "absent" means the same thing as "empty" here.
    const dependencies = (packageJson.default as { dependencies?: Record<string, string> })
      .dependencies;
    expect(Object.keys(dependencies ?? {})).toEqual([]);
  });

  it('bonus: bullet glyphs never appear in any text node, across all fixtures', () => {
    const glyphPattern = /[•▪●▫◊➔-]/;
    for (const fixture of fixtures) {
      const document = parseWordHtmlString(fixture.inputHtml);
      for (const block of walkBlocks(document.blocks)) {
        if (block.type !== 'paragraph') continue;
        for (const run of block.runs) {
          if (run.type !== 'text') continue;
          expect(glyphPattern.test(run.text), `${fixture.id}: ${JSON.stringify(run.text)}`).toBe(
            false,
          );
        }
      }
    }
  });
});
