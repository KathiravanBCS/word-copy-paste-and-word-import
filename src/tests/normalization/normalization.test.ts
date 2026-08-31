import { describe, expect, it } from 'vitest';
import { parseWordHtmlString } from '../../index.js';
import { buildListTree } from '../../normalization/NormalizeLists.js';
import { computeListIndentation } from '../../normalization/NormalizeUnits.js';
import { walkBlocks } from '../../model/Block.js';
import { loadFixtures } from '../support/fixtures.js';
import type { WordDocument } from '../../model/Document.js';

const fixtures = loadFixtures();
const byId = (id: string): WordDocument => {
  const fixture = fixtures.find((f) => f.id === id);
  if (!fixture) throw new Error(`Missing fixture: ${id}`);
  return parseWordHtmlString(fixture.inputHtml);
};

describe('list normalisation', () => {
  it('assigns one instance to a contiguous run of the same list', () => {
    const document = byId('numbering/simple-numbering');
    const items = [...walkBlocks(document.blocks)]
      .filter((b) => b.type === 'paragraph')
      .map((b) => (b.type === 'paragraph' ? b.listItem : undefined))
      .filter(Boolean);
    const instances = new Set(items.map((i) => i!.listInstanceId));
    expect(instances.size).toBe(1);
  });

  it('restarts a deeper level each time it is re-entered', () => {
    // 1 / 2 / a / b / i / 3 — the alpha level starts at "a" because it is
    // entered for the first time, and the roman level at "i".
    const document = byId('numbering/simple-numbering');
    const items = [...walkBlocks(document.blocks)]
      .filter((b) => b.type === 'paragraph')
      .map((b) => (b.type === 'paragraph' ? b.listItem : undefined))
      .filter(Boolean);
    expect(items.map((i) => i!.marker.text)).toEqual(['1.', '2.', 'a.', 'b.', 'i.', '3.']);
    expect(items.map((i) => i!.restart ?? false)).toEqual([true, false, true, false, true, false]);
  });

  it('carries numbering across an interruption for a continuing list', () => {
    const document = byId('mixed/list-continuation');
    const items = [...walkBlocks(document.blocks)]
      .filter((b) => b.type === 'paragraph')
      .map((b) => (b.type === 'paragraph' ? b.listItem : undefined))
      .filter(Boolean);
    expect(items.map((i) => i!.marker.text)).toEqual(['1.', '2.', '3.', '4.']);
    expect(new Set(items.map((i) => i!.listInstanceId)).size).toBe(1);
  });

  it('gives a list in each table cell its own scope', () => {
    const document = byId('tables/list-in-cell');
    const table = document.blocks.find((b) => b.type === 'table');
    if (table?.type !== 'table') throw new Error('no table');
    const [a, b] = table.rows[0]!.cells;
    const first = a!.blocks.find((x) => x.type === 'paragraph' && x.listItem);
    const second = b!.blocks.find((x) => x.type === 'paragraph' && x.listItem);
    if (first?.type !== 'paragraph' || second?.type !== 'paragraph') throw new Error('no items');
    expect(first.listItem!.listInstanceId).not.toBe(second.listItem!.listInstanceId);
    expect(second.listItem!.marker.text).toBe('1.');
  });

  it('builds a nesting tree from Word’s flat paragraph sequence', () => {
    const document = byId('numbering/multilevel-numbering');
    const tree = buildListTree(document.blocks, 0);
    expect(tree).not.toBeNull();
    const root = tree!.node;
    expect(root.level).toBe(0);
    expect(root.items).toHaveLength(2); // I. and II.
    const firstChildren = root.items[0]!.children;
    expect(firstChildren).toHaveLength(1);
    expect(firstChildren[0]!.items).toHaveLength(4); // 1.1 .. 1.4
    const grandchildren = firstChildren[0]!.items[3]!.children;
    expect(grandchildren).toHaveLength(1);
    expect(grandchildren[0]!.items).toHaveLength(1); // a.
    expect(tree!.nextIndex).toBe(7);
  });

  it('synthesises a spacer when Word jumps more than one level', () => {
    // Level 0 straight to level 2, which is legal in Word and must not be
    // flattened: the nesting depth is part of the document's meaning.
    const html = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office">
      <meta name=ProgId content=Word.Document>
      <style><!--
        @list l0 {mso-list-id:1;}
        @list l0:level1 {mso-level-number-format:roman-upper; mso-level-text:"%1\\.";}
        @list l0:level2 {mso-level-text:"%1\\.%2";}
        @list l0:level3 {mso-level-number-format:alpha-lower; mso-level-text:"%3\\.";}
      --></style>
      <body><div class=WordSection1>
      <p style='mso-list:l0 level1 lfo1'><![if !supportLists]><span style='mso-list:Ignore'>I.</span><![endif]>Top</p>
      <p style='mso-list:l0 level3 lfo1'><![if !supportLists]><span style='mso-list:Ignore'>a.</span><![endif]>Deep</p>
      </div></body></html>`;
    const document = parseWordHtmlString(html);
    const tree = buildListTree(document.blocks, 0)!;

    // The intermediate level is created and given a marker-less spacer item to
    // hold the deeper list, so the structure is
    //   ol > li "Top" > ol(level 1) > li(spacer) > ol(level 2) > li "Deep"
    // rather than "Deep" being flattened up to level 1.
    expect(tree.node.items).toHaveLength(1);
    const top = tree.node.items[0]!;
    expect(top.spacer).toBeUndefined();
    expect(top.children).toHaveLength(1);

    const intermediate = top.children[0]!;
    expect(intermediate.level).toBe(1);
    expect(intermediate.items).toHaveLength(1);
    expect(intermediate.items[0]!.spacer).toBe(true);

    const deepest = intermediate.items[0]!.children[0]!;
    expect(deepest.level).toBe(2);
    expect(deepest.items[0]!.item!.marker.text).toBe('a.');
  });
});

describe('indentation normalisation', () => {
  it('converts Word’s hanging indent into marker and text offsets', () => {
    // margin-left .5in, text-indent -.25in: marker at 0.25in, text at 0.5in.
    const document = byId('bullets/bullet-default');
    const first = document.blocks.find((b) => b.type === 'paragraph');
    if (first?.type !== 'paragraph' || !first.listItem) throw new Error('no list item');
    const indent = computeListIndentation(first.listItem, first.formatting);
    expect(indent.textOffsetPx).toBe(48);
    expect(indent.markerOffsetPx).toBe(24);
    expect(indent.hangingPx).toBe(24);
    expect(indent.explicit).toBe(true);
  });

  it('falls back to Word’s default ladder when nothing is declared', () => {
    const item = {
      listId: 'l0',
      level: 1,
      marker: { type: 'bullet' as const, source: 'list-definition' as const },
    };
    const indent = computeListIndentation(item, {});
    expect(indent.explicit).toBe(false);
    expect(indent.textOffsetPx).toBe(96);
  });
});

describe('table normalisation', () => {
  it('pads short rows so the grid stays rectangular', () => {
    const document = byId('tables/merged-cells');
    const table = document.blocks.find((b) => b.type === 'table');
    if (table?.type !== 'table') throw new Error('no table');
    for (const row of table.rows) {
      const occupied = row.cells.reduce((sum, cell) => sum + cell.colSpan, 0);
      // Every row either fills the grid or is completed by a rowspan above it.
      expect(occupied).toBeLessThanOrEqual(table.gridColumnCount);
    }
    expect(table.gridColumnCount).toBe(3);
  });

  it('derives percentage column widths from Word’s absolute ones', () => {
    const document = byId('tables/simple-table');
    const table = document.blocks.find((b) => b.type === 'table');
    if (table?.type !== 'table') throw new Error('no table');
    const percentages = table.columns.map((c) => c.widthPercent);
    expect(percentages).toEqual([50, 50]);
  });

  it('promotes Word’s repeating header row to a real header', () => {
    const document = byId('tables/simple-table');
    const table = document.blocks.find((b) => b.type === 'table');
    if (table?.type !== 'table') throw new Error('no table');
    expect(table.rows[0]!.section).toBe('head');
    expect(table.rows[1]!.section).toBe('body');
  });
});

describe('style normalisation', () => {
  it('folds away declarations that only restate the document default', () => {
    const document = byId('basic/plain-paragraphs');
    const first = document.blocks.find((b) => b.type === 'paragraph');
    if (first?.type !== 'paragraph') throw new Error('no paragraph');
    // MsoNormal is Calibri 11pt, so a plain run should carry neither.
    expect(first.runs[0]!.formatting.fontFamily).toBeUndefined();
    expect(first.runs[0]!.formatting.fontSize).toBeUndefined();
  });

  it('keeps a declaration that differs from the default', () => {
    const document = byId('formatting/character-formatting');
    const runs = [...walkBlocks(document.blocks)]
      .filter((b) => b.type === 'paragraph')
      .flatMap((b) => (b.type === 'paragraph' ? b.runs : []));
    expect(runs.some((r) => r.formatting.fontFamily === 'Courier New')).toBe(true);
  });

  it('merges adjacent runs with identical formatting but not across a boundary', () => {
    const document = byId('formatting/character-formatting');
    const last = [...walkBlocks(document.blocks)].filter((b) => b.type === 'paragraph').at(-1);
    if (last?.type !== 'paragraph') throw new Error('no paragraph');
    expect(last.runs.filter((r) => r.type === 'text')).toHaveLength(5);
  });
});

describe('structural normalisation', () => {
  it('unwraps Word’s WordSection div', () => {
    const document = byId('basic/plain-paragraphs');
    expect(document.blocks.some((b) => b.type === 'container')).toBe(false);
    expect(document.blocks.every((b) => b.type === 'paragraph')).toBe(true);
  });

  it('keeps a deliberately empty paragraph', () => {
    const document = byId('basic/plain-paragraphs');
    const empties = document.blocks.filter((b) => b.type === 'paragraph' && b.empty);
    expect(empties.length).toBeGreaterThan(0);
  });

  it('drops empty paragraphs only when asked to', () => {
    const fixture = fixtures.find((f) => f.id === 'basic/plain-paragraphs')!;
    const document = parseWordHtmlString(fixture.inputHtml, {
      normalize: { dropEmptyParagraphs: true },
    });
    expect(document.blocks.filter((b) => b.type === 'paragraph' && b.empty)).toHaveLength(0);
  });
});
