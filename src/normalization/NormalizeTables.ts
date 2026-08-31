import type { WordBlock } from '../model/Block.js';
import type { WordDocument } from '../model/Document.js';
import type { WordTable } from '../model/Table.js';
import type { DiagnosticCollector } from '../diagnostics/DiagnosticCollector.js';
import { DiagnosticCode } from '../diagnostics/UnsupportedFeature.js';
import { roundTo } from '../word/WordLengthParser.js';

/**
 * Table normalisation.
 *
 * The parser resolves each table's grid; this pass makes the grid *consistent*
 * across the document and fills in what Word left implicit:
 *
 *   - Rows that do not fill the grid get explicit padding cells, so a
 *     renderer never has to guess and a consumer walking the model always
 *     sees a rectangle.
 *   - Column widths get a percentage form, because Word writes absolute point
 *     widths that are meaningless once the content is in a page of a different
 *     width.
 *   - Header rows are promoted so `<thead>` can be emitted.
 */

export interface NormalizeTablesResult {
  tables: number;
  paddedRows: number;
  nested: number;
}

export function normalizeTables(
  document: WordDocument,
  diagnostics: DiagnosticCollector,
): NormalizeTablesResult {
  const result: NormalizeTablesResult = { tables: 0, paddedRows: 0, nested: 0 };

  const walk = (blocks: WordBlock[]): void => {
    for (const block of blocks) {
      if (block.type === 'table') {
        result.tables++;
        if (block.depth > 0) result.nested++;
        result.paddedRows += padRows(block, diagnostics);
        deriveColumnPercentages(block);
        promoteHeaderRow(block);
        for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
      } else if (block.type === 'container') {
        walk(block.blocks);
      }
    }
  };
  walk(document.blocks);
  return result;
}

/**
 * Give short rows explicit padding cells.
 *
 * A row is short when its cells plus the rowspans reaching into it from above
 * do not fill the grid. Word produces these when a table has merged cells and
 * the copy started part-way through, and a renderer that ignores the shortfall
 * produces a table with a ragged right edge.
 */
function padRows(table: WordTable, diagnostics: DiagnosticCollector): number {
  if (table.gridColumnCount === 0) return 0;
  let padded = 0;

  // Recompute occupancy so the padding accounts for rowspans from earlier rows.
  const carried: number[] = new Array(table.gridColumnCount).fill(0);
  for (const row of table.rows) {
    let occupied = 0;
    for (let c = 0; c < table.gridColumnCount; c++) {
      if ((carried[c] ?? 0) > 0) occupied++;
    }
    const own = row.cells.reduce((sum, cell) => sum + cell.colSpan, 0);
    const total = occupied + own;

    if (total < table.gridColumnCount) {
      const missing = table.gridColumnCount - total;
      for (let i = 0; i < missing; i++) {
        row.cells.push({ type: 'cell', blocks: [], colSpan: 1, rowSpan: 1, covered: false });
      }
      padded++;
    }

    // Advance the carry map for the next row.
    let column = 0;
    for (const cell of row.cells) {
      while ((carried[column] ?? 0) > 0) column++;
      for (let i = 0; i < cell.colSpan; i++) {
        carried[column + i] = cell.rowSpan;
      }
      column += cell.colSpan;
    }
    for (let c = 0; c < carried.length; c++) {
      if ((carried[c] ?? 0) > 0) carried[c] = (carried[c] ?? 0) - 1;
    }
  }

  if (padded > 0) {
    diagnostics.info(
      DiagnosticCode.WORD_TABLE_GRID_REPAIRED,
      `${padded} table row(s) did not fill the resolved ${table.gridColumnCount}-column grid; empty cells were added so the table stays rectangular. No content was removed.`,
      { details: { rows: padded, columns: table.gridColumnCount }, fidelity: 'EQUIVALENT' },
    );
  }
  return padded;
}

/**
 * Express column widths as percentages of the table.
 *
 * Word writes absolute widths in points, which is right for a fixed page and
 * wrong for a web page of any other width. The absolute value stays on the
 * model; the percentage is what makes the table usable in a flexible layout.
 */
function deriveColumnPercentages(table: WordTable): void {
  const widths = table.columns.map((column) => column.width?.px ?? 0);
  const total = widths.reduce((sum, width) => sum + width, 0);
  if (total <= 0) return;

  for (let i = 0; i < table.columns.length; i++) {
    const column = table.columns[i]!;
    if (column.widthPercent === undefined && column.width) {
      column.widthPercent = roundTo((column.width.px / total) * 100, 3);
    }
  }
}

/**
 * Mark the first row as a header when Word said so, or when every cell in it
 * is a `<th>`. Word's own signal is `mso-yfti-irow:-1`, already read by the
 * parser; this only settles the section assignment.
 */
function promoteHeaderRow(table: WordTable): void {
  const first = table.rows[0];
  if (!first) return;
  if (first.section === 'head') return;
  if (first.cells.length > 0 && first.cells.every((cell) => cell.header)) {
    first.section = 'head';
    first.header = true;
  }
}
