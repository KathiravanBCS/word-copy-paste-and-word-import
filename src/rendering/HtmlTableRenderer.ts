import type { WordTable } from '../model/Table.js';
import type { WordTableRow } from '../model/TableRow.js';
import type { WordTableCell } from '../model/TableCell.js';
import type { WordBlock } from '../model/Block.js';
import { escapeHtmlAttribute } from '../util/dom.js';
import {
  StyleBuilder,
  renderBorders,
  renderColor,
  renderLength,
  type StyleRenderOptions,
} from './HtmlStyleRenderer.js';
import { roundTo } from '../word/WordLengthParser.js';

/**
 * Table rendering.
 *
 * The model is already a resolved grid, so this is a straightforward
 * serialisation — with three decisions worth stating:
 *
 *   - `border-collapse: collapse` is emitted unless Word said otherwise,
 *     because that is what Word tables look like and the default (`separate`)
 *     puts a visible gap between every cell.
 *   - Column widths are emitted as percentages via `<colgroup>` when the
 *     model has them, so the table adapts to the page it lands in while
 *     keeping Word's proportions.
 *   - Header rows become a real `<thead>` with `<th scope="col">`, which is
 *     the difference between a table a screen reader can navigate and a grid
 *     of anonymous cells.
 */

export interface TableRenderOptions extends StyleRenderOptions {
  classPrefix?: string;
  /** Emit percentage column widths from the model. Default true. */
  emitColumnWidths?: boolean;
}

export interface TableRenderContext {
  renderBlocks: (blocks: WordBlock[]) => string;
}

export function renderTable(
  table: WordTable,
  context: TableRenderContext,
  options: TableRenderOptions = {},
): string {
  const prefix = options.classPrefix ?? 'wce';
  const style = new StyleBuilder();

  style.set('border-collapse', table.borderCollapse ?? 'collapse');
  if (table.widthPercent !== undefined) style.set('width', `${roundTo(table.widthPercent, 3)}%`);
  else if (table.width) style.set('width', renderLength(table.width, options));

  if (table.alignment === 'center') {
    style.set('margin-left', 'auto');
    style.set('margin-right', 'auto');
  } else if (table.alignment === 'right') {
    style.set('margin-left', 'auto');
  } else if (table.indent) {
    style.set('margin-left', renderLength(table.indent, options));
  }

  for (const [property, value] of renderBorders(table.borders, options)) style.set(property, value);
  if (table.shading?.fill) style.set('background-color', renderColor(table.shading.fill));
  if (table.borderCollapse === 'separate' && table.cellSpacing) {
    style.set('border-spacing', renderLength(table.cellSpacing, options));
  }

  const parts: string[] = [];
  parts.push(
    `<table class="${prefix}-table"${style.isEmpty ? '' : ` style="${escapeHtmlAttribute(style.toString())}"`}${
      options.includeWordMetadata && table.styleName
        ? ` data-word-table-style="${escapeHtmlAttribute(table.styleName)}"`
        : ''
    }>`,
  );

  if (table.caption) parts.push(`<caption>${escapeHtmlAttribute(table.caption)}</caption>`);

  if (options.emitColumnWidths !== false) {
    const colgroup = renderColgroup(table, options);
    if (colgroup) parts.push(colgroup);
  }

  const head = table.rows.filter((row) => row.section === 'head');
  const body = table.rows.filter((row) => row.section === 'body');
  const foot = table.rows.filter((row) => row.section === 'foot');

  if (head.length > 0) {
    parts.push('<thead>');
    for (const row of head) parts.push(renderRow(row, table, context, options, true));
    parts.push('</thead>');
  }
  if (body.length > 0 || head.length === 0) {
    parts.push('<tbody>');
    for (const row of body) parts.push(renderRow(row, table, context, options, false));
    parts.push('</tbody>');
  }
  if (foot.length > 0) {
    parts.push('<tfoot>');
    for (const row of foot) parts.push(renderRow(row, table, context, options, false));
    parts.push('</tfoot>');
  }

  parts.push('</table>');
  return parts.join('');
}

function renderColgroup(table: WordTable, options: TableRenderOptions): string {
  const columns = table.columns;
  if (columns.length === 0) return '';
  if (!columns.some((column) => column.widthPercent !== undefined || column.width)) return '';

  const cols = columns.map((column) => {
    const style = new StyleBuilder();
    if (column.widthPercent !== undefined) style.set('width', `${roundTo(column.widthPercent, 3)}%`);
    else if (column.width) style.set('width', renderLength(column.width, options));
    return style.isEmpty ? '<col>' : `<col style="${escapeHtmlAttribute(style.toString())}">`;
  });
  return `<colgroup>${cols.join('')}</colgroup>`;
}

function renderRow(
  row: WordTableRow,
  table: WordTable,
  context: TableRenderContext,
  options: TableRenderOptions,
  inHead: boolean,
): string {
  const style = new StyleBuilder();
  if (row.height) {
    style.set(row.heightRule === 'exact' ? 'height' : 'min-height', renderLength(row.height, options));
  }
  if (row.shading?.fill) style.set('background-color', renderColor(row.shading.fill));

  const cells = row.cells
    .filter((cell) => !cell.covered)
    .map((cell) => renderCell(cell, table, context, options, inHead || cell.header === true))
    .join('');

  return `<tr${style.isEmpty ? '' : ` style="${escapeHtmlAttribute(style.toString())}"`}>${cells}</tr>`;
}

function renderCell(
  cell: WordTableCell,
  table: WordTable,
  context: TableRenderContext,
  options: TableRenderOptions,
  asHeader: boolean,
): string {
  const tag = asHeader ? 'th' : 'td';
  const style = new StyleBuilder();

  if (cell.widthPercent !== undefined) style.set('width', `${roundTo(cell.widthPercent, 3)}%`);
  else if (cell.width) style.set('width', renderLength(cell.width, options));

  for (const [property, value] of renderBorders(cell.borders, options)) style.set(property, value);

  const background = cell.shading?.fill ?? cell.backgroundColor;
  if (background) style.set('background-color', renderColor(background));

  const padding = cell.padding ?? table.cellPadding;
  if (padding) {
    const box = [padding.top, padding.right, padding.bottom, padding.left]
      .map((side) => (side ? renderLength(side, options) : '0'))
      .join(' ');
    style.set('padding', box);
  }
  if (cell.verticalAlign) style.set('vertical-align', cell.verticalAlign);

  const attributes: string[] = [];
  if (cell.colSpan > 1) attributes.push(` colspan="${cell.colSpan}"`);
  if (cell.rowSpan > 1) attributes.push(` rowspan="${cell.rowSpan}"`);
  if (asHeader) attributes.push(' scope="col"');
  if (!style.isEmpty) attributes.push(` style="${escapeHtmlAttribute(style.toString())}"`);

  const content = context.renderBlocks(cell.blocks);
  return `<${tag}${attributes.join('')}>${content}</${tag}>`;
}

/** Default table styling added to the generated stylesheet. */
export function tableCss(prefix = 'wce'): string {
  return [
    `.${prefix}-table { max-width: 100%; }`,
    `.${prefix}-table td, .${prefix}-table th { vertical-align: top; }`,
    `.${prefix}-table > thead th { text-align: left; }`,
  ].join('\n');
}
