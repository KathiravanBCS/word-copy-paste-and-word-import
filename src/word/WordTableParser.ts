import type { WordTable, WordTableColumn } from '../model/Table.js';
import type { WordTableRow } from '../model/TableRow.js';
import type { WordTableCell } from '../model/TableCell.js';
import type { WordBlock } from '../model/Block.js';
import type { ParagraphFormatting, RunFormatting } from '../model/Style.js';
import { attr, childNodesOf, isElement, tagNameOf } from '../util/dom.js';
import { DiagnosticCode } from '../diagnostics/UnsupportedFeature.js';
import { parseBoxShorthand, parsePercent, parseWordLength } from './WordLengthParser.js';
import {
  parseAlignment,
  parseBorders,
  parseParagraphFormattingFromCss,
  parseRunFormattingFromCss,
  parseShading,
} from './WordFormattingParser.js';
import { parseWordColor } from './WordColorParser.js';
import { resolveElementCss } from './WordParagraphParser.js';
import { chargeBlock, type WordParseContext } from './WordParseContext.js';

/**
 * Word tables.
 *
 * Word's table HTML is close to ordinary HTML, which makes it tempting to pass
 * through untouched. That fails in three ways worth building a parser for:
 *
 *   - **Nested tables.** Word nests a table inside a `<td>`; passing the
 *     markup through keeps it looking right but leaves the nested content
 *     unstructured, so a consumer cannot walk it.
 *   - **The grid.** Rows disagree about their cell count whenever cells are
 *     merged, and Word emits `mso-` hints rather than a `<colgroup>`. Without
 *     resolving the grid, column widths cannot be applied and a ragged table
 *     renders wrong.
 *   - **Borders and shading.** Word puts them partly on the table, partly on
 *     every cell, partly in `mso-border-alt`. They have to be collected into
 *     the model, not left in a soup of inline styles.
 *
 * Cell content is parsed as a full block list, so a list inside a cell is a
 * real list and a nested table is a real table.
 */

export interface TableParseOptions {
  depth: number;
  inheritedRun?: RunFormatting;
  inheritedParagraph?: ParagraphFormatting;
  /** Callback that parses block-level children — supplied by WordHtmlParser. */
  parseBlocks: (
    element: Element,
    ctx: WordParseContext,
    options: { depth: number; inheritedRun?: RunFormatting; inheritedParagraph?: ParagraphFormatting },
  ) => WordBlock[];
}

export function parseTable(
  element: Element,
  ctx: WordParseContext,
  options: TableParseOptions,
): WordTable {
  const { css } = resolveElementCss(element, ctx);
  const lengthOptions = { defaultUnit: 'pt' as const, fontSizePx: 16 };

  const table: WordTable = {
    type: 'table',
    rows: [],
    columns: [],
    depth: options.depth,
    gridColumnCount: 0,
  };

  const width =
    parseWordLength(css['width'], lengthOptions) ??
    parseWordLength(attr(element, 'width'), { defaultUnit: 'px' });
  if (width) table.width = width;
  const widthPercent = parsePercent(css['width']) ?? parsePercent(attr(element, 'width'));
  if (widthPercent !== undefined) table.widthPercent = widthPercent;

  const alignment = parseAlignment(css['text-align']) ?? parseAlignment(attr(element, 'align'));
  if (alignment) table.alignment = alignment;

  const indent =
    parseWordLength(css['margin-left'], lengthOptions) ??
    parseWordLength(css['mso-table-lspace'], lengthOptions);
  if (indent) table.indent = indent;

  const borders = parseBorders(css, lengthOptions);
  if (borders) table.borders = borders;
  const shading = parseShading(css);
  if (shading) table.shading = shading;

  const collapse = css['border-collapse'];
  if (collapse) table.borderCollapse = /collapse/i.test(collapse) ? 'collapse' : 'separate';

  // `mso-padding-alt` is the table-wide default cell padding; the HTML
  // `cellpadding` attribute is the legacy spelling of the same thing.
  const padding =
    parseBoxShorthand(css['mso-padding-alt'], lengthOptions) ??
    parseBoxShorthand(css['padding'], lengthOptions) ??
    (attr(element, 'cellpadding')
      ? parseBoxShorthand(attr(element, 'cellpadding'), { defaultUnit: 'px' })
      : undefined);
  if (padding) table.cellPadding = padding;

  const spacing =
    parseWordLength(css['border-spacing'], lengthOptions) ??
    parseWordLength(attr(element, 'cellspacing'), { defaultUnit: 'px' });
  if (spacing) table.cellSpacing = spacing;

  const styleName = css['mso-table-style-name'] ?? undefined;
  if (styleName) table.styleName = styleName;

  if (options.depth > 0) {
    ctx.diagnostics.info(
      DiagnosticCode.WORD_NESTED_TABLE,
      `A table nested ${options.depth} level(s) deep was parsed as a structured table rather than passed through as markup.`,
      { details: { depth: options.depth }, fidelity: 'EXACT' },
    );
  }
  if (options.depth >= ctx.limits.maxTableDepth) {
    ctx.diagnostics.error(
      DiagnosticCode.LIMIT_DEPTH_EXCEEDED,
      `Table nesting exceeded ${ctx.limits.maxTableDepth} levels; deeper tables were not parsed.`,
    );
    return table;
  }

  const caption = findCaption(element);
  if (caption) table.caption = caption;

  collectRows(element, ctx, options, table, 'body');
  resolveGrid(table, ctx);
  collectColumns(element, table, lengthOptions);

  return table;
}

function findCaption(table: Element): string | undefined {
  for (const child of childNodesOf(table)) {
    if (isElement(child) && tagNameOf(child) === 'caption') {
      const text = (child.textContent ?? '').replace(/\s+/g, ' ').trim();
      if (text) return text;
    }
  }
  return undefined;
}

function collectRows(
  parent: Element,
  ctx: WordParseContext,
  options: TableParseOptions,
  table: WordTable,
  section: WordTableRow['section'],
): void {
  for (const child of childNodesOf(parent)) {
    if (!isElement(child)) continue;
    const tag = tagNameOf(child);
    if (tag === 'thead') {
      collectRows(child, ctx, options, table, 'head');
    } else if (tag === 'tfoot') {
      collectRows(child, ctx, options, table, 'foot');
    } else if (tag === 'tbody') {
      collectRows(child, ctx, options, table, section);
    } else if (tag === 'tr') {
      const row = parseRow(child, ctx, options, section);
      if (row) table.rows.push(row);
    } else if (tag === 'table') {
      // A stray nested table that is not inside a cell; Word does this when a
      // table follows a floating one. Parse it as a cell-less sibling row so
      // the content is not lost.
      ctx.diagnostics.warn(
        DiagnosticCode.WORD_TABLE_GRID_REPAIRED,
        'A <table> appeared directly inside another table without an enclosing cell; it was hoisted into a single-cell row so its content is preserved.',
      );
      const nested = parseTable(child, ctx, { ...options, depth: options.depth + 1 });
      table.rows.push({
        type: 'row',
        section,
        cells: [{ type: 'cell', blocks: [nested], colSpan: 1, rowSpan: 1 }],
      });
    }
  }
}

function parseRow(
  element: Element,
  ctx: WordParseContext,
  options: TableParseOptions,
  section: WordTableRow['section'],
): WordTableRow | null {
  if (!chargeBlock(ctx)) return null;
  const { css } = resolveElementCss(element, ctx);
  const lengthOptions = { defaultUnit: 'pt' as const, fontSizePx: 16 };

  const row: WordTableRow = { type: 'row', cells: [], section };
  const height =
    parseWordLength(css['height'], lengthOptions) ??
    parseWordLength(attr(element, 'height'), { defaultUnit: 'px' });
  if (height) row.height = height;

  const heightRule = css['mso-height-rule'];
  if (heightRule) {
    row.heightRule = /exact/i.test(heightRule) ? 'exact' : /at-?least/i.test(heightRule) ? 'atLeast' : 'auto';
  }
  // Word marks a repeating header row with `mso-yfti-irow:-1` or
  // `mso-row-heading:yes` — either is a real header row.
  if (/yes/i.test(css['mso-row-heading'] ?? '') || css['mso-yfti-irow'] === '-1') {
    row.header = true;
  }
  const shading = parseShading(css);
  if (shading) row.shading = shading;

  for (const child of childNodesOf(element)) {
    if (!isElement(child)) continue;
    const tag = tagNameOf(child);
    if (tag !== 'td' && tag !== 'th') continue;
    const cell = parseCell(child, ctx, options, tag === 'th' || row.header === true);
    if (cell) row.cells.push(cell);
  }

  if (row.header === undefined && row.cells.length > 0 && row.cells.every((c) => c.header)) {
    row.header = true;
  }
  if (row.header && section === 'body') row.section = 'head';

  return row;
}

function parseCell(
  element: Element,
  ctx: WordParseContext,
  options: TableParseOptions,
  isHeader: boolean,
): WordTableCell | null {
  if (!chargeBlock(ctx)) return null;
  const { css } = resolveElementCss(element, ctx);
  const lengthOptions = { defaultUnit: 'pt' as const, fontSizePx: 16 };

  const cell: WordTableCell = {
    type: 'cell',
    blocks: [],
    colSpan: readSpan(attr(element, 'colspan')),
    rowSpan: readSpan(attr(element, 'rowspan')),
  };
  if (isHeader) cell.header = true;

  const width =
    parseWordLength(css['width'], lengthOptions) ??
    parseWordLength(attr(element, 'width'), { defaultUnit: 'px' });
  if (width) cell.width = width;
  const widthPercent = parsePercent(css['width']) ?? parsePercent(attr(element, 'width'));
  if (widthPercent !== undefined) cell.widthPercent = widthPercent;

  const borders = parseBorders(css, lengthOptions);
  if (borders) cell.borders = borders;
  const shading = parseShading(css);
  if (shading) cell.shading = shading;
  const background = parseWordColor(css['background-color'] ?? css['background'] ?? attr(element, 'bgcolor'));
  if (background && background.hex !== 'transparent') cell.backgroundColor = background;

  const padding =
    parseBoxShorthand(css['padding'], lengthOptions) ??
    parseBoxShorthand(css['mso-padding-alt'], lengthOptions);
  if (padding) cell.padding = padding;

  const verticalAlign = (css['vertical-align'] ?? attr(element, 'valign') ?? '').toLowerCase();
  if (verticalAlign === 'top' || verticalAlign === 'bottom') cell.verticalAlign = verticalAlign;
  else if (verticalAlign === 'middle' || verticalAlign === 'center') cell.verticalAlign = 'middle';

  // Cell-level run and paragraph defaults cascade into the cell's content.
  const inheritedRun = parseRunFormattingFromCss(css);
  // A cell fill is a cell fill. Letting it cascade as `highlight` would put a
  // highlighter pen behind every word in the cell.
  delete inheritedRun.highlight;
  const inheritedParagraph = parseParagraphFormattingFromCss(css);
  // Alignment belongs to the paragraphs inside; margins on the cell do not.
  delete inheritedParagraph.marginLeft;
  delete inheritedParagraph.marginRight;
  delete inheritedParagraph.spaceBefore;
  delete inheritedParagraph.spaceAfter;
  delete inheritedParagraph.borders;
  delete inheritedParagraph.shading;
  delete inheritedParagraph.backgroundColor;

  const cellAlign = parseAlignment(attr(element, 'align'));
  if (cellAlign) inheritedParagraph.alignment = cellAlign;

  const blockOptions: {
    depth: number;
    inheritedRun?: RunFormatting;
    inheritedParagraph?: ParagraphFormatting;
  } = { depth: options.depth + 1 };
  const mergedRun = { ...(options.inheritedRun ?? {}), ...inheritedRun };
  if (Object.keys(mergedRun).length > 0) blockOptions.inheritedRun = mergedRun;
  if (Object.keys(inheritedParagraph).length > 0) blockOptions.inheritedParagraph = inheritedParagraph;

  cell.blocks = options.parseBlocks(element, ctx, blockOptions);
  return cell;
}

function readSpan(value: string | undefined): number {
  if (!value) return 1;
  const n = Number.parseInt(value, 10);
  if (!Number.isFinite(n) || n < 1) return 1;
  return Math.min(n, 1000);
}

/**
 * Resolve the table's column grid.
 *
 * Walks the rows placing each cell into the first free slot, honouring
 * rowspans carried down from earlier rows. Ragged rows are reported and
 * padded, because a table whose rows disagree about their width cannot be
 * rendered predictably and silently dropping the difference loses content.
 */
export function resolveGrid(table: WordTable, ctx: WordParseContext): void {
  const occupied: number[] = []; // remaining rowspan per grid column
  let gridWidth = 0;

  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r]!;
    let column = 0;
    for (const cell of row.cells) {
      while ((occupied[column] ?? 0) > 0) column++;
      cell.gridColumn = column;
      cell.gridRow = r;
      for (let i = 0; i < cell.colSpan; i++) {
        occupied[column + i] = cell.rowSpan;
      }
      column += cell.colSpan;
    }
    gridWidth = Math.max(gridWidth, column, occupied.length);
    for (let i = 0; i < occupied.length; i++) {
      if ((occupied[i] ?? 0) > 0) occupied[i] = (occupied[i] ?? 0) - 1;
    }
  }

  table.gridColumnCount = gridWidth;

  // Report rows that do not fill the grid. The renderer pads them so the table
  // stays rectangular; the diagnostic says that happened.
  for (const row of table.rows) {
    const span = row.cells.reduce((sum, cell) => sum + cell.colSpan, 0);
    if (span > 0 && span < gridWidth) {
      const covered = row.cells.some((c) => (c.gridColumn ?? 0) + c.colSpan > span);
      if (!covered) {
        ctx.diagnostics.info(
          DiagnosticCode.WORD_TABLE_GRID_REPAIRED,
          `A table row spans ${span} of ${gridWidth} grid columns; the shortfall is padded on render so the table stays rectangular.`,
          { details: { rowSpanTotal: span, gridColumns: gridWidth }, fidelity: 'EQUIVALENT' },
        );
      }
    }
  }
}

/** Column widths from `<col>` elements, else from the widest row's cells. */
function collectColumns(
  element: Element,
  table: WordTable,
  lengthOptions: { defaultUnit: 'pt'; fontSizePx: number },
): void {
  const columns: WordTableColumn[] = [];
  const cols = findColElements(element);
  if (cols.length > 0) {
    for (const col of cols) {
      const column: WordTableColumn = {};
      const style = attr(col, 'style');
      const width =
        parseWordLength(style ? extractWidth(style) : undefined, lengthOptions) ??
        parseWordLength(attr(col, 'width'), { defaultUnit: 'px' });
      if (width) column.width = width;
      const percent = parsePercent(attr(col, 'width') ?? '');
      if (percent !== undefined) column.widthPercent = percent;
      const span = readSpan(attr(col, 'span'));
      for (let i = 0; i < span; i++) columns.push({ ...column });
    }
  } else {
    // Derive from the first row that has no merged cells: Word writes every
    // cell's width, so one clean row describes the whole grid.
    const clean = table.rows.find((row) => row.cells.every((c) => c.colSpan === 1));
    if (clean) {
      for (const cell of clean.cells) {
        const column: WordTableColumn = {};
        if (cell.width) column.width = cell.width;
        if (cell.widthPercent !== undefined) column.widthPercent = cell.widthPercent;
        columns.push(column);
      }
    }
  }
  while (columns.length < table.gridColumnCount) columns.push({});
  table.columns = columns.slice(0, Math.max(table.gridColumnCount, columns.length));
}

function findColElements(table: Element): Element[] {
  const cols: Element[] = [];
  for (const child of childNodesOf(table)) {
    if (!isElement(child)) continue;
    const tag = tagNameOf(child);
    if (tag === 'col') cols.push(child);
    else if (tag === 'colgroup') {
      for (const grandchild of childNodesOf(child)) {
        if (isElement(grandchild) && tagNameOf(grandchild) === 'col') cols.push(grandchild);
      }
    }
  }
  return cols;
}

function extractWidth(style: string): string | undefined {
  const match = /(?:^|;)\s*width\s*:\s*([^;]+)/i.exec(style);
  return match?.[1]?.trim();
}
