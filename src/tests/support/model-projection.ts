import type { WordDocument } from '../../model/Document.js';
import type { WordBlock } from '../../model/Block.js';
import type { WordRun } from '../../model/Run.js';
import type { WordTable } from '../../model/Table.js';
import type { RunFormatting } from '../../model/Style.js';

/**
 * A stable projection of the canonical model, for golden fixtures.
 *
 * `expected-model.json` holds this rather than the whole `WordDocument`, for
 * two reasons:
 *
 *   - The full model embeds the raw payload and the entire parsed stylesheet.
 *     A fixture file would be mostly Word's CSS, and every irrelevant change
 *     to it would break every test.
 *   - A golden file is only useful if a human can read the diff. This
 *     projection is the set of facts the engine promises: block structure,
 *     list identity and markers, run boundaries and the formatting on them,
 *     table geometry, image resolution, links.
 *
 * Anything not projected here is still tested — by unit tests against the
 * module that produces it — but is not part of the golden contract.
 */

export interface ProjectedDocument {
  detection: { isWord: boolean; source: string; signalCount: number };
  lists: ProjectedListDefinition[];
  blocks: ProjectedBlock[];
  images: ProjectedImage[];
  hyperlinks: Array<{ id: string; href: string; rawHref: string; blocked?: boolean }>;
  diagnostics: Array<{ code: string; severity: string; fidelity: string; count: number }>;
}

export interface ProjectedListDefinition {
  listId: string;
  levels: Array<{
    level: number;
    type: string;
    numberFormat: string;
    levelText?: string;
    bulletGlyph?: string;
    bulletGlyphRaw?: string;
    bulletFont?: string;
    startAt?: number;
    textIndentPx?: number;
    marginLeftPx?: number;
    tabStopPx?: number;
  }>;
}

export type ProjectedBlock =
  | ProjectedParagraph
  | ProjectedTable
  | { type: 'page-break'; breakType: string }
  | { type: 'horizontal-rule' }
  | { type: 'image-block'; imageId: string; placement: string }
  | { type: 'unsupported'; objectType: string; code: string; fallbackText?: string }
  | { type: 'container'; role: string; blocks: ProjectedBlock[] };

export interface ProjectedParagraph {
  type: 'paragraph';
  text: string;
  styleName?: string;
  headingLevel?: number;
  empty?: boolean;
  alignment?: string;
  marginLeftPx?: number;
  textIndentPx?: number;
  listItem?: {
    listId: string;
    level: number;
    lfo?: string;
    instance?: string;
    restart?: boolean;
    startAt?: number;
    marker: {
      type: string;
      text?: string;
      glyph?: string;
      rawGlyph?: string;
      font?: string;
      numberFormat?: string;
      levelText?: string;
      source: string;
    };
  };
  runs: ProjectedRun[];
}

export interface ProjectedRun {
  type: string;
  text?: string;
  imageId?: string;
  link?: string;
  formatting?: Record<string, string | number | boolean>;
}

export interface ProjectedTable {
  type: 'table';
  depth: number;
  gridColumnCount: number;
  widthPercent?: number;
  columns: Array<{ widthPx?: number; widthPercent?: number }>;
  rows: Array<{
    section: string;
    cells: Array<{
      colSpan: number;
      rowSpan: number;
      header?: boolean;
      backgroundColor?: string;
      blocks: ProjectedBlock[];
    }>;
  }>;
}

export interface ProjectedImage {
  id: string;
  resolution: string;
  origin: string;
  originalSource: string;
  widthPx?: number;
  heightPx?: number;
  hasSource: boolean;
}

export function projectDocument(document: WordDocument): ProjectedDocument {
  return {
    detection: {
      isWord: document.detection.isWord,
      source: document.detection.source,
      signalCount: document.detection.signals.length,
    },
    lists: document.lists.map(projectListDefinition),
    blocks: document.blocks.map(projectBlock),
    images: Object.values(document.images).map(projectImage),
    hyperlinks: Object.values(document.hyperlinks).map((link) => ({
      id: link.id,
      href: link.href,
      rawHref: link.rawHref,
      ...(link.blocked ? { blocked: true } : {}),
    })),
    diagnostics: [...document.diagnostics]
      .map((d) => ({
        code: d.code,
        severity: d.severity,
        fidelity: d.fidelity,
        count: d.count ?? 1,
      }))
      .sort((a, b) => a.code.localeCompare(b.code)),
  };
}

function projectListDefinition(definition: WordDocument['lists'][number]): ProjectedListDefinition {
  return {
    listId: definition.listId,
    levels: definition.levels.map((level) => ({
      level: level.level,
      type: level.type,
      numberFormat: level.numberFormat,
      ...(level.levelText !== undefined ? { levelText: level.levelText } : {}),
      ...(level.bulletGlyph !== undefined ? { bulletGlyph: level.bulletGlyph } : {}),
      ...(level.bulletGlyphRaw !== undefined && level.bulletGlyphRaw !== level.bulletGlyph
        ? { bulletGlyphRaw: level.bulletGlyphRaw }
        : {}),
      ...(level.bulletFont !== undefined ? { bulletFont: level.bulletFont } : {}),
      ...(level.startAt !== undefined ? { startAt: level.startAt } : {}),
      ...(level.textIndent ? { textIndentPx: level.textIndent.px } : {}),
      ...(level.marginLeft ? { marginLeftPx: level.marginLeft.px } : {}),
      ...(level.tabStop ? { tabStopPx: level.tabStop.px } : {}),
    })),
  };
}

function projectBlock(block: WordBlock): ProjectedBlock {
  switch (block.type) {
    case 'paragraph': {
      const projected: ProjectedParagraph = {
        type: 'paragraph',
        text: runsText(block.runs),
        runs: block.runs.map(projectRun),
      };
      if (block.styleName) projected.styleName = block.styleName;
      if (block.headingLevel) projected.headingLevel = block.headingLevel;
      if (block.empty) projected.empty = true;
      if (block.formatting.alignment) projected.alignment = block.formatting.alignment;
      if (block.formatting.marginLeft) projected.marginLeftPx = block.formatting.marginLeft.px;
      if (block.formatting.textIndent) projected.textIndentPx = block.formatting.textIndent.px;
      if (block.listItem) {
        const item = block.listItem;
        projected.listItem = {
          listId: item.listId,
          level: item.level,
          ...(item.lfo ? { lfo: item.lfo } : {}),
          ...(item.listInstanceId ? { instance: item.listInstanceId } : {}),
          ...(item.restart ? { restart: true } : {}),
          ...(item.startAt !== undefined ? { startAt: item.startAt } : {}),
          marker: {
            type: item.marker.type,
            ...(item.marker.text !== undefined ? { text: item.marker.text } : {}),
            ...(item.marker.glyph !== undefined ? { glyph: item.marker.glyph } : {}),
            ...(item.marker.rawGlyph !== undefined && item.marker.rawGlyph !== item.marker.glyph
              ? { rawGlyph: item.marker.rawGlyph }
              : {}),
            ...(item.marker.font !== undefined ? { font: item.marker.font } : {}),
            ...(item.marker.numberFormat !== undefined
              ? { numberFormat: item.marker.numberFormat }
              : {}),
            ...(item.marker.levelText !== undefined ? { levelText: item.marker.levelText } : {}),
            source: item.marker.source,
          },
        };
      }
      return projected;
    }
    case 'table':
      return projectTable(block);
    case 'container':
      return { type: 'container', role: block.role, blocks: block.blocks.map(projectBlock) };
    case 'page-break':
      return { type: 'page-break', breakType: block.breakType };
    case 'horizontal-rule':
      return { type: 'horizontal-rule' };
    case 'image-block':
      return { type: 'image-block', imageId: block.imageId, placement: block.placement };
    case 'unsupported':
      return {
        type: 'unsupported',
        objectType: block.objectType,
        code: block.code,
        ...(block.fallbackText ? { fallbackText: block.fallbackText } : {}),
      };
    default:
      return { type: 'horizontal-rule' };
  }
}

function projectTable(table: WordTable): ProjectedTable {
  return {
    type: 'table',
    depth: table.depth,
    gridColumnCount: table.gridColumnCount,
    ...(table.widthPercent !== undefined ? { widthPercent: table.widthPercent } : {}),
    columns: table.columns.map((column) => ({
      ...(column.width ? { widthPx: column.width.px } : {}),
      ...(column.widthPercent !== undefined ? { widthPercent: column.widthPercent } : {}),
    })),
    rows: table.rows.map((row) => ({
      section: row.section,
      cells: row.cells.map((cell) => ({
        colSpan: cell.colSpan,
        rowSpan: cell.rowSpan,
        ...(cell.header ? { header: true } : {}),
        ...(cell.shading?.fill || cell.backgroundColor
          ? { backgroundColor: (cell.shading?.fill ?? cell.backgroundColor)!.hex }
          : {}),
        blocks: cell.blocks.map(projectBlock),
      })),
    })),
  };
}

function projectRun(run: WordRun): ProjectedRun {
  const projected: ProjectedRun = { type: run.type };
  if (run.type === 'text' || run.type === 'field' || run.type === 'note') projected.text = run.text;
  if (run.type === 'image') projected.imageId = run.imageId;
  if ('hyperlinkId' in run && run.hyperlinkId) projected.link = run.hyperlinkId;
  const formatting = projectFormatting(run.formatting);
  if (Object.keys(formatting).length > 0) projected.formatting = formatting;
  return projected;
}

/** Only the formatting a reader of a diff would care about. */
function projectFormatting(formatting: RunFormatting): Record<string, string | number | boolean> {
  const out: Record<string, string | number | boolean> = {};
  if (formatting.bold) out.bold = true;
  if (formatting.italic) out.italic = true;
  if (formatting.underline && formatting.underline !== 'none') out.underline = formatting.underline;
  if (formatting.strike) out.strike = true;
  if (formatting.smallCaps) out.smallCaps = true;
  if (formatting.allCaps) out.allCaps = true;
  if (formatting.color) out.color = formatting.color.hex;
  if (formatting.highlight) out.highlight = formatting.highlight.hex;
  if (formatting.fontFamily) out.fontFamily = formatting.fontFamily;
  if (formatting.fontSize) out.fontSizePx = formatting.fontSize.px;
  if (formatting.verticalAlign && formatting.verticalAlign !== 'baseline') {
    out.verticalAlign = formatting.verticalAlign;
  }
  if (formatting.letterSpacing) out.letterSpacingPx = formatting.letterSpacing.px;
  if (formatting.language) out.language = formatting.language;
  if (formatting.hidden) out.hidden = true;
  return out;
}

function projectImage(image: WordDocument['images'][string]): ProjectedImage {
  return {
    id: image.id,
    resolution: image.resolution,
    origin: image.origin,
    originalSource: image.originalSource,
    ...(image.width ? { widthPx: image.width.px } : {}),
    ...(image.height ? { heightPx: image.height.px } : {}),
    hasSource: image.src.length > 0,
  };
}

function runsText(runs: WordRun[]): string {
  let text = '';
  for (const run of runs) {
    if (run.type === 'text' || run.type === 'field' || run.type === 'note') text += run.text;
    else if (run.type === 'tab') text += '\t';
    else if (run.type === 'break') text += '\n';
    else if (run.type === 'image') text += '￼';
  }
  return text;
}
