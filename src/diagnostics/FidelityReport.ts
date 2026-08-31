import type { WordDocument } from '../model/Document.js';
import type { WordBlock } from '../model/Block.js';
import { walkBlocks } from '../model/Block.js';
import type { WordDiagnostic, FidelityClass } from './UnsupportedFeature.js';

/**
 * A count of what survived the round trip, plus what did not.
 *
 * The report is how the engine keeps its promise not to claim "100% fidelity":
 * every approximation and every unsupported construct is visible here.
 */
export interface FidelityReport {
  wordDetected: boolean;
  detectionConfidence: number;
  detectionSignals: string[];

  blocks: number;
  paragraphs: number;
  headings: number;
  runs: number;
  characters: number;

  lists: number;
  listInstances: number;
  listItems: number;
  listLevels: number;
  /** Deepest list level index observed, zero-based. */
  maxListDepth: number;
  bulletItems: number;
  numberedItems: number;
  customGlyphs: string[];

  tables: number;
  nestedTables: number;
  tableRows: number;
  tableCells: number;
  mergedCells: number;

  images: number;
  resolvedImages: number;
  unresolvedImages: number;

  hyperlinks: number;
  bookmarks: number;
  pageBreaks: number;
  unsupportedObjects: number;

  /** Distinct diagnostic codes classified as APPROXIMATED or UNSUPPORTED. */
  unsupportedFeatures: string[];
  approximatedFeatures: string[];
  warnings: number;
  errors: number;
  /** Per-class diagnostic counts. */
  fidelityBreakdown: Record<FidelityClass, number>;
  diagnostics: WordDiagnostic[];
}

/** Build a FidelityReport by walking a finished document. */
export function buildFidelityReport(doc: WordDocument): FidelityReport {
  const report: FidelityReport = {
    wordDetected: doc.detection.isWord,
    detectionConfidence: doc.detection.confidence,
    detectionSignals: doc.detection.signals.slice(),

    blocks: 0,
    paragraphs: 0,
    headings: 0,
    runs: 0,
    characters: 0,

    lists: doc.lists.length,
    listInstances: 0,
    listItems: 0,
    listLevels: 0,
    maxListDepth: -1,
    bulletItems: 0,
    numberedItems: 0,
    customGlyphs: [],

    tables: 0,
    nestedTables: 0,
    tableRows: 0,
    tableCells: 0,
    mergedCells: 0,

    images: Object.keys(doc.images).length,
    resolvedImages: 0,
    unresolvedImages: 0,

    hyperlinks: Object.keys(doc.hyperlinks).length,
    bookmarks: Object.keys(doc.bookmarks).length,
    pageBreaks: 0,
    unsupportedObjects: 0,

    unsupportedFeatures: [],
    approximatedFeatures: [],
    warnings: 0,
    errors: 0,
    fidelityBreakdown: { EXACT: 0, EQUIVALENT: 0, APPROXIMATED: 0, UNSUPPORTED: 0 },
    diagnostics: doc.diagnostics,
  };

  for (const def of doc.lists) {
    report.listLevels += def.levels.length;
  }

  const instances = new Set<string>();
  const glyphs = new Set<string>();

  for (const block of walkBlocks(doc.blocks)) {
    report.blocks++;
    countBlock(block, report, instances, glyphs);
  }

  report.listInstances = instances.size;
  report.customGlyphs = [...glyphs].sort();

  for (const image of Object.values(doc.images)) {
    if (image.resolution === 'unresolved') report.unresolvedImages++;
    else report.resolvedImages++;
  }

  const unsupported = new Set<string>();
  const approximated = new Set<string>();
  for (const d of doc.diagnostics) {
    const n = d.count ?? 1;
    if (d.severity === 'warning') report.warnings += n;
    if (d.severity === 'error') report.errors += n;
    report.fidelityBreakdown[d.fidelity] += n;
    if (d.fidelity === 'UNSUPPORTED') unsupported.add(d.code);
    else if (d.fidelity === 'APPROXIMATED') approximated.add(d.code);
  }
  report.unsupportedFeatures = [...unsupported].sort();
  report.approximatedFeatures = [...approximated].sort();

  return report;
}

function countBlock(
  block: WordBlock,
  report: FidelityReport,
  instances: Set<string>,
  glyphs: Set<string>,
): void {
  switch (block.type) {
    case 'paragraph': {
      report.paragraphs++;
      if (block.headingLevel) report.headings++;
      for (const run of block.runs) {
        report.runs++;
        if (run.type === 'text' || run.type === 'field' || run.type === 'note') {
          report.characters += run.text.length;
        }
      }
      const item = block.listItem;
      if (item) {
        report.listItems++;
        if (item.listInstanceId) instances.add(item.listInstanceId);
        if (item.level > report.maxListDepth) report.maxListDepth = item.level;
        if (item.marker.type === 'bullet') {
          report.bulletItems++;
          if (item.marker.glyph) glyphs.add(item.marker.glyph);
        } else if (item.marker.type === 'number') {
          report.numberedItems++;
        }
      }
      break;
    }
    case 'table': {
      report.tables++;
      if (block.depth > 0) report.nestedTables++;
      for (const row of block.rows) {
        report.tableRows++;
        for (const cell of row.cells) {
          if (cell.covered) continue;
          report.tableCells++;
          if (cell.colSpan > 1 || cell.rowSpan > 1) report.mergedCells++;
        }
      }
      break;
    }
    case 'page-break':
      report.pageBreaks++;
      break;
    case 'unsupported':
      report.unsupportedObjects++;
      break;
    default:
      break;
  }
}

/** A compact, human-readable summary suitable for a console or the lab UI. */
export function formatFidelityReport(report: FidelityReport): string {
  const lines = [
    `Word detected: ${report.wordDetected} (confidence ${report.detectionConfidence.toFixed(2)})`,
    `Blocks ${report.blocks}  paragraphs ${report.paragraphs}  headings ${report.headings}  runs ${report.runs}`,
    `Lists ${report.lists} definitions / ${report.listInstances} instances / ${report.listItems} items ` +
      `(${report.bulletItems} bullet, ${report.numberedItems} numbered, max depth ${report.maxListDepth + 1})`,
    `Tables ${report.tables} (${report.nestedTables} nested, ${report.mergedCells} merged cells)`,
    `Images ${report.images} (${report.resolvedImages} resolved, ${report.unresolvedImages} unresolved)`,
    `Hyperlinks ${report.hyperlinks}  bookmarks ${report.bookmarks}  page breaks ${report.pageBreaks}`,
    `Fidelity: EXACT ${report.fidelityBreakdown.EXACT}, EQUIVALENT ${report.fidelityBreakdown.EQUIVALENT}, ` +
      `APPROXIMATED ${report.fidelityBreakdown.APPROXIMATED}, UNSUPPORTED ${report.fidelityBreakdown.UNSUPPORTED}`,
  ];
  if (report.unsupportedFeatures.length) {
    lines.push(`Unsupported: ${report.unsupportedFeatures.join(', ')}`);
  }
  if (report.approximatedFeatures.length) {
    lines.push(`Approximated: ${report.approximatedFeatures.join(', ')}`);
  }
  return lines.join('\n');
}
