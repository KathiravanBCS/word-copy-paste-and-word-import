import type { WordBlock } from '../model/Block.js';
import type { WordDocument } from '../model/Document.js';
import type { WordParagraph } from '../model/Paragraph.js';
import type { RunFormatting } from '../model/Style.js';
import type { WordRun } from '../model/Run.js';
import { runFormattingEquals } from '../word/WordFormattingParser.js';

/**
 * Style normalisation.
 *
 * Word restates the document default on almost every run: a paragraph of plain
 * text arrives as a dozen spans that each say `font-family:"Calibri",sans-serif;
 * font-size:11.0pt`. Keeping all of it produces output that is technically
 * faithful and practically unusable — every word carries its own font
 * declaration, and any editor the content lands in inherits the mess.
 *
 * So redundant declarations are folded away *against the document default*,
 * which is itself read from Word's own `p.MsoNormal` rule rather than assumed.
 * Nothing that differs from that default is touched. This is a lossless
 * transformation with respect to appearance: what is removed is exactly what
 * the default already supplies.
 */

export interface NormalizeStylesOptions {
  /**
   * Fold declarations that merely restate the document default. Default true.
   * Turn it off to keep the model byte-faithful to the payload.
   */
  foldDocumentDefaults?: boolean;
  /** Merge adjacent runs whose formatting is identical. Default true. */
  mergeAdjacentRuns?: boolean;
}

export interface DocumentDefaults {
  fontFamily?: string;
  fontSizePx?: number;
}

/** Read the document default formatting from Word's `MsoNormal` style. */
export function readDocumentDefaults(document: WordDocument): DocumentDefaults {
  const normal =
    document.styles.styles['msonormal'] ??
    document.styles.styles['element:body'] ??
    document.styles.styles['element:p'];
  const defaults: DocumentDefaults = {};
  if (!normal) return defaults;
  if (normal.run.fontFamily) defaults.fontFamily = normal.run.fontFamily;
  if (normal.run.fontSize) defaults.fontSizePx = normal.run.fontSize.px;
  return defaults;
}

export function normalizeStyles(
  document: WordDocument,
  options: NormalizeStylesOptions = {},
): DocumentDefaults {
  const defaults = readDocumentDefaults(document);
  const fold = options.foldDocumentDefaults !== false;
  const merge = options.mergeAdjacentRuns !== false;

  const walk = (blocks: WordBlock[]): void => {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        if (fold) foldParagraphDefaults(block, defaults);
        if (merge) block.runs = mergeRuns(block.runs);
      } else if (block.type === 'container') {
        walk(block.blocks);
      } else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
      }
    }
  };
  walk(document.blocks);
  return defaults;
}

function foldParagraphDefaults(paragraph: WordParagraph, defaults: DocumentDefaults): void {
  for (const run of paragraph.runs) {
    foldRunDefaults(run.formatting, defaults);
  }
}

function foldRunDefaults(formatting: RunFormatting, defaults: DocumentDefaults): void {
  if (
    defaults.fontFamily &&
    formatting.fontFamily &&
    formatting.fontFamily.toLowerCase() === defaults.fontFamily.toLowerCase()
  ) {
    delete formatting.fontFamily;
    delete formatting.fontFamilyRaw;
  }
  if (
    defaults.fontSizePx !== undefined &&
    formatting.fontSize &&
    Math.abs(formatting.fontSize.px - defaults.fontSizePx) < 0.01
  ) {
    delete formatting.fontSize;
  }
  // `bold:false` etc. carry no information once nothing above them set it.
  if (formatting.bold === false) delete formatting.bold;
  if (formatting.italic === false) delete formatting.italic;
  if (formatting.underline === 'none') delete formatting.underline;
  if (formatting.verticalAlign === 'baseline') delete formatting.verticalAlign;
  if (formatting.strike === false) delete formatting.strike;
}

/**
 * Merge adjacent runs whose formatting is identical.
 *
 * This is not "flattening formatting across the paragraph" — the check is
 * structural equality of every formatting property, so a bold word still
 * starts a new run. What it removes is Word's habit of splitting one styled
 * phrase across several spans for reasons of its own (a spell-check boundary,
 * a proofing-language marker, a smart tag).
 */
export function mergeRuns(runs: WordRun[]): WordRun[] {
  if (runs.length < 2) return runs;
  const out: WordRun[] = [];
  for (const run of runs) {
    const previous = out[out.length - 1];
    if (
      previous &&
      previous.type === 'text' &&
      run.type === 'text' &&
      previous.hyperlinkId === run.hyperlinkId &&
      !run.bookmarks &&
      runFormattingEquals(previous.formatting, run.formatting) &&
      sameStyleChain(previous.styleChain, run.styleChain)
    ) {
      previous.text += run.text;
      continue;
    }
    out.push(run);
  }
  return out;
}

function sameStyleChain(a: string[] | undefined, b: string[] | undefined): boolean {
  if (a === b) return true;
  if (!a || !b) return (a?.length ?? 0) === (b?.length ?? 0);
  if (a.length !== b.length) return false;
  return a.every((value, index) => value === b[index]);
}
