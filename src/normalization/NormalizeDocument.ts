import type { WordBlock } from '../model/Block.js';
import type { WordDocument } from '../model/Document.js';
import type { WordParagraph } from '../model/Paragraph.js';
import { DiagnosticCollector } from '../diagnostics/DiagnosticCollector.js';
import { normalizeLists, type NormalizeListsResult } from './NormalizeLists.js';
import { normalizeStyles, type DocumentDefaults, type NormalizeStylesOptions } from './NormalizeStyles.js';
import { normalizeTables, type NormalizeTablesResult } from './NormalizeTables.js';
import { normalizeImages, type NormalizeImagesResult } from './NormalizeImages.js';

/**
 * The normalisation pass.
 *
 * Everything here operates on the model, never on markup. The parser's job was
 * to understand the payload; this pass's job is to make the result coherent —
 * list identity resolved, tables rectangular, images sized, structural noise
 * removed — without inventing anything the payload did not say.
 *
 * Order matters:
 *
 *   1. Structural cleanup, so later passes see the final block sequence.
 *   2. Lists, which depend on block adjacency.
 *   3. Styles, tables and images, which are independent of each other.
 */

export interface NormalizeOptions extends NormalizeStylesOptions {
  /**
   * Drop paragraphs that Word emitted purely as spacing artefacts. Off by
   * default: an empty paragraph in Word is usually a deliberate blank line,
   * and removing it changes the document.
   */
  dropEmptyParagraphs?: boolean;
  /** Unwrap containers that carry no formatting of their own. Default true. */
  unwrapRedundantContainers?: boolean;
  /** Convert break-only paragraphs into page-break blocks. Default true. */
  liftPageBreaks?: boolean;
}

export interface NormalizeResult {
  lists: NormalizeListsResult;
  tables: NormalizeTablesResult;
  images: NormalizeImagesResult;
  defaults: DocumentDefaults;
}

export function normalizeDocument(
  document: WordDocument,
  options: NormalizeOptions = {},
): NormalizeResult {
  // Diagnostics raised here join the ones the parser already produced, so the
  // fidelity report covers the whole pipeline rather than just parsing.
  const diagnostics = new DiagnosticCollector();

  document.blocks = normalizeStructure(document.blocks, options);

  const lists = normalizeLists(document, diagnostics);
  const defaults = normalizeStyles(document, options);
  const tables = normalizeTables(document, diagnostics);
  const images = normalizeImages(document, diagnostics);

  for (const diagnostic of diagnostics.all()) document.diagnostics.push(diagnostic);

  return { lists, tables, images, defaults };
}

/* -------------------------------------------------------------------------
 * Structural cleanup
 * ---------------------------------------------------------------------- */

function normalizeStructure(blocks: WordBlock[], options: NormalizeOptions): WordBlock[] {
  const out: WordBlock[] = [];

  for (const block of blocks) {
    switch (block.type) {
      case 'container': {
        block.blocks = normalizeStructure(block.blocks, options);
        if (options.unwrapRedundantContainers !== false && isRedundantContainer(block)) {
          out.push(...block.blocks);
          continue;
        }
        out.push(block);
        continue;
      }
      case 'table': {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            cell.blocks = normalizeStructure(cell.blocks, options);
          }
        }
        out.push(block);
        continue;
      }
      case 'paragraph': {
        if (options.liftPageBreaks !== false) {
          const lifted = liftPageBreak(block);
          if (lifted) {
            out.push(lifted);
            continue;
          }
        }
        if (options.dropEmptyParagraphs && block.empty && !block.listItem) continue;
        out.push(block);
        continue;
      }
      default:
        out.push(block);
    }
  }
  return out;
}

/**
 * A container earns its place only if it carries formatting or a role that
 * changes the meaning of what is inside it. Word's `div.WordSection1` wrapper
 * and the anonymous divs it uses for structure do not.
 */
function isRedundantContainer(block: Extract<WordBlock, { type: 'container' }>): boolean {
  if (block.role === 'blockquote' || block.role === 'text-box') return false;
  const formatting = block.formatting;
  const carriesFormatting =
    formatting.borders !== undefined ||
    formatting.shading !== undefined ||
    formatting.backgroundColor !== undefined ||
    formatting.alignment !== undefined ||
    formatting.marginLeft !== undefined ||
    formatting.marginRight !== undefined;
  return !carriesFormatting;
}

/**
 * A paragraph whose only content is an explicit page break is a page break.
 *
 * Word writes one as:
 *
 *     <p class=MsoNormal><span style='mso-special-character:line-break;
 *        page-break-before:always'></span></p>
 *
 * Leaving it as a paragraph puts a stray empty line in the output where the
 * document meant "new page".
 */
function liftPageBreak(paragraph: WordParagraph): WordBlock | null {
  const meaningful = paragraph.runs.filter(
    (run) => !(run.type === 'text' && run.text.replace(/[\s\u00a0]/g, '').length === 0),
  );
  if (meaningful.length === 0) return null;
  if (!meaningful.every((run) => run.type === 'break' && run.breakType === 'page')) return null;
  return { type: 'page-break', breakType: 'page' };
}
