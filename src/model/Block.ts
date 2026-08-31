import type { WordParagraph } from './Paragraph.js';
import type { WordTable } from './Table.js';
import type { ImagePlacement } from './Image.js';
import type { Borders, ParagraphFormatting } from './Style.js';

/** A block-level image (a paragraph whose only content is a picture). */
export interface WordImageBlock {
  type: 'image-block';
  imageId: string;
  placement: ImagePlacement;
  formatting: ParagraphFormatting;
}

export interface WordPageBreak {
  type: 'page-break';
  /** `page` for an explicit break, `section` for a Word section boundary. */
  breakType: 'page' | 'section' | 'column';
  /** Section properties when Word described them (`mso-*` on `div.WordSection`). */
  sectionName?: string;
}

/** A horizontal rule Word emitted (`<hr>` or a bottom-bordered empty paragraph). */
export interface WordHorizontalRule {
  type: 'horizontal-rule';
  borders?: Borders;
}

/**
 * Anything Word expressed that the engine deliberately refuses to guess at.
 *
 * The construct is *kept* — with its raw markup — rather than deleted, and a
 * matching diagnostic is emitted. The renderer emits an inert, clearly marked
 * placeholder for it.
 */
export interface WordUnsupportedObject {
  type: 'unsupported';
  /** Machine-readable kind, e.g. `ole-object`, `vml-shape`, `activex`. */
  objectType: string;
  /** Diagnostic code raised alongside this block. */
  code: string;
  /** Attributes / metadata scraped off the construct. */
  rawMetadata: Record<string, string>;
  /** The original markup, truncated to a safe length. */
  rawMarkup?: string;
  /** Any text Word rendered as a fallback, preserved verbatim. */
  fallbackText?: string;
}

/**
 * A `div`-level container Word used for structure (`div.WordSection1`,
 * a bordered container, a floating text box). Kept so nesting is not flattened.
 */
export interface WordContainer {
  type: 'container';
  role: 'section' | 'div' | 'blockquote' | 'text-box';
  blocks: WordBlock[];
  formatting: ParagraphFormatting;
  className?: string;
}

export type WordBlock =
  | WordParagraph
  | WordTable
  | WordImageBlock
  | WordPageBreak
  | WordHorizontalRule
  | WordUnsupportedObject
  | WordContainer;

export const isParagraph = (b: WordBlock): b is WordParagraph => b.type === 'paragraph';
export const isTable = (b: WordBlock): b is WordTable => b.type === 'table';
export const isContainer = (b: WordBlock): b is WordContainer => b.type === 'container';

/** Depth-first walk over every block, descending into tables and containers. */
export function* walkBlocks(blocks: WordBlock[]): Generator<WordBlock> {
  for (const block of blocks) {
    yield block;
    if (block.type === 'container') {
      yield* walkBlocks(block.blocks);
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          yield* walkBlocks(cell.blocks);
        }
      }
    }
  }
}
