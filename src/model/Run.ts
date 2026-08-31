import type { RunFormatting } from './Style.js';

/**
 * The leaf content of a paragraph.
 *
 * A run is the smallest span of text that shares identical formatting. Runs are
 * never merged across a formatting boundary, so `Hello <b>world</b>` stays two
 * runs.
 */
export interface WordTextRun {
  type: 'text';
  text: string;
  formatting: RunFormatting;
  /** Word paragraph/character style names in effect, outermost first. */
  styleChain?: string[];
  /** Id of the hyperlink this run sits inside, if any. */
  hyperlinkId?: string;
  /** Bookmark names that open at this run. */
  bookmarks?: string[];
}

export interface WordLineBreakRun {
  type: 'break';
  /** `line` for `<br>`, `page` for an explicit page break inside a paragraph. */
  breakType: 'line' | 'page' | 'column';
  formatting: RunFormatting;
}

export interface WordTabRun {
  type: 'tab';
  formatting: RunFormatting;
}

/** An inline image sitting inside a paragraph. */
export interface WordImageRun {
  type: 'image';
  imageId: string;
  formatting: RunFormatting;
  hyperlinkId?: string;
}

/**
 * A Word field result the engine could not model semantically (page numbers,
 * cross references, TOC entries…). The visible text is preserved so nothing
 * disappears, and a diagnostic is raised.
 */
export interface WordFieldRun {
  type: 'field';
  /** Field instruction when recoverable, e.g. `PAGEREF _Toc123 \\h`. */
  instruction?: string;
  /** The rendered result Word produced. */
  text: string;
  formatting: RunFormatting;
}

/** A footnote/endnote reference mark. */
export interface WordNoteRun {
  type: 'note';
  noteType: 'footnote' | 'endnote';
  reference: string;
  text: string;
  formatting: RunFormatting;
}

export type WordRun =
  | WordTextRun
  | WordLineBreakRun
  | WordTabRun
  | WordImageRun
  | WordFieldRun
  | WordNoteRun;

export const emptyRunFormatting = (): RunFormatting => ({});
