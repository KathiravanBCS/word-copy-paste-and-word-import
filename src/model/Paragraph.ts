import type { ParagraphFormatting } from './Style.js';
import type { WordRun } from './Run.js';
import type { ListItemInfo } from './List.js';

/**
 * A paragraph — the workhorse block.
 *
 * `headingLevel` is only set when Word itself said so (an `MsoHeadingN` style,
 * an `mso-outline-level` declaration, or a real `<h1>`..`<h6>` element). It is
 * never guessed from font size.
 */
export interface WordParagraph {
  type: 'paragraph';
  runs: WordRun[];
  /** Word style name, e.g. `Heading 1`, `List Paragraph`, `Normal`. */
  styleName?: string;
  /** Normalised style id used to look the style up in the stylesheet. */
  styleId?: string;
  /** Class names Word put on the element, preserved for diagnostics. */
  classNames?: string[];
  formatting: ParagraphFormatting;
  /** Set when this paragraph is a list item. */
  listItem?: ListItemInfo;
  /** 1..6 when this paragraph is a heading. */
  headingLevel?: number;
  /** True when the paragraph has no visible content (Word's spacer paragraphs). */
  empty?: boolean;
  /** Bookmarks anchored at the start of this paragraph. */
  bookmarks?: string[];
}
