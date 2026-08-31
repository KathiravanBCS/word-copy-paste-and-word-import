import type { Borders, Length, ParagraphAlignment, Shading } from './Style.js';
import type { WordTableRow } from './TableRow.js';

export interface WordTableColumn {
  width?: Length;
  widthPercent?: number;
}

export interface WordTable {
  type: 'table';
  rows: WordTableRow[];
  /** Column definitions derived from the widest row / Word's own `<col>`s. */
  columns: WordTableColumn[];
  width?: Length;
  widthPercent?: number;
  alignment?: ParagraphAlignment;
  /** Table indentation from the left margin (`mso-table-lspace`). */
  indent?: Length;
  borders?: Borders;
  shading?: Shading;
  /** Default cell padding declared on the table (`mso-padding-alt`). */
  cellPadding?: {
    top?: Length;
    right?: Length;
    bottom?: Length;
    left?: Length;
  };
  /** `border-collapse` as Word declared it. */
  borderCollapse?: 'collapse' | 'separate';
  cellSpacing?: Length;
  /** Word table style name, e.g. `TableGrid`. */
  styleName?: string;
  /** Nesting depth; 0 for a top-level table. */
  depth: number;
  /** Number of grid columns after span resolution. */
  gridColumnCount: number;
  caption?: string;
}
