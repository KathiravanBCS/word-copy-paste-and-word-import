import type { Length, Shading } from './Style.js';
import type { WordTableCell } from './TableCell.js';

export interface WordTableRow {
  type: 'row';
  cells: WordTableCell[];
  height?: Length;
  /** `mso-height-rule` / `height` semantics. */
  heightRule?: 'auto' | 'atLeast' | 'exact';
  /** True when Word marked the row as a repeating header row. */
  header?: boolean;
  /** `mso-row-margin-*` / cell spacing inherited from the table. */
  shading?: Shading;
  /** Which section of the table this row belongs to. */
  section: 'head' | 'body' | 'foot';
}
