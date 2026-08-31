import type { Borders, Color, Length, Shading } from './Style.js';
import type { WordBlock } from './Block.js';

export type CellVerticalAlign = 'top' | 'middle' | 'bottom';

export interface WordTableCell {
  type: 'cell';
  /** Cell content is a full block list, so nested tables and lists survive. */
  blocks: WordBlock[];
  colSpan: number;
  rowSpan: number;
  width?: Length;
  /** `%` width when Word expressed it proportionally. */
  widthPercent?: number;
  borders?: Borders;
  shading?: Shading;
  backgroundColor?: Color;
  padding?: {
    top?: Length;
    right?: Length;
    bottom?: Length;
    left?: Length;
  };
  verticalAlign?: CellVerticalAlign;
  /** True for `<th>` or a cell Word marked as a header row cell. */
  header?: boolean;
  /** True when this slot is covered by a span from another cell. */
  covered?: boolean;
  /** Zero-based grid coordinates assigned by NormalizeTables. */
  gridColumn?: number;
  gridRow?: number;
}
