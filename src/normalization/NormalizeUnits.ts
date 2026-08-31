import type { Length, ParagraphFormatting } from '../model/Style.js';
import type { ListItemInfo } from '../model/List.js';
import { lengthFromPx, roundTo } from '../word/WordLengthParser.js';

/**
 * Unit normalisation.
 *
 * All conversion has already happened in `parseWordLength`; what is left is
 * the indentation arithmetic that is specific to list items, and it is
 * genuinely fiddly, so it lives in exactly one place.
 *
 * Word expresses a list item's indentation as a *hanging indent*:
 *
 *     margin-left:  .5in    where the text sits
 *     text-indent: -.25in   how far left of that the marker starts
 *
 * So the marker begins at 0.25in and the text at 0.5in. CSS lists work
 * differently — the marker sits in the padding box — so the two numbers have
 * to be converted rather than copied, or every pasted list comes out with the
 * wrong indentation at every level.
 */

/** The indentation of a list item, expressed the way CSS needs it. */
export interface ListIndentation {
  /** Distance from the container's left edge to the marker, in px. */
  markerOffsetPx: number;
  /** Distance from the container's left edge to the text, in px. */
  textOffsetPx: number;
  /** The gap the marker occupies: `textOffset - markerOffset`. */
  hangingPx: number;
  /** True when the values came from Word rather than from a default. */
  explicit: boolean;
}

/** Word's default list indentation step: half an inch per level. */
export const DEFAULT_LEVEL_INDENT_PX = 48;

/**
 * Work out where a list item's marker and text belong.
 *
 * Prefers what the paragraph itself declared, falls back to the level
 * definition, and only then to Word's default half-inch-per-level ladder.
 */
export function computeListIndentation(
  item: ListItemInfo,
  formatting: ParagraphFormatting,
): ListIndentation {
  const marginLeft = item.marginLeft ?? formatting.marginLeft ?? item.levelDefinition?.marginLeft;
  const textIndent = item.textIndent ?? formatting.textIndent ?? item.levelDefinition?.textIndent;

  if (marginLeft || textIndent) {
    const textOffsetPx = marginLeft?.px ?? 0;
    // A negative text-indent is a hanging indent: the marker starts that far
    // to the left of the text.
    const markerOffsetPx = textOffsetPx + (textIndent?.px ?? 0);
    return {
      textOffsetPx: roundTo(textOffsetPx, 2),
      markerOffsetPx: roundTo(Math.max(0, markerOffsetPx), 2),
      hangingPx: roundTo(Math.max(0, textOffsetPx - Math.max(0, markerOffsetPx)), 2),
      explicit: true,
    };
  }

  // No declaration anywhere: Word's default ladder.
  const tabStop = item.levelDefinition?.tabStop?.px;
  const textOffsetPx = tabStop ?? DEFAULT_LEVEL_INDENT_PX * (item.level + 1);
  const markerOffsetPx = Math.max(0, textOffsetPx - DEFAULT_LEVEL_INDENT_PX / 2);
  return {
    textOffsetPx: roundTo(textOffsetPx, 2),
    markerOffsetPx: roundTo(markerOffsetPx, 2),
    hangingPx: roundTo(textOffsetPx - markerOffsetPx, 2),
    explicit: false,
  };
}

/**
 * The indentation a nested list should carry, relative to its parent list.
 *
 * Nested `<ol>`/`<ul>` elements inherit their parent's indentation, so the
 * child must only add the *difference* — copying the absolute value would
 * double the indent at every level.
 */
export function relativeIndent(child: ListIndentation, parent: ListIndentation | null): number {
  if (!parent) return child.markerOffsetPx;
  return roundTo(Math.max(0, child.markerOffsetPx - parent.markerOffsetPx), 2);
}

/** Round every length in a formatting object, to keep golden output stable. */
export function roundLength(length: Length | undefined): Length | undefined {
  if (!length) return undefined;
  return { ...length, px: roundTo(length.px, 2) };
}

/** Build a Length for a px value, used when the engine synthesises indentation. */
export function px(value: number): Length {
  return lengthFromPx(roundTo(value, 2));
}
