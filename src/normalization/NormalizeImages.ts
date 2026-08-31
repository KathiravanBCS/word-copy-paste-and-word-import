import type { WordDocument } from '../model/Document.js';
import type { WordImage } from '../model/Image.js';
import type { DiagnosticCollector } from '../diagnostics/DiagnosticCollector.js';
import { DiagnosticCode } from '../diagnostics/UnsupportedFeature.js';
import { roundTo } from '../word/WordLengthParser.js';
import { px } from './NormalizeUnits.js';

/**
 * Image normalisation.
 *
 * Two things need settling after parsing:
 *
 *   - **Aspect ratio.** Word frequently states only one dimension when the
 *     picture is scaled. Emitting one dimension makes the browser use the
 *     image's natural size for the other, which is right for a resolved image
 *     and impossible for an unresolved one — so a placeholder needs both.
 *   - **Floating placement.** Word floats pictures using `mso-position-*`
 *     against the page. HTML can float against the text column and nothing
 *     else, so the placement is downgraded to a float with a diagnostic
 *     saying exactly what was lost.
 */

export interface NormalizeImagesResult {
  images: number;
  resolved: number;
  unresolved: number;
  floating: number;
}

/** Size used for a placeholder when Word gave no dimensions at all. */
const PLACEHOLDER_WIDTH_PX = 240;
const PLACEHOLDER_HEIGHT_PX = 135;

export function normalizeImages(
  document: WordDocument,
  diagnostics: DiagnosticCollector,
): NormalizeImagesResult {
  const result: NormalizeImagesResult = { images: 0, resolved: 0, unresolved: 0, floating: 0 };

  for (const image of Object.values(document.images)) {
    result.images++;
    if (image.resolution === 'unresolved') result.unresolved++;
    else result.resolved++;

    if (image.placement === 'floating') {
      result.floating++;
      diagnostics.warn(
        DiagnosticCode.WORD_FLOATING_IMAGE_APPROXIMATED,
        'A floating picture was converted to a CSS float. Word anchors floating pictures to a paragraph and positions them against the page or margin; HTML can only float them within the text column, so the exact position is not reproduced.',
        { details: { imageId: image.id } },
      );
    }

    ensureDimensions(image);
  }
  return result;
}

function ensureDimensions(image: WordImage): void {
  if (image.width && image.height) return;

  if (image.resolution === 'unresolved') {
    // A placeholder has to be laid out, and there is nothing to measure.
    if (!image.width) {
      image.width = px(image.height ? roundTo(image.height.px * (16 / 9), 2) : PLACEHOLDER_WIDTH_PX);
    }
    if (!image.height) {
      image.height = px(image.width ? roundTo(image.width.px * (9 / 16), 2) : PLACEHOLDER_HEIGHT_PX);
    }
    return;
  }

  // A resolved image with one dimension: leave the other to the browser, which
  // knows the natural size. Inventing a value here would distort the picture.
}
