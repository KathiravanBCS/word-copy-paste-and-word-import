/**
 * Individual Word signals and their weights.
 *
 * Detection is deliberately additive rather than "does it contain mso-". A
 * single weak signal (say, a stray `class="MsoNormal"` that survived a trip
 * through another editor) must not be enough on its own; the engine wants
 * corroboration before it applies Word-specific parsing rules to a payload.
 */

export type SignalStrength = 'decisive' | 'strong' | 'moderate' | 'weak';

export interface WordSignalDefinition {
  id: string;
  description: string;
  strength: SignalStrength;
  /** Test run against the raw HTML string. */
  test: (html: string) => boolean;
}

export interface DetectedSignal {
  id: string;
  description: string;
  strength: SignalStrength;
  weight: number;
}

export const SIGNAL_WEIGHTS: Record<SignalStrength, number> = {
  decisive: 1.0,
  strong: 0.45,
  moderate: 0.25,
  weak: 0.1,
};

/**
 * Signals are matched against the raw HTML text rather than a parsed DOM,
 * because several of them (conditional comments, the `<style>` comment
 * wrapper, `xmlns:` prefixes) are altered or normalised away by HTML parsing.
 */
export const WORD_SIGNALS: WordSignalDefinition[] = [
  {
    id: 'meta-progid-word',
    description: '<meta name=ProgId content=Word.Document>',
    strength: 'decisive',
    test: (h) => /<meta[^>]+name=["']?ProgId["']?[^>]*content=["']?Word\.(Document|Sheet)/i.test(h),
  },
  {
    id: 'meta-generator-word',
    description: '<meta name=Generator content="Microsoft Word ...">',
    strength: 'decisive',
    test: (h) => /<meta[^>]+name=["']?Generator["']?[^>]*content=["']?Microsoft\s+Word/i.test(h),
  },
  {
    id: 'ns-office',
    description: 'urn:schemas-microsoft-com:office namespace',
    strength: 'strong',
    test: (h) => /urn:schemas-microsoft-com:office/i.test(h),
  },
  {
    id: 'ns-xmlns-o',
    description: 'xmlns:o (Office namespace prefix)',
    strength: 'moderate',
    test: (h) => /xmlns:o\s*=/i.test(h),
  },
  {
    id: 'ns-xmlns-w',
    description: 'xmlns:w (Word namespace prefix)',
    strength: 'strong',
    test: (h) => /xmlns:w\s*=/i.test(h),
  },
  {
    id: 'ns-xmlns-m',
    description: 'xmlns:m (Office Math namespace prefix)',
    strength: 'moderate',
    test: (h) => /xmlns:m\s*=/i.test(h),
  },
  {
    id: 'ns-xmlns-v',
    description: 'xmlns:v (VML namespace prefix)',
    strength: 'moderate',
    test: (h) => /xmlns:v\s*=/i.test(h),
  },
  {
    id: 'class-mso',
    description: 'class="Mso..." (Word style classes)',
    strength: 'moderate',
    test: (h) => /class=["']?Mso[A-Za-z]/i.test(h),
  },
  {
    id: 'css-mso-declaration',
    description: 'mso-* CSS declarations',
    strength: 'strong',
    test: (h) => /\bmso-[a-z-]+\s*:/i.test(h),
  },
  {
    id: 'css-mso-list',
    description: 'mso-list declarations (Word numbering)',
    strength: 'strong',
    test: (h) => /\bmso-list\s*:/i.test(h),
  },
  {
    id: 'css-at-list',
    description: '@list list definitions',
    strength: 'decisive',
    test: (h) => /@list\s+l\d+/i.test(h),
  },
  {
    id: 'css-font-face-mso',
    description: '@font-face with mso-font-* declarations',
    strength: 'strong',
    test: (h) => /@font-face[^}]*mso-(font-charset|generic-font-family|font-alt)/i.test(h),
  },
  {
    id: 'conditional-comment-mso',
    description: 'Word conditional comment (<!--[if ... mso ...]>)',
    strength: 'strong',
    test: (h) => /<!--\[if\s+(gte\s+)?mso\s|<!--\[if\s+!?(supportLists|vml|mso)/i.test(h),
  },
  {
    id: 'conditional-comment-supportlists',
    description: 'Downlevel-revealed <![if !supportLists]> list marker block',
    strength: 'decisive',
    test: (h) => /<!\[if\s+!supportLists\]>/i.test(h),
  },
  {
    id: 'conditional-comment-vml',
    description: 'VML conditional comment (<!--[if gte vml 1]>)',
    strength: 'strong',
    test: (h) => /<!--\[if\s+(gte\s+)?vml\s/i.test(h),
  },
  {
    id: 'element-o-p',
    description: '<o:p> Office paragraph marker element',
    strength: 'strong',
    test: (h) => /<o:p[\s>/]/i.test(h),
  },
  {
    id: 'element-vml-shape',
    description: '<v:shape> / <v:imagedata> VML markup',
    strength: 'moderate',
    test: (h) => /<v:(shape|imagedata|shapetype|rect|oval|group|line)[\s>/]/i.test(h),
  },
  {
    id: 'element-word-document-xml',
    description: '<w:WordDocument> settings block',
    strength: 'decisive',
    test: (h) => /<w:WordDocument[\s>]/i.test(h),
  },
  {
    id: 'element-office-document-properties',
    description: '<o:DocumentProperties> metadata block',
    strength: 'strong',
    test: (h) => /<o:DocumentProperties[\s>]/i.test(h),
  },
  {
    id: 'div-word-section',
    description: 'div.WordSectionN section container',
    strength: 'strong',
    test: (h) => /class=["']?WordSection\d/i.test(h),
  },
  {
    id: 'style-mso-style-name',
    description: 'mso-style-name declarations in the stylesheet',
    strength: 'moderate',
    test: (h) => /mso-style-(name|parent|link|type)\s*:/i.test(h),
  },
  {
    id: 'fragment-markers',
    description: '<!--StartFragment--> / <!--EndFragment--> CF_HTML markers',
    strength: 'weak',
    test: (h) => /<!--\s*StartFragment\s*-->/i.test(h),
  },
  {
    id: 'clip-image-reference',
    description: 'clip_imageNNN image reference',
    strength: 'moderate',
    test: (h) => /clip_image\d+/i.test(h),
  },
  {
    id: 'mso-spacerun',
    description: "span style='mso-spacerun:yes' whitespace runs",
    strength: 'moderate',
    test: (h) => /mso-spacerun\s*:/i.test(h),
  },
  {
    id: 'ole-object',
    description: '<o:OLEObject> embedded object',
    strength: 'moderate',
    test: (h) => /<o:OLEObject[\s>]/i.test(h),
  },
  {
    id: 'panose-1',
    description: 'panose-1 font metrics (Office @font-face)',
    strength: 'weak',
    test: (h) => /panose-1\s*:/i.test(h),
  },
];

/** Run every signal against the payload. */
export function detectSignals(html: string): DetectedSignal[] {
  const found: DetectedSignal[] = [];
  for (const signal of WORD_SIGNALS) {
    let matched = false;
    try {
      matched = signal.test(html);
    } catch {
      matched = false;
    }
    if (matched) {
      found.push({
        id: signal.id,
        description: signal.description,
        strength: signal.strength,
        weight: SIGNAL_WEIGHTS[signal.strength],
      });
    }
  }
  return found;
}
