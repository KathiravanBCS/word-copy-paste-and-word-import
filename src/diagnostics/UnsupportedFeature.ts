/**
 * Diagnostic codes.
 *
 * Rule: nothing in a Word payload is ever silently dropped. If the engine
 * cannot represent a construct exactly, it emits one of these, keeps whatever
 * raw material it has, and the FidelityReport surfaces it.
 */
export const DiagnosticCode = {
  // --- images -------------------------------------------------------------
  WORD_UNRESOLVED_IMAGE: 'WORD_UNRESOLVED_IMAGE',
  WORD_LOCAL_FILE_IMAGE: 'WORD_LOCAL_FILE_IMAGE',
  WORD_CID_IMAGE: 'WORD_CID_IMAGE',
  WORD_IMAGE_CROP_APPROXIMATED: 'WORD_IMAGE_CROP_APPROXIMATED',
  WORD_FLOATING_IMAGE_APPROXIMATED: 'WORD_FLOATING_IMAGE_APPROXIMATED',

  // --- drawing / embedded objects ----------------------------------------
  WORD_VML_OBJECT: 'WORD_VML_OBJECT',
  WORD_VML_SHAPE_APPROXIMATED: 'WORD_VML_SHAPE_APPROXIMATED',
  WORD_OLE_OBJECT: 'WORD_OLE_OBJECT',
  WORD_ACTIVEX_OBJECT: 'WORD_ACTIVEX_OBJECT',
  WORD_UNSUPPORTED_SHAPE: 'WORD_UNSUPPORTED_SHAPE',
  WORD_TEXT_BOX_APPROXIMATED: 'WORD_TEXT_BOX_APPROXIMATED',
  WORD_EQUATION_UNSUPPORTED: 'WORD_EQUATION_UNSUPPORTED',
  WORD_SMARTART_UNSUPPORTED: 'WORD_SMARTART_UNSUPPORTED',
  WORD_CHART_UNSUPPORTED: 'WORD_CHART_UNSUPPORTED',

  // --- fields / notes -----------------------------------------------------
  WORD_UNSUPPORTED_FIELD: 'WORD_UNSUPPORTED_FIELD',
  WORD_FOOTNOTE_APPROXIMATED: 'WORD_FOOTNOTE_APPROXIMATED',
  WORD_FORM_FIELD_UNSUPPORTED: 'WORD_FORM_FIELD_UNSUPPORTED',
  WORD_CONTENT_CONTROL_APPROXIMATED: 'WORD_CONTENT_CONTROL_APPROXIMATED',

  // --- lists --------------------------------------------------------------
  WORD_LIST_DEFINITION_MISSING: 'WORD_LIST_DEFINITION_MISSING',
  WORD_LIST_LEVEL_MISSING: 'WORD_LIST_LEVEL_MISSING',
  WORD_LIST_MARKER_HEURISTIC: 'WORD_LIST_MARKER_HEURISTIC',
  WORD_LIST_NUMBER_FORMAT_APPROXIMATED: 'WORD_LIST_NUMBER_FORMAT_APPROXIMATED',
  WORD_SYMBOL_FONT_MAPPED: 'WORD_SYMBOL_FONT_MAPPED',
  WORD_SYMBOL_FONT_UNMAPPED: 'WORD_SYMBOL_FONT_UNMAPPED',

  // --- tables -------------------------------------------------------------
  WORD_TABLE_GRID_REPAIRED: 'WORD_TABLE_GRID_REPAIRED',
  WORD_NESTED_TABLE: 'WORD_NESTED_TABLE',

  // --- structural ---------------------------------------------------------
  WORD_SECTION_BREAK_APPROXIMATED: 'WORD_SECTION_BREAK_APPROXIMATED',
  WORD_UNKNOWN_ELEMENT: 'WORD_UNKNOWN_ELEMENT',
  WORD_NAMESPACE_ELEMENT_DROPPED: 'WORD_NAMESPACE_ELEMENT_DROPPED',
  WORD_CSS_PARSE_WARNING: 'WORD_CSS_PARSE_WARNING',
  WORD_FRAGMENT_BOUNDARY_MISSING: 'WORD_FRAGMENT_BOUNDARY_MISSING',
  WORD_REVISION_MARK_FLATTENED: 'WORD_REVISION_MARK_FLATTENED',

  // --- security / limits --------------------------------------------------
  SECURITY_SCRIPT_REMOVED: 'SECURITY_SCRIPT_REMOVED',
  SECURITY_EVENT_HANDLER_REMOVED: 'SECURITY_EVENT_HANDLER_REMOVED',
  SECURITY_URL_BLOCKED: 'SECURITY_URL_BLOCKED',
  SECURITY_EXTERNAL_RESOURCE: 'SECURITY_EXTERNAL_RESOURCE',
  LIMIT_DOCUMENT_TRUNCATED: 'LIMIT_DOCUMENT_TRUNCATED',
  LIMIT_DEPTH_EXCEEDED: 'LIMIT_DEPTH_EXCEEDED',
  LIMIT_NODE_BUDGET_EXCEEDED: 'LIMIT_NODE_BUDGET_EXCEEDED',

  // --- detection ----------------------------------------------------------
  NOT_WORD_CONTENT: 'NOT_WORD_CONTENT',
} as const;

export type DiagnosticCodeValue = (typeof DiagnosticCode)[keyof typeof DiagnosticCode];

export type DiagnosticSeverity = 'info' | 'warning' | 'error';

/**
 * How faithfully a construct made it across, per the fidelity classification
 * in docs/FIDELITY.md.
 */
export type FidelityClass = 'EXACT' | 'EQUIVALENT' | 'APPROXIMATED' | 'UNSUPPORTED';

export interface DiagnosticLocation {
  /** A CSS-ish path to the offending node in the parsed clipboard DOM. */
  path?: string;
  /** Index of the top-level block the problem belongs to. */
  blockIndex?: number;
  /** Tag name of the offending element. */
  tagName?: string;
  /** A short excerpt of the raw markup, truncated. */
  excerpt?: string;
}

export interface WordDiagnostic {
  code: DiagnosticCodeValue | string;
  severity: DiagnosticSeverity;
  fidelity: FidelityClass;
  message: string;
  location?: DiagnosticLocation;
  /** Structured extra facts, e.g. `{ font: 'Wingdings', glyph: '\\uF0A7' }`. */
  details?: Record<string, string | number | boolean>;
  /** How many times this exact code+message pair occurred. */
  count?: number;
}

/** Human-readable one-liner explaining what a code means. */
export const DiagnosticExplanations: Record<string, string> = {
  WORD_UNRESOLVED_IMAGE:
    'Word referenced an image whose bytes were not present in the clipboard. A placeholder was emitted instead of a broken image.',
  WORD_LOCAL_FILE_IMAGE:
    'Word referenced an image by local file path (file:///). Browsers cannot load these; the reference was preserved as metadata.',
  WORD_CID_IMAGE:
    'Word referenced an image by content id (cid:). The bytes live in the MIME part, not the HTML flavour.',
  WORD_VML_OBJECT:
    'A VML drawing was found. Where it wrapped a picture the picture was extracted; the VML markup itself is preserved as metadata and not rendered.',
  WORD_OLE_OBJECT:
    'An embedded OLE object was detected. OLE has no HTML equivalent; its metadata was preserved and a placeholder emitted.',
  WORD_LIST_DEFINITION_MISSING:
    'A paragraph referenced an mso-list definition that was not present in the clipboard stylesheet. The rendered marker Word emitted was used instead.',
  WORD_LIST_MARKER_HEURISTIC:
    'No mso-list marker span was present, so the marker was recovered heuristically from leading text. Verify against the source document.',
  WORD_SYMBOL_FONT_MAPPED:
    'A bullet glyph was encoded in a symbol font (Symbol/Wingdings). It was mapped to its Unicode equivalent; the raw byte and font are preserved.',
  WORD_SYMBOL_FONT_UNMAPPED:
    'A bullet glyph in a symbol font had no known Unicode equivalent. The raw glyph and its font are preserved and rendered with that font.',
  WORD_TABLE_GRID_REPAIRED:
    'Row lengths disagreed with the resolved column grid; missing cells were filled so the table stays rectangular.',
  WORD_SECTION_BREAK_APPROXIMATED:
    'A Word section break was represented as a page break; page setup differences between sections are not expressible in HTML.',
  SECURITY_SCRIPT_REMOVED:
    'Script content in the clipboard payload was removed. Clipboard HTML is untrusted input and is never executed.',
  SECURITY_URL_BLOCKED:
    'A URL using a non-permitted scheme (javascript:, data: on a link, vbscript:) was blocked.',
  LIMIT_NODE_BUDGET_EXCEEDED:
    'The pasted document exceeded the configured node budget and was truncated to keep parsing bounded.',
};
