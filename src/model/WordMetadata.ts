/** Document-level facts recovered from the clipboard envelope. */
export interface WordDocumentMetadata {
  /** `<meta name=Generator>` value, e.g. `Microsoft Word 15 (filtered medium)`. */
  generator?: string;
  /** `<meta name=ProgId>` value, e.g. `Word.Document`. */
  progId?: string;
  /** Word version parsed out of the generator string, when present. */
  wordVersion?: number;
  /** Declared charset. */
  charset?: string;
  /** XML namespaces declared on `<html>`, keyed by prefix. */
  namespaces: Record<string, string>;
  /** `<o:DocumentProperties>` values from the office conditional comment. */
  documentProperties: Record<string, string>;
  /** `<w:WordDocument>` settings, e.g. `View`, `Zoom`, `TrackMoves`. */
  wordSettings: Record<string, string>;
  /** Section names seen (`WordSection1`, …). */
  sections: string[];
  /** True when a `<!--StartFragment-->` boundary was found and honoured. */
  fragmentBoundaryFound: boolean;
  /** Byte length of the raw HTML payload. */
  rawHtmlLength: number;
  /** The source application as detected, e.g. `word-desktop`, `word-online`. */
  sourceApplication?: string;
}
