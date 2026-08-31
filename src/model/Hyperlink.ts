/**
 * Hyperlinks are stored out-of-line in `WordDocument.hyperlinks` and referenced
 * by id from runs, so that run formatting inside a link is never flattened.
 */
export interface WordHyperlink {
  id: string;
  /** Resolved, security-screened href. Empty when the target was rejected. */
  href: string;
  /** The href exactly as Word wrote it (may be `file:///…` or a local anchor). */
  rawHref: string;
  /** `#name` internal jumps carry the bookmark name here. */
  anchor?: string;
  title?: string;
  target?: string;
  /** True when the href was dropped for security reasons (`javascript:` …). */
  blocked?: boolean;
}

export interface WordBookmark {
  id: string;
  name: string;
  /** True for Word's own TOC bookmarks (`_Toc…`, `_Ref…`, `_GoBack`). */
  internal: boolean;
}
