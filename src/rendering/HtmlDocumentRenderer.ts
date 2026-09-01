import type { WordDocument } from '../model/Document.js';
import { escapeHtmlAttribute, escapeHtmlText } from '../util/dom.js';
import { renderWordDocument, type RenderOptions, type RenderResult } from './HtmlRenderer.js';
import { buildFidelityReport, type FidelityReport } from '../diagnostics/FidelityReport.js';
import { parseBoxShorthand, parseWordLength, roundTo } from '../word/WordLengthParser.js';

/**
 * Standalone HTML document output.
 *
 * This is what a "download as HTML" button produces: a complete, self-
 * contained `.html` file that opens in any browser and looks like what was
 * copied out of Word, with no external stylesheet, no script, and no network
 * dependency beyond any images that were genuinely external in the first
 * place.
 *
 * The generated stylesheet — including the `@counter-style` rules compiled
 * from Word's own numbering definitions — goes in the head, which is what
 * makes the native list markers survive the trip to a file on disk.
 */

export interface DocumentRenderOptions extends RenderOptions {
  /** `<title>` of the produced document. */
  title?: string;
  /** Page language. Default `en`. */
  lang?: string;
  /**
   * Include a page shell — a white page on a grey background at the width Word
   * used. Default true; set false for a bare document with no chrome.
   */
  pageShell?: boolean;
  /**
   * Content width for the page shell. Defaults to the payload's own `@page`
   * size and margins when Word declared one (its page width minus its own
   * left/right margins), falling back to `6.5in` — Letter minus 1in margins —
   * only when it did not. Set explicitly to override either.
   */
  contentWidth?: string;
  /**
   * Page padding for the page shell, as a CSS `padding` value. Defaults to
   * the payload's own `@page` margins, falling back to `0.6in 0.75in`.
   */
  pagePadding?: string;
  /**
   * Append a fidelity summary as an HTML comment at the end of the document.
   * Default true — it costs nothing and makes a downloaded file
   * self-documenting about what did and did not survive.
   */
  includeFidelityComment?: boolean;
  /** Append a visible appendix listing diagnostics. Default false. */
  includeDiagnosticsAppendix?: boolean;
}

export interface DocumentRenderResult extends RenderResult {
  /** The complete `<!doctype html>` document. */
  document: string;
  report: FidelityReport;
}

export function renderStandaloneHtml(
  document: WordDocument,
  options: DocumentRenderOptions = {},
): DocumentRenderResult {
  const rendered = renderWordDocument(document, { ...options, cssMode: 'separate' });
  const report = buildFidelityReport(document);

  const title =
    options.title ??
    document.metadata.documentProperties['Title'] ??
    'Pasted from Microsoft Word';
  const lang = options.lang ?? 'en';
  const prefix = options.classPrefix ?? 'wce';
  const geometry = resolvePageGeometry(document);

  const parts: string[] = [
    '<!doctype html>',
    `<html lang="${escapeHtmlAttribute(lang)}">`,
    '<head>',
    '<meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    `<title>${escapeHtmlText(title)}</title>`,
    '<style>',
    documentShellCss(prefix, options, geometry),
    rendered.css,
    '</style>',
    '</head>',
    '<body>',
    options.pageShell === false
      ? rendered.html
      : `<main class="${prefix}-page">${rendered.html}</main>`,
  ];

  if (options.includeDiagnosticsAppendix) {
    parts.push(renderDiagnosticsAppendix(document, report, prefix));
  }
  parts.push('</body>', '</html>');

  if (options.includeFidelityComment !== false) {
    parts.push(renderFidelityComment(report));
  }

  return { ...rendered, document: parts.join('\n'), report };
}

/**
 * The page geometry the payload itself declared.
 *
 * Word writes one `@page` rule per section (`@page WordSection1 { size:8.5in
 * 11.0in; margin:1.0in 1.0in 1.0in 1.0in; }`), which the stylesheet parser
 * already captures as raw declarations. This is the one place they are read
 * back out, so the downloaded file's page width and margins come from the
 * document that was actually pasted rather than from a Letter-with-1in-margins
 * guess that happens to be right for a great many documents and wrong for the
 * rest.
 *
 * Only the first section's page setup is used. A document with a genuine
 * section-by-section page-size change is already flagged with
 * `WORD_SECTION_BREAK_APPROXIMATED`, and picking one geometry for a single
 * scrolling page is the same approximation whichever section it comes from.
 */
interface PageGeometry {
  contentWidth: string;
  margin: string;
}

function resolvePageGeometry(document: WordDocument): PageGeometry | null {
  const sectionName = document.metadata.sections[0];
  const declarations = sectionName
    ? document.styles.pages[sectionName]
    : Object.values(document.styles.pages)[0];
  if (!declarations) return null;

  const size = declarations['size'];
  const pageWidth = size ? parseWordLength(size.trim().split(/\s+/)[0], { defaultUnit: 'in' }) : undefined;

  const marginRaw = declarations['margin'];
  const margins = marginRaw ? parseBoxShorthand(marginRaw, { defaultUnit: 'in' }) : undefined;
  const marginTop = margins?.top ?? parseWordLength(declarations['margin-top'], { defaultUnit: 'in' });
  const marginRight = margins?.right ?? parseWordLength(declarations['margin-right'], { defaultUnit: 'in' });
  const marginBottom = margins?.bottom ?? parseWordLength(declarations['margin-bottom'], { defaultUnit: 'in' });
  const marginLeft = margins?.left ?? parseWordLength(declarations['margin-left'], { defaultUnit: 'in' });

  if (!pageWidth && !marginTop && !marginRight && !marginBottom && !marginLeft) return null;

  const geometry: PageGeometry = { margin: '0.6in 0.75in', contentWidth: '6.5in' };
  if (marginTop || marginRight || marginBottom || marginLeft) {
    geometry.margin = [marginTop, marginRight, marginBottom, marginLeft]
      .map((m) => (m ? `${m.px}px` : '0'))
      .join(' ');
  }
  if (pageWidth) {
    const contentWidthPx = pageWidth.px - (marginLeft?.px ?? 0) - (marginRight?.px ?? 0);
    if (contentWidthPx > 0) geometry.contentWidth = `${roundTo(contentWidthPx, 2)}px`;
  }
  return geometry;
}

/**
 * Page chrome plus a typographic baseline.
 *
 * Word content arrives with its own fonts and sizes on almost everything, so
 * this only has to supply what Word leaves to the application: the page box,
 * a default font for anything unstyled, and print rules so the downloaded
 * file prints the way it looked.
 */
function documentShellCss(
  prefix: string,
  options: DocumentRenderOptions,
  geometry: PageGeometry | null,
): string {
  const width = options.contentWidth ?? geometry?.contentWidth ?? '6.5in';
  const pagePadding = options.pagePadding ?? geometry?.margin ?? '0.6in 0.75in';
  return `
:root { color-scheme: light dark; }
* { box-sizing: border-box; }
body {
  margin: 0;
  padding: 24px 16px;
  background: #f4f4f5;
  color: #111;
  font-family: Calibri, "Segoe UI", system-ui, -apple-system, sans-serif;
  font-size: 11pt;
  line-height: 1.35;
}
.${prefix}-page {
  max-width: ${width};
  margin: 0 auto;
  padding: ${pagePadding};
  background: #fff;
  border-radius: 3px;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.14), 0 6px 18px rgba(0, 0, 0, 0.06);
}
.${prefix}-document > *:first-child { margin-top: 0; }
.${prefix}-document img { max-width: 100%; height: auto; }
.${prefix}-document table { max-width: 100%; }
.${prefix}-document a { color: #0563c1; }
h1, h2, h3, h4, h5, h6 { font-weight: 600; }
@media (prefers-color-scheme: dark) {
  body { background: #18181b; color: #e8e8ea; }
  .${prefix}-page { background: #232327; box-shadow: none; }
  .${prefix}-document a { color: #6aa9f0; }
}
@media print {
  body { background: #fff; padding: 0; }
  .${prefix}-page { max-width: none; margin: 0; padding: 0; box-shadow: none; border-radius: 0; }
  .${prefix}-page-break { break-before: page; page-break-before: always; }
}
`.trim();
}

/**
 * A machine-readable summary of what survived, as a trailing comment.
 *
 * Deliberately a comment rather than visible content: the file is the pasted
 * document, not a report about it. But anyone who opens the source — or greps
 * a directory of exports — can see exactly what was approximated.
 */
function renderFidelityComment(report: FidelityReport): string {
  const summary = {
    engine: 'word-clipboard-engine',
    wordDetected: report.wordDetected,
    confidence: report.detectionConfidence,
    paragraphs: report.paragraphs,
    runs: report.runs,
    lists: report.lists,
    listItems: report.listItems,
    maxListDepth: report.maxListDepth + 1,
    tables: report.tables,
    images: report.images,
    unresolvedImages: report.unresolvedImages,
    hyperlinks: report.hyperlinks,
    fidelity: report.fidelityBreakdown,
    approximated: report.approximatedFeatures,
    unsupported: report.unsupportedFeatures,
  };
  // `--` cannot appear inside an HTML comment.
  const json = JSON.stringify(summary, null, 2).replace(/--/g, '-‑');
  return `<!-- word-clipboard-engine fidelity report\n${json}\n-->`;
}

function renderDiagnosticsAppendix(
  document: WordDocument,
  report: FidelityReport,
  prefix: string,
): string {
  if (document.diagnostics.length === 0) return '';
  const rows = document.diagnostics
    .map((diagnostic) => {
      const count = (diagnostic.count ?? 1) > 1 ? ` (x${diagnostic.count})` : '';
      return (
        `<li><code>${escapeHtmlText(diagnostic.code)}</code> ` +
        `<span class="${prefix}-diagnostic-fidelity">${escapeHtmlText(diagnostic.fidelity)}</span>${count} — ` +
        `${escapeHtmlText(diagnostic.message)}</li>`
      );
    })
    .join('');
  return `
<aside class="${prefix}-diagnostics">
  <h2>Conversion notes</h2>
  <p>${report.fidelityBreakdown.EXACT} exact, ${report.fidelityBreakdown.EQUIVALENT} equivalent,
     ${report.fidelityBreakdown.APPROXIMATED} approximated, ${report.fidelityBreakdown.UNSUPPORTED} unsupported.</p>
  <ul>${rows}</ul>
</aside>`;
}

/**
 * A `Blob` of the standalone document, ready for a download link.
 *
 * Kept here rather than in the demo so any consumer gets the same file, with
 * the same BOM-free UTF-8 encoding Word content needs to survive being
 * reopened.
 */
export function toHtmlBlob(html: string): Blob {
  return new Blob([html], { type: 'text/html;charset=utf-8' });
}

/** A filename-safe slug derived from the document title. */
export function suggestFileName(document: WordDocument, fallback = 'word-paste'): string {
  const title = document.metadata.documentProperties['Title'] ?? '';
  const source = title || firstHeadingText(document) || fallback;
  const slug = source
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);
  return `${slug || fallback}.html`;
}

function firstHeadingText(document: WordDocument): string {
  for (const block of document.blocks) {
    if (block.type === 'paragraph' && block.headingLevel) {
      const text = block.runs
        .map((run) => (run.type === 'text' ? run.text : ''))
        .join('')
        .trim();
      if (text) return text;
    }
  }
  return '';
}
