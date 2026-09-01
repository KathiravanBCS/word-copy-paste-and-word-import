import {
  captureFromNavigatorClipboard,
  clipboardPayloadFromHtml,
  detectWordHtml,
  formatFidelityReport,
  getFidelityReport,
  installPasteCapture,
  parseWordClipboard,
  renderStandaloneHtml,
  renderWordDocument,
  suggestFileName,
  type ClipboardPayload,
  type ListMarkerMode,
  type WordDocument,
} from '../../index.js';
import type { WordBlock } from '../../model/Block.js';
import type { WordDiagnostic } from '../../diagnostics/UnsupportedFeature.js';
import { DiagnosticExplanations } from '../../diagnostics/UnsupportedFeature.js';

/**
 * The clipboard lab.
 *
 * This is the engine's development environment, and it is deliberately built
 * around one idea: **show the payload, the model and the output side by side,
 * and never hide the difference between them.**
 *
 * When a paste comes out wrong, the question is always "which stage got it
 * wrong?" — the parser reading the payload, the model representing it, or the
 * renderer emitting it. A preview alone cannot answer that. Three panes can.
 *
 * Everything here reads the engine's public API. The lab has no privileged
 * access and no parsing logic of its own, so what it shows is what a consumer
 * would get.
 */

/* -------------------------------------------------------------------------
 * Fixtures, bundled at build time
 * ---------------------------------------------------------------------- */

const fixtureModules = import.meta.glob('../../fixtures/word/**/input.html', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>;

const fixtures = Object.entries(fixtureModules)
  .map(([path, html]) => ({
    id: path.replace(/^.*fixtures\/word\//, '').replace(/\/input\.html$/, ''),
    html,
  }))
  .sort((a, b) => a.id.localeCompare(b.id));

/* -------------------------------------------------------------------------
 * State
 * ---------------------------------------------------------------------- */

interface LabState {
  payload: ClipboardPayload | null;
  document: WordDocument | null;
  html: string;
  css: string;
  markerMode: ListMarkerMode;
  rawTab: string;
  modelTab: string;
  outputTab: string;
  sourceLabel: string;
}

const state: LabState = {
  payload: null,
  document: null,
  html: '',
  css: '',
  markerMode: 'native',
  rawTab: 'html',
  modelTab: 'outline',
  outputTab: 'preview',
  sourceLabel: '',
};

const $ = <T extends HTMLElement>(id: string): T => {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing element: ${id}`);
  return element as T;
};

const els = {
  pasteTarget: $('paste-target'),
  panes: $('panes'),
  emptyState: $('empty-state'),
  summary: $('summary'),
  summaryGrid: $('summary-grid'),
  rawBody: $('raw-body'),
  rawSize: $('raw-size'),
  modelBody: $('model-body'),
  outputBody: $('output-body'),
  diagnosticsCount: $('diagnostics-count'),
  markerMode: $<HTMLSelectElement>('marker-mode'),
  fixturePicker: $<HTMLSelectElement>('fixture-picker'),
  fileInput: $<HTMLInputElement>('file-input'),
  footerStatus: $('footer-status'),
};

/* -------------------------------------------------------------------------
 * Pipeline
 * ---------------------------------------------------------------------- */

function analyse(payload: ClipboardPayload, sourceLabel: string): void {
  state.payload = payload;
  state.sourceLabel = sourceLabel;

  if (!payload.html) {
    state.document = null;
    state.html = '';
    state.css = '';
    setStatus('no text/html flavour on the clipboard');
    render();
    return;
  }

  const started = performance.now();
  try {
    state.document = parseWordClipboard(payload);
    const rendered = renderWordDocument(state.document, {
      markerMode: state.markerMode,
      includeWordMetadata: true,
      cssMode: 'separate',
    });
    state.html = rendered.html;
    state.css = rendered.css;
    const elapsed = performance.now() - started;
    setStatus(`${sourceLabel} · parsed and rendered in ${elapsed.toFixed(1)} ms`);
  } catch (error) {
    // A crash is a bug worth seeing in full, not a toast that disappears.
    state.document = null;
    state.html = '';
    state.css = '';
    setStatus(`error: ${(error as Error).message}`);
    els.outputBody.textContent = String((error as Error).stack ?? error);
  }
  render();
}

function rerender(): void {
  if (!state.document) return;
  const rendered = renderWordDocument(state.document, {
    markerMode: state.markerMode,
    includeWordMetadata: true,
    cssMode: 'separate',
  });
  state.html = rendered.html;
  state.css = rendered.css;
  renderOutputPane();
}

function setStatus(message: string): void {
  els.footerStatus.textContent = message;
}

/* -------------------------------------------------------------------------
 * Rendering the lab itself
 * ---------------------------------------------------------------------- */

function render(): void {
  const hasContent = Boolean(state.payload);
  els.panes.hidden = !hasContent;
  els.summary.hidden = !hasContent;
  els.emptyState.hidden = hasContent;
  if (!hasContent) return;

  renderSummary();
  renderRawPane();
  renderModelPane();
  renderOutputPane();
}

function renderSummary(): void {
  const document_ = state.document;
  els.summaryGrid.replaceChildren();
  if (!document_) {
    els.summaryGrid.append(stat('—', 'no model', 'is-error'));
    return;
  }
  const report = getFidelityReport(document_);
  const errors = document_.diagnostics.filter((d) => d.severity === 'error').length;
  const warnings = document_.diagnostics.filter((d) => d.severity === 'warning').length;

  els.summaryGrid.append(
    stat(
      report.wordDetected ? `${Math.round(report.detectionConfidence * 100)}%` : 'no',
      report.wordDetected ? `Word (${document_.detection.source})` : 'not Word',
      report.wordDetected ? 'is-ok' : 'is-warn',
    ),
    stat(String(report.paragraphs), 'paragraphs'),
    stat(String(report.runs), 'runs'),
    stat(`${report.lists}/${report.listItems}`, 'lists / items'),
    stat(String(report.maxListDepth + 1), 'list depth'),
    stat(String(report.tables), 'tables'),
    stat(
      `${report.images - report.unresolvedImages}/${report.images}`,
      'images resolved',
      report.unresolvedImages > 0 ? 'is-warn' : undefined,
    ),
    stat(String(report.hyperlinks), 'links'),
    stat(String(warnings), 'warnings', warnings > 0 ? 'is-warn' : undefined),
    stat(String(errors), 'errors', errors > 0 ? 'is-error' : undefined),
  );
}

function stat(value: string, label: string, modifier?: string): HTMLElement {
  const wrapper = document.createElement('div');
  wrapper.className = `stat${modifier ? ` ${modifier}` : ''}`;
  const valueEl = document.createElement('span');
  valueEl.className = 'stat-value';
  valueEl.textContent = value;
  const labelEl = document.createElement('span');
  labelEl.className = 'stat-label';
  labelEl.textContent = label;
  wrapper.append(valueEl, labelEl);
  return wrapper;
}

/* ------------------------------------------------------------- raw pane */

function renderRawPane(): void {
  const payload = state.payload;
  if (!payload) return;
  const html = payload.html ?? '';
  els.rawSize.textContent = `${formatBytes(html.length)}`;

  switch (state.rawTab) {
    case 'text':
      els.rawBody.textContent = payload.text ?? '(no text/plain flavour)';
      return;
    case 'types': {
      const lines = [
        `Clipboard flavours (${payload.types.length}):`,
        ...payload.types.map((type) => `  ${type}`),
        '',
        `Image items: ${payload.images.length}`,
        ...payload.images.map(
          (image, index) =>
            `  [${index}] ${image.mimeType} ${formatBytes(image.byteLength ?? 0)}${
              image.name ? ` — ${image.name}` : ''
            }`,
        ),
        '',
        `Captured: ${payload.capturedAt ? new Date(payload.capturedAt).toISOString() : 'n/a'}`,
        `Source: ${state.sourceLabel}`,
      ];
      els.rawBody.textContent = lines.join('\n');
      return;
    }
    case 'css':
      els.rawBody.textContent =
        state.document?.styles.rawCss || '(no <style> block in the payload)';
      return;
    default:
      els.rawBody.textContent = html || '(no text/html flavour)';
  }
}

/* ----------------------------------------------------------- model pane */

function renderModelPane(): void {
  const document_ = state.document;
  els.modelBody.replaceChildren();
  if (!document_) {
    els.modelBody.append(note('The payload could not be parsed into a model.'));
    return;
  }

  switch (state.modelTab) {
    case 'json':
      els.modelBody.append(code(JSON.stringify(modelForJson(document_), null, 2)));
      return;
    case 'lists':
      renderListDefinitions(document_);
      return;
    case 'styles':
      renderStyles(document_);
      return;
    case 'assets':
      renderAssets(document_);
      return;
    case 'signals':
      renderSignals(document_);
      return;
    default:
      renderOutline(document_);
  }
}

/**
 * The outline is the pane that answers "did the parser understand this?".
 *
 * Every list item shows its marker in a chip *beside* the text, which is the
 * whole point made visible: if a glyph appears in the text column instead of
 * the chip, the marker leaked into the content and there is a bug.
 */
function renderOutline(document_: WordDocument): void {
  const container = document.createElement('div');
  container.className = 'outline';

  const walk = (blocks: WordBlock[], parent: HTMLElement, depth: number): void => {
    for (const block of blocks) {
      const node = document.createElement('div');
      node.className = 'outline-node';
      node.style.marginLeft = depth === 0 ? '0' : '10px';

      switch (block.type) {
        case 'paragraph': {
          const kind = document.createElement('span');
          kind.className = 'outline-kind';
          kind.textContent = block.headingLevel ? `h${block.headingLevel}` : 'p';
          node.append(kind, ' ');

          if (block.listItem) {
            const marker = document.createElement('span');
            marker.className = 'outline-marker';
            marker.textContent =
              block.listItem.marker.type === 'bullet'
                ? (block.listItem.marker.glyph ?? '•')
                : (block.listItem.marker.text ?? '?');
            marker.title = describeMarker(block.listItem);
            node.append(marker);
          }

          const text = document.createElement('span');
          text.className = 'outline-text';
          const content = block.runs
            .map((run) =>
              run.type === 'text' ? run.text : run.type === 'image' ? '[image]' : run.type === 'tab' ? '→' : '',
            )
            .join('');
          text.textContent = content || (block.empty ? '(empty paragraph)' : '');
          node.append(text);

          const meta = document.createElement('span');
          meta.className = 'outline-meta';
          const bits: string[] = [];
          if (block.styleName) bits.push(block.styleName);
          if (block.listItem) bits.push(`L${block.listItem.level + 1}`);
          bits.push(`${block.runs.length} run${block.runs.length === 1 ? '' : 's'}`);
          meta.textContent = `  · ${bits.join(' · ')}`;
          node.append(meta);
          break;
        }
        case 'table': {
          const kind = document.createElement('span');
          kind.className = 'outline-kind';
          kind.textContent = 'table';
          const meta = document.createElement('span');
          meta.className = 'outline-meta';
          meta.textContent = `  ${block.rows.length}×${block.gridColumnCount}${
            block.depth > 0 ? ` · nested depth ${block.depth}` : ''
          }`;
          node.append(kind, meta);
          parent.append(node);
          for (const row of block.rows) {
            for (const cell of row.cells) {
              if (cell.blocks.length > 0) walk(cell.blocks, node, depth + 1);
            }
          }
          continue;
        }
        case 'container': {
          const kind = document.createElement('span');
          kind.className = 'outline-kind';
          kind.textContent = block.role;
          node.append(kind);
          parent.append(node);
          walk(block.blocks, node, depth + 1);
          continue;
        }
        default: {
          const kind = document.createElement('span');
          kind.className = 'outline-kind';
          kind.textContent = block.type;
          node.append(kind);
          if (block.type === 'unsupported') {
            const meta = document.createElement('span');
            meta.className = 'outline-meta';
            meta.textContent = `  ${block.objectType}`;
            node.append(meta);
          }
        }
      }
      parent.append(node);
    }
  };

  walk(document_.blocks, container, 0);
  els.modelBody.append(container);
}

function describeMarker(item: NonNullable<Extract<WordBlock, { type: 'paragraph' }>['listItem']>): string {
  const parts = [
    `list ${item.listId}${item.lfo ? ` (${item.lfo})` : ''}`,
    `level ${item.level + 1}`,
    `type ${item.marker.type}`,
  ];
  if (item.marker.numberFormat) parts.push(item.marker.numberFormat);
  if (item.marker.levelText) parts.push(`levelText ${item.marker.levelText}`);
  if (item.marker.rawGlyph && item.marker.rawGlyph !== item.marker.glyph) {
    parts.push(`raw U+${item.marker.rawGlyph.codePointAt(0)!.toString(16).toUpperCase()}`);
  }
  if (item.marker.font) parts.push(`font ${item.marker.font}`);
  parts.push(`source ${item.marker.source}`);
  return parts.join(' · ');
}

function renderListDefinitions(document_: WordDocument): void {
  if (document_.lists.length === 0) {
    els.modelBody.append(note('No @list definitions in this payload.'));
    return;
  }
  for (const definition of document_.lists) {
    els.modelBody.append(heading(`@list ${definition.listId}${definition.listType ? ` · ${definition.listType}` : ''}`));
    els.modelBody.append(
      table(
        ['lvl', 'type', 'format', 'level text', 'glyph', 'raw', 'font', 'start', 'indent'],
        definition.levels.map((level) => [
          String(level.level + 1),
          level.type,
          level.numberFormat,
          level.levelText ?? '',
          { glyph: level.bulletGlyph ?? '' },
          level.bulletGlyphRaw && level.bulletGlyphRaw !== level.bulletGlyph
            ? `U+${level.bulletGlyphRaw.codePointAt(0)!.toString(16).toUpperCase()}`
            : '',
          level.bulletFont ?? '',
          level.startAt !== undefined ? String(level.startAt) : '',
          level.textIndent ? `${level.textIndent.px}px` : '',
        ]),
      ),
    );
  }
}

function renderStyles(document_: WordDocument): void {
  const styles = Object.values(document_.styles.styles);
  const fonts = Object.values(document_.styles.fonts);

  els.modelBody.append(heading(`Styles (${styles.length})`));
  if (styles.length === 0) {
    els.modelBody.append(note('No style rules found.'));
  } else {
    els.modelBody.append(
      table(
        ['name', 'id', 'type', 'parent', 'declarations'],
        styles
          .slice(0, 200)
          .map((style) => [
            style.name,
            style.id,
            style.type,
            style.parent ?? '',
            String(Object.keys(style.declarations).length),
          ]),
      ),
    );
  }

  els.modelBody.append(heading(`Fonts (${fonts.length})`));
  if (fonts.length === 0) {
    els.modelBody.append(note('No @font-face rules found.'));
  } else {
    els.modelBody.append(
      table(
        ['family', 'symbol font', 'charset', 'generic'],
        fonts.map((font) => [
          font.family,
          font.isSymbolFont ? 'yes' : 'no',
          font.charset ?? '',
          font.genericFamily ?? '',
        ]),
      ),
    );
  }
}

function renderAssets(document_: WordDocument): void {
  const images = Object.values(document_.images);
  const links = Object.values(document_.hyperlinks);
  const bookmarks = Object.values(document_.bookmarks);

  els.modelBody.append(heading(`Images (${images.length})`));
  if (images.length === 0) els.modelBody.append(note('No images.'));
  else
    els.modelBody.append(
      table(
        ['id', 'resolution', 'origin', 'size', 'original reference'],
        images.map((image) => [
          image.id,
          image.resolution,
          image.origin,
          image.width && image.height ? `${round(image.width.px)}×${round(image.height.px)}` : '',
          image.originalSource.slice(0, 90),
        ]),
      ),
    );

  els.modelBody.append(heading(`Hyperlinks (${links.length})`));
  if (links.length === 0) els.modelBody.append(note('No hyperlinks.'));
  else
    els.modelBody.append(
      table(
        ['id', 'href', 'blocked'],
        links.map((link) => [link.id, link.href || link.rawHref, link.blocked ? 'blocked' : '']),
      ),
    );

  els.modelBody.append(heading(`Bookmarks (${bookmarks.length})`));
  if (bookmarks.length === 0) els.modelBody.append(note('No bookmarks.'));
  else
    els.modelBody.append(
      table(
        ['name', 'internal'],
        bookmarks.map((bookmark) => [bookmark.name, bookmark.internal ? 'Word internal' : '']),
      ),
    );
}

function renderSignals(document_: WordDocument): void {
  const detection = document_.detection;
  els.modelBody.append(
    heading(
      `Detection: ${detection.isWord ? 'Word' : 'not Word'} · confidence ${detection.confidence.toFixed(2)} · ${detection.source}`,
    ),
  );
  els.modelBody.append(
    table(
      ['signal', 'strength', 'detail'],
      detection.detail.map((signal) => [signal.id, signal.strength, signal.description]),
    ),
  );

  const metadata = document_.metadata;
  els.modelBody.append(heading('Payload metadata'));
  const rows: Array<[string, string]> = [
    ['generator', metadata.generator ?? ''],
    ['progId', metadata.progId ?? ''],
    ['word version', metadata.wordVersion !== undefined ? String(metadata.wordVersion) : ''],
    ['fragment markers', metadata.fragmentBoundaryFound ? 'found' : 'absent'],
    ['sections', metadata.sections.join(', ')],
    ['namespaces', Object.keys(metadata.namespaces).join(', ')],
    ['raw length', formatBytes(metadata.rawHtmlLength ?? 0)],
  ];
  for (const [key, value] of Object.entries(metadata.documentProperties)) {
    rows.push([`property: ${key}`, value]);
  }
  els.modelBody.append(table(['key', 'value'], rows.filter(([, value]) => value !== '')));
}

/**
 * The JSON view drops the raw payload and the parsed stylesheet.
 *
 * Both are available in their own tabs, and including them here would bury the
 * content under a megabyte of Word's CSS — the opposite of a debugging aid.
 */
function modelForJson(document_: WordDocument): unknown {
  const { rawHtml, rawText, styles, ...rest } = document_;
  void rawHtml;
  void rawText;
  void styles;
  return { ...rest, styles: '(see the Styles tab)' };
}

/* ---------------------------------------------------------- output pane */

function renderOutputPane(): void {
  els.outputBody.replaceChildren();
  const document_ = state.document;
  if (!document_) {
    els.outputBody.append(note('Nothing to render.'));
    return;
  }

  const diagnostics = document_.diagnostics;
  const errors = diagnostics.filter((d) => d.severity === 'error').length;
  const warnings = diagnostics.filter((d) => d.severity === 'warning').length;
  els.diagnosticsCount.textContent = String(diagnostics.length);
  els.diagnosticsCount.className = `badge${errors > 0 ? ' is-error' : warnings > 0 ? ' is-warn' : ''}`;

  switch (state.outputTab) {
    case 'html':
      els.outputBody.append(code(prettyHtml(state.html)));
      return;
    case 'css':
      els.outputBody.append(code(state.css));
      return;
    case 'report':
      els.outputBody.append(code(formatFidelityReport(getFidelityReport(document_))));
      return;
    case 'diagnostics':
      renderDiagnostics(diagnostics);
      return;
    default:
      renderPreview();
  }
}

/**
 * The preview shows the engine's output *with the engine's own stylesheet*.
 *
 * The generated CSS is scoped into a `<style>` alongside the fragment, exactly
 * as a consumer would place it. Nothing in the lab's stylesheet reaches inside
 * — a preview that the lab has prettified is a preview that lies about what
 * the engine produced.
 */
function renderPreview(): void {
  const wrapper = document.createElement('div');
  wrapper.className = 'preview';
  // The rendered HTML is the engine's own output, built by string construction
  // from escaped model values — never a clone of the untrusted payload.
  wrapper.innerHTML = `<style>${state.css}</style>${state.html}`;
  els.outputBody.append(wrapper);
}

function renderDiagnostics(diagnostics: WordDiagnostic[]): void {
  if (diagnostics.length === 0) {
    els.outputBody.append(note('No diagnostics. Everything in this payload was represented exactly.'));
    return;
  }

  const order = { error: 0, warning: 1, info: 2 } as const;
  const sorted = [...diagnostics].sort((a, b) => order[a.severity] - order[b.severity]);

  for (const diagnostic of sorted) {
    const item = document.createElement('div');
    item.className = `diagnostic sev-${diagnostic.severity}`;

    const head = document.createElement('div');
    head.className = 'diagnostic-head';

    const codeEl = document.createElement('span');
    codeEl.className = 'diagnostic-code';
    codeEl.textContent = diagnostic.code;

    const fidelity = document.createElement('span');
    fidelity.className = `fidelity fidelity-${diagnostic.fidelity}`;
    fidelity.textContent = diagnostic.fidelity;

    head.append(codeEl, fidelity);
    if ((diagnostic.count ?? 1) > 1) {
      const count = document.createElement('span');
      count.className = 'pill';
      count.textContent = `×${diagnostic.count}`;
      head.append(count);
    }

    const message = document.createElement('div');
    message.className = 'diagnostic-message';
    message.textContent = diagnostic.message;

    item.append(head, message);

    const explanation = DiagnosticExplanations[diagnostic.code];
    if (explanation && explanation !== diagnostic.message) {
      const why = document.createElement('div');
      why.className = 'diagnostic-details';
      why.textContent = explanation;
      item.append(why);
    }

    if (diagnostic.location || diagnostic.details) {
      const details = document.createElement('div');
      details.className = 'diagnostic-details';
      const bits: string[] = [];
      if (diagnostic.location?.tagName) bits.push(`<${diagnostic.location.tagName}>`);
      if (diagnostic.location?.path) bits.push(diagnostic.location.path);
      if (diagnostic.details) {
        for (const [key, value] of Object.entries(diagnostic.details)) {
          bits.push(`${key}=${String(value).slice(0, 80)}`);
        }
      }
      details.textContent = bits.join('  ');
      item.append(details);
    }

    els.outputBody.append(item);
  }
}

/* -------------------------------------------------------------------------
 * Small DOM helpers
 * ---------------------------------------------------------------------- */

function note(text: string): HTMLElement {
  const element = document.createElement('p');
  element.className = 'note';
  element.textContent = text;
  return element;
}

function heading(text: string): HTMLElement {
  const element = document.createElement('h3');
  element.className = 'section-heading';
  element.textContent = text;
  return element;
}

function code(text: string): HTMLElement {
  const element = document.createElement('pre');
  element.className = 'code';
  element.textContent = text;
  return element;
}

type Cell = string | { glyph: string };

function table(headers: string[], rows: Cell[][]): HTMLElement {
  const element = document.createElement('table');
  element.className = 'data-table';

  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  for (const header of headers) {
    const th = document.createElement('th');
    th.textContent = header;
    headRow.append(th);
  }
  thead.append(headRow);

  const tbody = document.createElement('tbody');
  for (const row of rows) {
    const tr = document.createElement('tr');
    for (const cell of row) {
      const td = document.createElement('td');
      if (typeof cell === 'string') {
        td.textContent = cell;
      } else {
        td.className = 'glyph-cell';
        td.textContent = cell.glyph;
      }
      tr.append(td);
    }
    tbody.append(tr);
  }

  element.append(thead, tbody);
  return element;
}

function formatBytes(count: number): string {
  if (count < 1024) return `${count} B`;
  if (count < 1024 * 1024) return `${(count / 1024).toFixed(1)} kB`;
  return `${(count / (1024 * 1024)).toFixed(2)} MB`;
}

function round(value: number): number {
  return Math.round(value * 10) / 10;
}

/** Line-break the rendered HTML so the pane is readable. Display only. */
function prettyHtml(html: string): string {
  return html.replace(/></g, '>\n<');
}

/* -------------------------------------------------------------------------
 * Wiring
 * ---------------------------------------------------------------------- */

installPasteCapture(els.pasteTarget, (payload) => {
  analyse(payload, 'paste event');
  els.pasteTarget.replaceChildren();
  const summary = document.createElement('p');
  summary.className = 'paste-hint';
  summary.textContent = `Pasted ${formatBytes((payload.html ?? '').length)} of clipboard HTML. Paste again to replace.`;
  els.pasteTarget.append(summary);
});

$('read-clipboard').addEventListener('click', () => {
  void (async () => {
    try {
      const payload = await captureFromNavigatorClipboard();
      if (!payload.html) {
        setStatus('the async clipboard API returned no text/html flavour');
        return;
      }
      analyse(payload, 'navigator.clipboard');
    } catch (error) {
      // Reading the clipboard programmatically needs permission and a user
      // gesture, and browsers differ on both. Say which one failed.
      setStatus(`clipboard read failed: ${(error as Error).message} — try pasting instead`);
    }
  })();
});

els.fileInput.addEventListener('change', () => {
  const file = els.fileInput.files?.[0];
  if (!file) return;
  void file.text().then((text) => {
    analyse(clipboardPayloadFromHtml(text), `file: ${file.name}`);
  });
});

for (const fixture of fixtures) {
  const option = document.createElement('option');
  option.value = fixture.id;
  option.textContent = fixture.id;
  els.fixturePicker.append(option);
}

els.fixturePicker.addEventListener('change', () => {
  const fixture = fixtures.find((f) => f.id === els.fixturePicker.value);
  if (!fixture) return;
  analyse(clipboardPayloadFromHtml(fixture.html), `fixture: ${fixture.id}`);
});

$('clear').addEventListener('click', () => {
  state.payload = null;
  state.document = null;
  state.html = '';
  state.css = '';
  els.fixturePicker.value = '';
  els.pasteTarget.replaceChildren();
  const hint = document.createElement('p');
  hint.className = 'paste-hint';
  hint.textContent = 'Copy content in Word, click here, and paste.';
  els.pasteTarget.append(hint);
  setStatus('idle');
  render();
});

els.markerMode.addEventListener('change', () => {
  state.markerMode = els.markerMode.value as ListMarkerMode;
  rerender();
});

/* ------------------------------------------------------------------ tabs */

function wireTabs(containerId: string, attribute: string, apply: (value: string) => void): void {
  const container = $(containerId);
  container.addEventListener('click', (event) => {
    const target = (event.target as HTMLElement).closest<HTMLElement>(`[${attribute}]`);
    if (!target) return;
    for (const tab of container.querySelectorAll('.tab')) tab.classList.remove('is-active');
    target.classList.add('is-active');
    apply(target.getAttribute(attribute)!);
  });
}

wireTabs('raw-tabs', 'data-raw-tab', (value) => {
  state.rawTab = value;
  renderRawPane();
});
wireTabs('model-tabs', 'data-model-tab', (value) => {
  state.modelTab = value;
  renderModelPane();
});
wireTabs('output-tabs', 'data-output-tab', (value) => {
  state.outputTab = value;
  renderOutputPane();
});

/* ------------------------------------------------------------ copy/export */

document.addEventListener('click', (event) => {
  const button = (event.target as HTMLElement).closest<HTMLElement>('[data-copy]');
  if (!button) return;
  const what = button.getAttribute('data-copy');
  const text =
    what === 'raw'
      ? (state.payload?.html ?? '')
      : what === 'model'
        ? JSON.stringify(modelForJson(state.document ?? ({} as WordDocument)), null, 2)
        : state.html;
  void navigator.clipboard.writeText(text).then(
    () => setStatus(`copied ${formatBytes(text.length)} to the clipboard`),
    (error: Error) => setStatus(`copy failed: ${error.message}`),
  );
});

$('download-html').addEventListener('click', () => {
  if (!state.document) return;
  const { document: file } = renderStandaloneHtml(state.document, {
    markerMode: state.markerMode,
    includeWordMetadata: true,
    includeDiagnosticsAppendix: false,
  });
  downloadFile(file, suggestFileName(state.document), 'text/html;charset=utf-8');
  setStatus(`downloaded ${formatBytes(file.length)} of standalone HTML`);
});

function downloadFile(content: string, fileName: string, mimeType: string): void {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  // Revoking immediately can cancel the download in some browsers.
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/* -------------------------------------------------------------------------
 * Fixture export
 *
 * The spec's rule is that every bug found in real Word content becomes a
 * fixture. This turns whatever is currently loaded into the three files the
 * fixture library expects, so capturing a regression is a single click rather
 * than a manual transcription of a megabyte of clipboard HTML.
 * ---------------------------------------------------------------------- */

const exportButton = document.createElement('button');
exportButton.type = 'button';
exportButton.className = 'button button-small';
exportButton.textContent = 'Export fixture';
exportButton.addEventListener('click', () => {
  const payload = state.payload;
  if (!payload?.html) return;
  const name = window.prompt(
    'Fixture id (e.g. bullets/real-word-16-bullets):',
    'captured/word-paste',
  );
  if (!name) return;
  const safe = name.replace(/[^a-z0-9/_-]/gi, '-');
  downloadFile(payload.html, `${safe.split('/').pop()}--input.html`, 'text/html;charset=utf-8');
  setStatus(
    `saved input.html — put it in src/fixtures/word/${safe}/ and run UPDATE_FIXTURES=1 npm test to bless it`,
  );
});
document.querySelector('.pane-tools')?.append(exportButton);

/* -------------------------------------------------------------------- init */

// Load a fixture on first visit so the lab is never an empty page: seeing what
// correct output looks like is the fastest way to recognise incorrect output.
// The gallery links here with ?fixture=<id>, so a specific one can be opened.
const requested = new URLSearchParams(window.location.search).get('fixture');
const initial =
  (requested ? fixtures.find((f) => f.id === requested) : undefined) ??
  fixtures.find((f) => f.id === 'complex/engagement-report') ??
  fixtures[0];
if (initial) {
  els.fixturePicker.value = initial.id;
  analyse(clipboardPayloadFromHtml(initial.html), `fixture: ${initial.id}`);
}

// Detection runs on every payload; expose it for console debugging too.
Object.assign(window as unknown as Record<string, unknown>, {
  wce: {
    get state() {
      return state;
    },
    detect: detectWordHtml,
    parse: parseWordClipboard,
    render: renderWordDocument,
    fixtures,
  },
});
