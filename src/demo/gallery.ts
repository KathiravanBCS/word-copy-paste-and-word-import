import { getFidelityReport, parseWordHtmlString, renderWordDocument } from '../index.js';

/**
 * The fixture gallery.
 *
 * Every fixture rendered on one page, through the same code path a real paste
 * takes. It is the fastest way to see the effect of a change across every Word
 * construct at once — a regression in bullet glyphs or table borders is
 * obvious here in a way it is not in a diff.
 */

const inputs = import.meta.glob('../fixtures/word/**/input.html', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>;

const notes = import.meta.glob('../fixtures/word/**/notes.md', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>;

interface GalleryFixture {
  id: string;
  category: string;
  name: string;
  html: string;
  summary: string;
}

const fixtures: GalleryFixture[] = Object.entries(inputs)
  .map(([path, html]) => {
    const id = path.replace(/^.*fixtures\/word\//, '').replace(/\/input\.html$/, '');
    const noteKey = path.replace(/input\.html$/, 'notes.md');
    const note = notes[noteKey] ?? '';
    const [category = 'root', ...rest] = id.split('/');
    return {
      id,
      category,
      name: rest.join('/') || category,
      html,
      summary: firstParagraph(note),
    };
  })
  .sort((a, b) => a.id.localeCompare(b.id));

/** The first prose paragraph of a fixture's notes.md, for the card. */
function firstParagraph(markdown: string): string {
  const lines = markdown.split('\n');
  const body: string[] = [];
  for (const line of lines) {
    if (line.startsWith('#')) continue;
    if (line.trim() === '') {
      if (body.length > 0) break;
      continue;
    }
    body.push(line.trim());
  }
  return body.join(' ').replace(/`/g, '').slice(0, 220);
}

const container = document.getElementById('gallery-body')!;
const status = document.getElementById('gallery-status')!;

const byCategory = new Map<string, GalleryFixture[]>();
for (const fixture of fixtures) {
  const list = byCategory.get(fixture.category) ?? [];
  list.push(fixture);
  byCategory.set(fixture.category, list);
}

let totalDiagnostics = 0;

for (const [category, list] of byCategory) {
  const heading = document.createElement('h2');
  heading.textContent = category;
  container.append(heading);

  for (const fixture of list) {
    container.append(renderFixture(fixture));
  }
}

status.textContent = `${fixtures.length} fixtures · ${totalDiagnostics} diagnostics`;

function renderFixture(fixture: GalleryFixture): HTMLElement {
  const section = document.createElement('section');
  section.className = 'fixture-card';
  section.style.marginBottom = '14px';

  const title = document.createElement('h3');
  title.textContent = fixture.id;
  section.append(title);

  if (fixture.summary) {
    const summary = document.createElement('p');
    summary.textContent = fixture.summary;
    section.append(summary);
  }

  try {
    const document_ = parseWordHtmlString(fixture.html);
    const { html, css } = renderWordDocument(document_, { markerMode: 'native' });
    const report = getFidelityReport(document_);
    totalDiagnostics += document_.diagnostics.length;

    const stats = document.createElement('p');
    stats.style.marginTop = '6px';
    stats.textContent =
      `${report.paragraphs} paragraphs · ${report.listItems} list items · ` +
      `${report.tables} tables · ${report.images} images · ` +
      `${report.fidelityBreakdown.APPROXIMATED} approximated · ` +
      `${report.fidelityBreakdown.UNSUPPORTED} unsupported`;
    section.append(stats);

    const preview = document.createElement('div');
    preview.className = 'preview';
    preview.style.marginTop = '10px';
    // The engine's own output plus the engine's own stylesheet, exactly as a
    // consumer would place them. The gallery adds nothing.
    preview.innerHTML = `<style>${css}</style>${html}`;
    section.append(preview);
  } catch (error) {
    const failure = document.createElement('pre');
    failure.className = 'code';
    failure.style.color = 'var(--error)';
    failure.textContent = `Failed to render: ${(error as Error).stack ?? String(error)}`;
    section.append(failure);
  }

  const link = document.createElement('p');
  link.style.marginTop = '8px';
  const anchor = document.createElement('a');
  anchor.href = `./clipboard-lab/?fixture=${encodeURIComponent(fixture.id)}`;
  anchor.textContent = 'Open in Clipboard Lab →';
  link.append(anchor);
  section.append(link);

  return section;
}
