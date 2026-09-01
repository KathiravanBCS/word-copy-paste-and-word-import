import { describe, expect, it } from 'vitest';
import { parseWordHtmlString, renderWordDocument, renderStandaloneHtml, suggestFileName } from '../../index.js';
import { loadFixtures } from '../support/fixtures.js';
import type { WordDocument } from '../../model/Document.js';

const fixtures = loadFixtures();
const byId = (id: string): WordDocument => {
  const fixture = fixtures.find((f) => f.id === id);
  if (!fixture) throw new Error(`Missing fixture: ${id}`);
  return parseWordHtmlString(fixture.inputHtml);
};

describe('list rendering', () => {
  it('compiles Word bullets into a native counter style', () => {
    const { html, css } = renderWordDocument(byId('bullets/bullet-default'));
    expect(html).toContain('<ul');
    expect(html).not.toContain('<span class="wce-marker"');
    expect(css).toContain('system: cyclic');
    expect(css).toContain('symbols: "•"');
    expect(css).toContain('symbols: "o"');
    expect(css).toContain('symbols: "▪"');
    expect(css).toContain('list-style-type: wce-');
  });

  it('compiles a simple Word number format into a native counter style', () => {
    const { css } = renderWordDocument(byId('numbering/roman-numbering'));
    expect(css).toContain('system: extends upper-roman');
    expect(css).toContain('suffix: ". "');
  });

  it('falls back to a real ::marker for a level text CSS cannot express', () => {
    const { html, css } = renderWordDocument(byId('numbering/multilevel-numbering'));
    // `%1.%2` has no counter-style equivalent, so the literal marker is put on
    // the item and drawn by ::marker — still a marker, never text.
    expect(html).toContain('data-marker="1.1"');
    expect(css).toContain('::marker');
    expect(css).toContain('content: attr(data-marker)');
    expect(html).not.toMatch(/>1\.1\s*Orbis/);
  });

  it('renders a marker element when asked to', () => {
    const { html, css } = renderWordDocument(byId('bullets/bullet-default'), {
      markerMode: 'element',
    });
    expect(html).toContain('<span class="wce-marker" aria-hidden="true">•</span>');
    expect(css).toContain('list-style-type: none');
    expect(css).toContain('text-indent: -24px');
  });

  it('sets an explicit value so a copied list keeps Word’s numbers', () => {
    const { html } = renderWordDocument(byId('numbering/start-at'));
    expect(html).toContain('value="5"');
    expect(html).toContain('value="6"');
    expect(html).toContain('value="7"');
  });

  it('preserves the marker font when the glyph could not be mapped', () => {
    const { css } = renderWordDocument(byId('bullets/bullet-default'));
    // Word's level-2 bullet really is the letter "o" in Courier New.
    expect(css).toContain('font-family: "Courier New"');
  });

  it('nests lists rather than repeating the indent', () => {
    const { css } = renderWordDocument(byId('bullets/bullet-default'));
    const paddings = [...css.matchAll(/padding-left: ([\d.]+)px/g)].map((m) => Number(m[1]));
    // Each level adds half an inch relative to its parent, not absolutely.
    expect(paddings).toEqual([48, 48, 48]);
  });
});

describe('run rendering', () => {
  it('prefers semantic elements over styled spans', () => {
    const { html } = renderWordDocument(byId('formatting/character-formatting'));
    expect(html).toContain('<strong>Bold</strong>');
    expect(html).toContain('<em>Italic</em>');
    expect(html).toContain('<u>Underline</u>');
    expect(html).toContain('<s>Strikethrough</s>');
    expect(html).toContain('<sub>2</sub>');
    expect(html).toContain('<sup>2</sup>');
  });

  it('keeps run boundaries in the output', () => {
    const { html } = renderWordDocument(byId('formatting/character-formatting'));
    expect(html).toContain('Hello <span style="color:#c00000"><strong>world</strong></span> and <em>goodbye</em>.');
  });

  it('keeps formatting inside a hyperlink', () => {
    const { html } = renderWordDocument(byId('mixed/hyperlinks'));
    expect(html).toContain('<a class="wce-link" href="https://example.com/report">the <strong>full report</strong></a>');
  });

  it('marks an unfollowable link rather than emitting a dead href', () => {
    const { html } = renderWordDocument(byId('mixed/hyperlinks'));
    expect(html).toContain('data-word-unresolved-href="file:///');
    // No *navigable* href — the reference survives only as data.
    expect(html).not.toMatch(/\shref="file:/);
  });
});

describe('table rendering', () => {
  it('emits thead, colgroup and spans', () => {
    const { html } = renderWordDocument(byId('tables/simple-table'));
    expect(html).toContain('<thead>');
    expect(html).toContain('<th');
    expect(html).toContain('scope="col"');
    expect(html).toContain('<colgroup>');
    expect(html).toContain('border-collapse:collapse');
  });

  it('renders merged cells with colspan and rowspan', () => {
    const { html } = renderWordDocument(byId('tables/merged-cells'));
    expect(html).toContain('colspan="3"');
    expect(html).toContain('rowspan="2"');
    expect(html).toContain('colspan="2"');
  });

  it('renders a nested table as a real nested table', () => {
    const { html } = renderWordDocument(byId('tables/nested-table'));
    const tables = html.match(/<table/g) ?? [];
    expect(tables).toHaveLength(2);
    expect(html).toMatch(/<td[^>]*>.*<table/s);
  });
});

describe('security', () => {
  const hostile = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office">
    <meta name=ProgId content=Word.Document>
    <meta name=Generator content="Microsoft Word 15">
    <style><!-- p.MsoNormal {margin:0in;} --></style>
    <body><div class=WordSection1>
      <script>window.stolen = document.cookie;</script>
      <p class=MsoNormal onclick="alert(1)" onmouseover="alert(2)">Click me</p>
      <p class=MsoNormal><a href="javascript:alert(3)">A link</a></p>
      <p class=MsoNormal><a href="  jav&#x09;ascript:alert(4)">Obfuscated</a></p>
      <p class=MsoNormal><img src="javascript:alert(5)" alt="x"></p>
      <p class=MsoNormal><iframe src="https://evil.example"></iframe></p>
      <p class=MsoNormal>Text with &lt;script&gt;alert(6)&lt;/script&gt; escaped.</p>
      <p class=MsoNormal><span style="background:url(javascript:alert(7))">Styled</span></p>
    </div></body></html>`;

  it('never emits script, handlers or javascript URLs', () => {
    const document = parseWordHtmlString(hostile);
    const { html } = renderWordDocument(document);
    expect(html).not.toMatch(/<script/i);
    expect(html).not.toMatch(/<iframe/i);
    expect(html).not.toMatch(/\son[a-z]+\s*=/i);
    expect(html).not.toMatch(/javascript:/i);
    expect(html).not.toMatch(/window\.stolen/);
  });

  it('escapes text that looks like markup', () => {
    const document = parseWordHtmlString(hostile);
    const { html } = renderWordDocument(document);
    expect(html).toContain('&lt;script&gt;alert(6)&lt;/script&gt;');
  });

  it('reports what it removed instead of removing it silently', () => {
    const document = parseWordHtmlString(hostile);
    const codes = document.diagnostics.map((d) => d.code);
    expect(codes).toContain('SECURITY_SCRIPT_REMOVED');
    expect(codes).toContain('SECURITY_EVENT_HANDLER_REMOVED');
    expect(codes).toContain('SECURITY_URL_BLOCKED');
  });

  it('keeps the raw payload untouched for inspection', () => {
    const document = parseWordHtmlString(hostile);
    expect(document.rawHtml).toBe(hostile);
    expect(document.rawHtml).toContain('<script>');
  });

  /**
   * The regression that mattered most.
   *
   * `<link>` is a void element. Word puts `<link rel=File-List …>` in the head
   * of every clipboard payload. Neutralising it by *renaming* it to a custom
   * element turns it into a non-void one: a real browser then hoists it into
   * the body, leaves it open, and every subsequent element becomes its child.
   * Removing it afterwards took the entire document with it, and a 162 kB
   * paste produced zero paragraphs — silently, because nothing threw.
   *
   * happy-dom hoists it differently and did not reproduce this, which is
   * exactly why the fixtures now carry the real Word head.
   */
  it('survives the <link> elements Word puts in every payload', () => {
    const withLinks = `<html xmlns:o="urn:schemas-microsoft-com:office:office">
      <meta name=ProgId content=Word.Document>
      <meta name=Generator content="Microsoft Word 15 (filtered medium)">
      <link rel=File-List href="file:///C:/Users/x/Temp/msohtmlclip1/01/clip_filelist.xml">
      <link rel=Edit-Time-Data href="file:///C:/Users/x/Temp/msohtmlclip1/01/clip_editdata.mso">
      <style><!-- p.MsoNormal {margin:0in;} --></style>
      <body><div class=WordSection1><!--StartFragment-->
      <p class=MsoNormal>First paragraph.<o:p></o:p></p>
      <p class=MsoNormal>Second paragraph.<o:p></o:p></p>
      <!--EndFragment--></div></body></html>`;

    const document = parseWordHtmlString(withLinks);
    expect(document.blocks).toHaveLength(2);
    expect(renderWordDocument(document).html).toContain('First paragraph.');
    expect(renderWordDocument(document).html).not.toContain('clip_filelist');
  });

  it('removes void elements from the text rather than renaming them', () => {
    // A renamed void element becomes a container; deleting from the source
    // text is the only handling that cannot swallow the document.
    for (const tag of ['link', 'base', 'embed', 'frame']) {
      const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office">
        <meta name=ProgId content=Word.Document>
        <style><!-- p.MsoNormal {margin:0in;} --></style>
        <body><div class=WordSection1>
        <${tag} href="https://evil.example" src="https://evil.example">
        <p class=MsoNormal>Content after the ${tag}.</p>
        </div></body></html>`;
      const document = parseWordHtmlString(html);
      const text = document.blocks
        .map((b) => (b.type === 'paragraph' ? b.runs.map((r) => (r.type === 'text' ? r.text : '')).join('') : ''))
        .join('');
      expect(text, tag).toContain(`Content after the ${tag}.`);
      expect(renderWordDocument(document).html, tag).not.toContain('evil.example');
    }
  });

  it('bounds a pathologically nested payload', () => {
    const deep = '<div>'.repeat(400) + '<p>deep</p>' + '</div>'.repeat(400);
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office">
      <meta name=ProgId content=Word.Document><body>${deep}</body></html>`;
    const document = parseWordHtmlString(html, { limits: { maxDepth: 32 } });
    expect(document.blocks.length).toBeGreaterThanOrEqual(0);
    expect(() => renderWordDocument(document)).not.toThrow();
  });
});

describe('standalone document', () => {
  it('produces a complete, self-contained HTML file', () => {
    const document = byId('complex/engagement-report');
    const { document: file, report } = renderStandaloneHtml(document);

    expect(file.startsWith('<!doctype html>')).toBe(true);
    expect(file).toContain('<meta charset="utf-8">');
    expect(file).toContain('<title>');
    expect(file).toContain('</html>');
    // The generated counter styles must be in the head, or the native markers
    // do not survive the trip to a file on disk.
    expect(file).toContain('@counter-style');
    expect(file).toContain('<style>');
    expect(file).not.toContain('<script');
    expect(report.paragraphs).toBeGreaterThan(0);
  });

  it('appends a machine-readable fidelity report', () => {
    const { document: file } = renderStandaloneHtml(byId('images/unresolved-image'));
    expect(file).toContain('word-clipboard-engine fidelity report');
    expect(file).toContain('"unresolvedImages": 1');
  });

  it('can include a visible diagnostics appendix', () => {
    const { document: file } = renderStandaloneHtml(byId('images/unresolved-image'), {
      includeDiagnosticsAppendix: true,
    });
    expect(file).toContain('Conversion notes');
    expect(file).toContain('WORD_UNRESOLVED_IMAGE');
  });

  it('suggests a filename from the document', () => {
    expect(suggestFileName(byId('basic/headings'))).toBe('background.html');
    expect(suggestFileName(byId('basic/plain-paragraphs'))).toBe('word-paste.html');
  });
});
