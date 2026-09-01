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
  it('defaults to element markers, whose gap matches Word — a native ::marker does not', () => {
    // This is the whole reason 'element' is the default rather than 'native'.
    // A native CSS counter-style / ::marker is right-aligned within its
    // gutter: the browser piles any blank space to the LEFT of the digits,
    // flush against the text on the right. Word's numbering is a literal tab
    // after the number, landing at a fixed column with the blank space AFTER
    // the number, before the text. Verified directly in Chromium with
    // identical gutter widths: native renders "1. Text" (no visible gap),
    // element renders "1.1.1    Text" (Word's actual spacing) — see the
    // 'element marker mode' describe block below for the geometry assertion.
    const { html } = renderWordDocument(byId('bullets/bullet-default'));
    expect(html).toContain('<span class="wce-marker" aria-hidden="true">•</span>');
    expect(html).not.toContain('list-style-type: wce-');
  });

  it('compiles Word bullets into a native counter style, when explicitly asked for', () => {
    const { html, css } = renderWordDocument(byId('bullets/bullet-default'), {
      markerMode: 'native',
    });
    expect(html).toContain('<ul');
    expect(html).not.toContain('<span class="wce-marker"');
    expect(css).toContain('system: cyclic');
    expect(css).toContain('symbols: "•"');
    expect(css).toContain('symbols: "o"');
    expect(css).toContain('symbols: "▪"');
    expect(css).toContain('list-style-type: wce-');
  });

  it('compiles a simple Word number format into a native counter style, when asked for', () => {
    const { css } = renderWordDocument(byId('numbering/roman-numbering'), { markerMode: 'native' });
    expect(css).toContain('system: extends upper-roman');
    expect(css).toContain('suffix: ". "');
  });

  it('falls back to a real ::marker for a level text CSS cannot express, in native mode', () => {
    const { html, css } = renderWordDocument(byId('numbering/multilevel-numbering'), {
      markerMode: 'native',
    });
    // `%1.%2` has no counter-style equivalent, so the literal marker is put on
    // the item and drawn by ::marker — still a marker, never text.
    expect(html).toContain('data-marker="1.1"');
    expect(css).toContain('::marker');
    expect(css).toContain('content: attr(data-marker)');
    expect(html).not.toMatch(/>1\.1\s*Orbis/);
  });

  it('renders a marker element by default, and can render native markers when asked to', () => {
    const byDefault = renderWordDocument(byId('bullets/bullet-default'));
    expect(byDefault.html).toContain('<span class="wce-marker" aria-hidden="true">•</span>');
    expect(byDefault.css).toContain('list-style-type: none');
    expect(byDefault.css).toContain('text-indent: -24px');

    const native = renderWordDocument(byId('bullets/bullet-default'), { markerMode: 'native' });
    expect(native.html).not.toContain('wce-marker');
    expect(native.css).toContain('list-style-type: wce-');
  });

  it('shows Word’s own numbers as the literal marker text by default', () => {
    // In element mode the number IS the visible text (no CSS counter to keep
    // in sync), so the start-at value shows up as ordinary marker text rather
    // than as a native <li value>.
    const { html } = renderWordDocument(byId('numbering/start-at'));
    expect(html).toContain('>5.</span>');
    expect(html).toContain('>6.</span>');
    expect(html).toContain('>7.</span>');
  });

  it('sets an explicit <li value> in native mode so a copied list keeps Word’s numbers', () => {
    const { html } = renderWordDocument(byId('numbering/start-at'), { markerMode: 'native' });
    expect(html).toContain('value="5"');
    expect(html).toContain('value="6"');
    expect(html).toContain('value="7"');
  });

  it('preserves the marker font when the glyph could not be mapped', () => {
    const { css } = renderWordDocument(byId('bullets/bullet-default'), { markerMode: 'native' });
    // Word's level-2 bullet really is the letter "o" in Courier New.
    expect(css).toContain('font-family: "Courier New"');
  });

  it('nests lists rather than repeating the indent', () => {
    const { css } = renderWordDocument(byId('bullets/bullet-default'), { markerMode: 'native' });
    const paddings = [...css.matchAll(/padding-left: ([\d.]+)px/g)].map((m) => Number(m[1]));
    // Each level adds half an inch relative to its parent, not absolutely.
    expect(paddings).toEqual([48, 48, 48]);
  });
});

describe('element marker mode geometry', () => {
  it('reserves a fixed-width gutter so the marker sits left-aligned with a real gap before the text', () => {
    // The property under test, stated precisely: in element mode the marker
    // span is `min-width: hangingPx` and the <li> uses `text-indent:
    // -hangingPx`, which is what makes an inline-block marker narrower than
    // its reserved width leave blank space to its OWN right (between the
    // marker and the text) rather than to its left (Chromium's native
    // ::marker behaviour) — see docs/LIST-PARSING.md.
    const { css } = renderWordDocument(byId('numbering/multilevel-numbering'));
    expect(css).toMatch(/> li > \.wce-marker \{[^}]*display: inline-block;/);
    expect(css).toMatch(/> li > \.wce-marker \{[^}]*min-width: \d+px;/);
    expect(css).toMatch(/> li \{[^}]*text-indent: -\d+px;/);
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
    // The generated list-marker geometry must be in the head, or the markers
    // do not survive the trip to a file on disk.
    expect(file).toContain('wce-marker');
    expect(file).toContain('min-width:');
    expect(file).toContain('<style>');
    expect(file).not.toContain('<script');
    expect(report.paragraphs).toBeGreaterThan(0);
  });

  it('still puts generated @counter-style rules in the head in native mode', () => {
    const document = byId('numbering/roman-numbering');
    const { document: file } = renderStandaloneHtml(document, { markerMode: 'native' });
    expect(file).toContain('@counter-style');
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

describe('standalone document page geometry', () => {
  it('uses the payload’s own @page size and margins for the page shell', () => {
    // The engagement-report fixture declares @page WordSection1 {size:8.5in
    // 11.0in; margin:1.0in 1.0in 1.0in 1.0in;} — content width should be
    // 8.5in - 1in - 1in = 6.5in, not the 6.5in *default* that would also
    // happen to match here by coincidence; the asymmetric case below is the
    // one that actually distinguishes "read from the payload" from "guessed".
    const document = byId('complex/engagement-report');
    const { document: file } = renderStandaloneHtml(document);
    expect(file).toContain('max-width: 624px');
    expect(file).toContain('padding: 96px 96px 96px 96px');
  });

  it('reflects asymmetric margins and a non-Letter page size', () => {
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office">
      <meta name=ProgId content=Word.Document>
      <style><!--
        @page WordSection1 {size:8.27in 11.69in; margin:0.5in 1.5in 0.5in 2.0in;}
        div.WordSection1 {page:WordSection1;}
      --></style>
      <body><div class=WordSection1><!--StartFragment-->
      <p class=MsoNormal>A4 page, uneven margins.<o:p></o:p></p>
      <!--EndFragment--></div></body></html>`;
    const document = parseWordHtmlString(html);
    const { document: file } = renderStandaloneHtml(document);
    // content width = 8.27in - 2.0in - 1.5in = 4.77in = 4.77*96 = 457.92px
    expect(file).toContain('max-width: 457.92px');
    // margin order is CSS box order: top right bottom left
    expect(file).toContain('padding: 48px 144px 48px 192px');
  });

  it('falls back to a Letter-minus-1in default when the payload declares no @page', () => {
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office">
      <meta name=ProgId content=Word.Document>
      <body><div class=WordSection1><p class=MsoNormal>x</p></div></body></html>`;
    const document = parseWordHtmlString(html, { forceWord: true });
    const { document: file } = renderStandaloneHtml(document);
    expect(file).toContain('max-width: 6.5in');
    expect(file).toContain('padding: 0.6in 0.75in');
  });

  it('lets the caller override the resolved geometry explicitly', () => {
    const document = byId('complex/engagement-report');
    const { document: file } = renderStandaloneHtml(document, {
      contentWidth: '500px',
      pagePadding: '10px',
    });
    expect(file).toContain('max-width: 500px');
    expect(file).toContain('padding: 10px');
    expect(file).not.toContain('max-width: 624px');
  });
});
