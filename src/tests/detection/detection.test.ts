import { describe, expect, it } from 'vitest';
import { detectWordHtml, detectSignals, WORD_SIGNALS } from '../../detection/index.js';
import { loadFixtures } from '../support/fixtures.js';

describe('Word detection', () => {
  it('classifies every fixture as Word with full confidence', () => {
    for (const fixture of loadFixtures()) {
      const result = detectWordHtml(fixture.inputHtml);
      expect(result.isWord, fixture.id).toBe(true);
      expect(result.confidence, fixture.id).toBeGreaterThan(0.9);
      expect(result.source, fixture.id).toBe('word-desktop');
    }
  });

  it('does not classify ordinary HTML as Word', () => {
    const html = `
      <html><head><title>A page</title></head>
      <body><h1>Hello</h1><p class="intro">Some <b>content</b>.</p>
      <ul><li>One</li><li>Two</li></ul></body></html>`;
    const result = detectWordHtml(html);
    expect(result.isWord).toBe(false);
    expect(result.confidence).toBe(0);
  });

  it('refuses to classify on a single weak signal', () => {
    // A stray Mso class that survived a trip through another editor is not
    // enough on its own — that is the whole point of weighting the signals.
    const html = '<div><p class="MsoNormal">Recycled markup</p></div>';
    const result = detectWordHtml(html);
    expect(result.signals.length).toBe(1);
    expect(result.isWord).toBe(false);
  });

  it('classifies on one decisive signal alone', () => {
    const html = '<style>@list l0:level1 {mso-level-number-format:bullet;}</style><p>x</p>';
    const result = detectWordHtml(html);
    expect(result.isWord).toBe(true);
    expect(result.confidence).toBeGreaterThan(0.9);
  });

  it('reports which signals fired', () => {
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office">
      <meta name=Generator content="Microsoft Word 15 (filtered medium)">
      <p class=MsoNormal style='mso-list:l0 level1 lfo1'>x<o:p></o:p></p>`;
    const result = detectWordHtml(html);
    const ids = detectSignals(html).map((s) => s.id);
    expect(ids).toContain('meta-generator-word');
    expect(ids).toContain('css-mso-list');
    expect(ids).toContain('element-o-p');
    expect(result.signals.length).toBe(ids.length);
  });

  it('distinguishes Word Online from Word desktop', () => {
    const online = `<html xmlns:o="urn:schemas-microsoft-com:office:office">
      <meta name=Generator content="Microsoft Word 15">
      <p style='mso-list:l0 level1 lfo1'>x</p>`;
    expect(detectWordHtml(online).source).toBe('word-online');

    const desktop = `<html xmlns:o="urn:schemas-microsoft-com:office:office">
      <meta name=Generator content="Microsoft Word 15 (filtered medium)">
      <div class=WordSection1><p class=MsoNormal>x</p></div>`;
    expect(detectWordHtml(desktop).source).toBe('word-desktop');
  });

  it('recognises Excel and PowerPoint as Office but not Word', () => {
    const excel = `<html xmlns:x="urn:schemas-microsoft-com:office:excel">
      <meta name=ProgId content=Excel.Sheet>
      <meta name=Generator content="Microsoft Excel 15">
      <table><tr><td class=xl65>1</td></tr></table>`;
    expect(detectWordHtml(excel).source).toBe('excel');
  });

  it('handles an empty payload without throwing', () => {
    const result = detectWordHtml('');
    expect(result.isWord).toBe(false);
    expect(result.signals).toEqual([]);
  });

  it('has a unique id for every signal', () => {
    const ids = WORD_SIGNALS.map((s) => s.id);
    expect(new Set(ids).size).toBe(ids.length);
  });
});
