import { describe, expect, it } from 'vitest';
import { parseWordHtml } from '../word/WordHtmlParser.js';

const WORD_BULLETS = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">
<head><meta name=ProgId content=Word.Document><meta name=Generator content="Microsoft Word 15 (filtered medium)">
<style><!--
 @font-face {font-family:Symbol; panose-1:5 0 0 0 0 0 0 0 0 0; mso-font-charset:2;}
 @font-face {font-family:Wingdings; panose-1:5 0 0 0 0 0 0 0 0 0; mso-font-charset:2;}
 p.MsoNormal, li.MsoNormal, div.MsoNormal {mso-style-parent:""; margin:0in; font-size:11.0pt; font-family:"Calibri",sans-serif;}
 p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph {mso-style-name:"List Paragraph"; margin-left:.5in;}
 @list l0 {mso-list-id:1587389017; mso-list-type:hybrid;}
 @list l0:level1 {mso-level-number-format:bullet; mso-level-text:\\F0B7; mso-level-tab-stop:.5in; mso-level-number-position:left; text-indent:-.25in; font-family:Symbol;}
 @list l0:level2 {mso-level-number-format:bullet; mso-level-text:o; mso-level-tab-stop:1.0in; mso-level-number-position:left; text-indent:-.25in; font-family:"Courier New";}
 @list l0:level3 {mso-level-number-format:bullet; mso-level-text:\\F0A7; mso-level-tab-stop:1.5in; mso-level-number-position:left; text-indent:-.25in; font-family:Wingdings;}
-->
</style></head>
<body lang=EN-US>
<div class=WordSection1>
<!--StartFragment-->
<p class=MsoListParagraphCxSpFirst style='text-indent:-.25in;mso-list:l0 level1 lfo1'><![if !supportLists]><span style='font-family:Symbol;mso-fareast-font-family:Symbol'><span style='mso-list:Ignore'>&middot;<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]>Parent<o:p></o:p></p>
<p class=MsoListParagraphCxSpMiddle style='margin-left:1.0in;text-indent:-.25in;mso-list:l0 level2 lfo1'><![if !supportLists]><span style='font-family:"Courier New"'><span style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]>Child<o:p></o:p></p>
<p class=MsoListParagraphCxSpLast style='margin-left:1.5in;text-indent:-.25in;mso-list:l0 level3 lfo1'><![if !supportLists]><span style='font-family:Wingdings'><span style='mso-list:Ignore'>&sect;<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]>Grandchild<o:p></o:p></p>
<!--EndFragment-->
</div>
</body></html>`;

describe('smoke: bullet levels', () => {
  it('lifts bullet glyphs out of text and maps symbol fonts', () => {
    const { document } = parseWordHtml(WORD_BULLETS);

    expect(document.detection.isWord).toBe(true);
    expect(document.lists).toHaveLength(1);
    expect(document.lists[0]!.levels).toHaveLength(3);

    const paragraphs = document.blocks.filter((b) => b.type === 'paragraph');
    expect(paragraphs).toHaveLength(3);

    const texts = paragraphs.map((p) =>
      p.type === 'paragraph' ? p.runs.map((r) => (r.type === 'text' ? r.text : '')).join('') : '',
    );
    expect(texts).toEqual(['Parent', 'Child', 'Grandchild']);

    const markers = paragraphs.map((p) => (p.type === 'paragraph' ? p.listItem?.marker : undefined));
    expect(markers[0]?.glyph).toBe('•');
    expect(markers[1]?.glyph).toBe('o');
    expect(markers[2]?.glyph).toBe('▪');

    const levels = paragraphs.map((p) => (p.type === 'paragraph' ? p.listItem?.level : undefined));
    expect(levels).toEqual([0, 1, 2]);
  });
});
