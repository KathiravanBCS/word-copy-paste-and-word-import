# What Word actually puts on the clipboard

Reference for the payload this engine parses. Everything here is observable in
the Clipboard Lab's first pane, and every claim is exercised by a fixture.

## Flavours

A Word copy puts several representations on the clipboard at once:

| flavour | contents |
|---|---|
| `text/html` | CF_HTML — a whole HTML document with fragment markers |
| `text/plain` | text with markers rendered as literal characters |
| `text/rtf` | the RTF representation (not parsed by this engine) |
| image files | on some platforms, the pictures as separate clipboard items |

The engine reads `text/html` as its source, keeps `text/plain` for reference,
and uses the image items to recover pictures the HTML could not resolve.

## The document shell

```html
<html xmlns:v="urn:schemas-microsoft-com:vml"
      xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 15">
<meta name=Originator content="Microsoft Word 15">
<style><!-- … --></style>
</head>
<body lang=EN-US style='tab-interval:.5in'>
<div class=WordSection1>
<!--StartFragment-->
  … the selection …
<!--EndFragment-->
</div>
</body>
</html>
```

Note the unquoted attribute values (`content=Word.Document`). Word emits
HTML 4.0, and any parser reading the payload as XML will fail on it.

`Generator` distinguishes sources: `Microsoft Word 15 (filtered medium)` is the
desktop app's filtered export; the same string without the suffix is typically
Word Online. `div.WordSection1` is page-setup section 1 — a second section
appears as `WordSection2`, and the transition is a section break.

## The `<style>` block

Wrapped in an HTML comment (`<style><!-- … --></style>`), which is a habit from
the 1990s and still emitted. Inside are four things the parser needs:

### `@font-face` — which fonts are symbol fonts

```css
@font-face
	{font-family:Wingdings;
	panose-1:5 0 0 0 0 0 0 0 0 0;
	mso-font-charset:2;
	mso-generic-font-family:decorative;}
```

`mso-font-charset:2` means "symbol font": its bytes are not Unicode code points.
This is how the engine knows a bullet glyph needs decoding rather than passing
through.

### `@list` — the numbering definitions

```css
@list l0
	{mso-list-id:1587389017;
	mso-list-type:hybrid;
	mso-list-template-ids:1263389898 67698689 …;}
@list l0:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:none;
	mso-level-number-position:left;
	text-indent:-.25in;
	font-family:Symbol;}
```

The header rule identifies the list; one `:levelN` rule per level describes it.
`@list` is not standard CSS, so a conforming parser discards all of it — see
[ARCHITECTURE.md](ARCHITECTURE.md#why-the-stylesheet-is-mined-from-raw-text).

| declaration | meaning |
|---|---|
| `mso-level-number-format` | `bullet`, `alpha-lower`, `roman-upper`, `ordinal`, `chicago`, … — **absent means arabic**, not unknown |
| `mso-level-text` | the marker pattern: `\F0B7` (a font byte) or `"%1\.%2"` (a hierarchy) |
| `mso-level-start-at` | the first number; absent means 1 |
| `mso-level-tab-stop` | where the text sits |
| `mso-level-number-position` | marker alignment within its space |
| `text-indent` / `margin-left` | the hanging indent |
| `font-family` | the font the marker is drawn in |

### Style rules

```css
p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
	{mso-style-name:"List Paragraph";
	margin-left:.5in;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
```

`mso-style-name` is the human style name — the only way to know that
`MsoListParagraphCxSpFirst` is "List Paragraph". Word declares one style across
several selectors; the engine registers it under each so a class lookup works.

The `CxSpFirst` / `CxSpMiddle` / `CxSpLast` suffixes mean "first / middle / last
in a run of same-styled paragraphs" — contextual spacing, not a different style.
Word usually emits a rule for each variant, but not always; when it does not,
the engine falls back to the base style, which is the difference between a list
keeping its half-inch indent and losing it.

### `@page`

```css
@page WordSection1 {size:8.5in 11.0in; margin:1.0in 1.0in 1.0in 1.0in;}
div.WordSection1 {page:WordSection1;}
```

Page setup per section. HTML has no equivalent; the engine records it and
diagnoses section breaks as page breaks.

## Lists

The single most important part of the payload. Full treatment in
[LIST-PARSING.md](LIST-PARSING.md); the shape is:

```html
<p class=MsoListParagraphCxSpFirst style='text-indent:-.25in;mso-list:l0 level1 lfo1'>
  <![if !supportLists]><span style='font-family:Symbol'><span style='mso-list:Ignore'>&middot;<span
  style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp; </span></span></span><![endif]>Parent<o:p></o:p></p>
```

- `mso-list:l0 level1 lfo1` — list `l0`, level 1 (one-based), format override `lfo1`
- `<![if !supportLists]>…<![endif]>` — a downlevel-**revealed** conditional
  comment: browsers parse the content as elements, so the marker is live text in
  the DOM
- `<span style='mso-list:Ignore'>` — the marker itself, which Word is telling you
  to ignore
- the innermost small-font span of `&nbsp;` is padding to the tab stop, not
  marker text
- `<o:p></o:p>` — Word's paragraph mark; `<o:p>&nbsp;</o:p>` means "this
  paragraph is deliberately empty"

## Conditional comments

Word uses both kinds and they behave differently:

**Downlevel-hidden** — `<!--[if gte vml 1]> … <![endif]-->`. A real HTML
comment. Browsers see a comment node; the content is never parsed into elements.
Used for VML drawings and Office XML metadata. The engine mines it out of the
raw text.

**Downlevel-revealed** — `<![if !supportLists]> … <![endif]>`. Not a comment at
all — browsers treat the delimiters as bogus comments and parse the content
normally. Used for list markers and the `<img>` twin of a picture.

Getting these the wrong way round is a common failure: the list marker must be
*removed* (it is in the DOM and would otherwise become text), and the VML must be
*scanned* (it is not in the DOM and would otherwise be invisible).

Host parsers vary in how they treat downlevel-hidden comments — happy-dom, for
instance, splits them into a comment plus live elements. The engine handles both
by treating the raw-text scan as authoritative and deferring to the `<img>` twin
where one exists.

## Images

```html
<!--[if gte vml 1]><v:shape id="Picture_x0020_1" o:spid="_x0000_i1025"
 style='width:451.2pt;height:184.8pt'>
 <v:imagedata src="file:///C:/Users/jdoe/AppData/Local/Temp/msohtmlclip1/01/clip_image001.png"
  o:title="Architecture diagram"/>
</v:shape><![endif]--><![if !vml]><img width=602 height=246
src="file:///C:/Users/jdoe/AppData/Local/Temp/msohtmlclip1/01/clip_image001.png"
alt="Architecture diagram" v:shapes="Picture_x0020_1"><![endif]>
```

Every picture appears twice — VML for Word and IE, `<img>` for everyone else,
linked by `v:shapes`. Both point at a temp file that exists on exactly one
machine in the world.

Emitting that `src` produces content that looks right to whoever pasted it
(their browser may even have the file cached) and is broken for everyone else.
That is the most common way a paste-from-Word implementation appears to work and
does not. The engine never emits it: the reference becomes a labelled
placeholder carrying the original path, unless actual bytes were found on the
clipboard.

Data URIs (`<img src="data:image/png;base64,…">`) do occur — from Word Online,
and from some copy paths — and resolve normally.

## Tables

Close to ordinary HTML, with `mso-` hints layered on:

```html
<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0
 style='border-collapse:collapse;mso-padding-alt:0in 5.4pt 0in 5.4pt'>
 <tr style='mso-yfti-irow:-1;mso-yfti-firstrow:yes'>
  <td width=208 valign=top style='width:156.0pt;border:solid windowtext 1.0pt;
  background:#D9E2F3;padding:0in 5.4pt 0in 5.4pt'>
```

- `mso-yfti-irow:-1` marks a repeating header row → a real `<thead>`
- `mso-padding-alt` is the table-wide default cell padding
- widths are stated twice, as a pixel attribute and a point CSS value
- borders are split across the table rule, each cell, and `mso-border-*-alt`
- `windowtext` is a system colour meaning black

Nested tables sit inside a `<td>` and need their own grid resolved
independently.

## Formatting quirks worth knowing

**`mso-bidi-font-weight`.** Word writes every bold run as
`<b style='mso-bidi-font-weight:normal'>Bold</b>`. The `mso-bidi-*` properties
describe the *complex-script* run, not the Latin one. Reading that `normal` as
the font weight cancels the `<b>` and silently un-bolds the entire document.
This engine had that bug; a fixture caught it.

**`mso-spacerun:yes`.** Marks a span whose spaces are significant. Collapsing
them loses deliberate spacing.

**`mso-tab-count:N`.** A tab, written as a run of non-breaking spaces. It must
become N tab runs, not six spaces of text.

**`mso-special-character:line-break`.** A line break; combined with
`page-break-before:always`, a page break.

**Highlight vs. background.** `background:yellow` on a `<span>` is the
highlighter pen. The same declaration on a `<td>` is cell shading — letting it
cascade into the cell's runs puts a highlighter behind every word in the cell.

**`&nbsp;`.** Word uses it as real content. It must not be collapsed as
whitespace.

## Word Online, Outlook and the rest

Word Online omits the marker spans and often emits real `<ol>`/`<ul>`. Outlook
uses the same engine as Word desktop but adds its own quoting wrappers. Excel
sends `xl` classes and a table with no document structure. The engine detects
which it is dealing with (`document.detection.source`) and handles the
differences; HTML lists are converted to the same `listItem` model as `mso-list`
paragraphs, with `fromHtmlList: true` recording where they came from, so
downstream code has exactly one representation of a list to deal with.
