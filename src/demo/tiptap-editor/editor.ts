import { Editor } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { TableKit } from '@tiptap/extension-table';
import { Image } from '@tiptap/extension-image';
import { WordList, WordListItem } from './WordListNodes.js';
import { WordClipboardExtension } from './WordClipboardExtension.js';

/**
 * The TipTap integration demo.
 *
 * `StarterKit` supplies everything ordinary (paragraphs, headings, marks,
 * undo) unmodified; its own `bulletList`/`orderedList`/`listItem` are turned
 * off because `WordList`/`WordListItem` (`WordListNodes.ts`) replace them —
 * two schemas for the same job would conflict over which one owns `<li>`.
 * `Table`/`Image` are TipTap's own, also unmodified. `WordClipboardExtension`
 * is the only piece this project wrote.
 */

const host = document.getElementById('tiptap-editor') as HTMLDivElement;

const editor = new Editor({
  element: host,
  extensions: [
    StarterKit.configure({
      bulletList: false,
      orderedList: false,
      listItem: false,
      listKeymap: false,
    }),
    WordList,
    WordListItem,
    TableKit.configure({ table: { resizable: false } }),
    Image,
    WordClipboardExtension,
  ],
  content: '<p></p>',
});

const statusEl = document.getElementById('status')!;
const footerStatusEl = document.getElementById('footer-status')!;
const fixturePicker = document.getElementById('fixture-picker') as HTMLSelectElement;

function setStatus(message: string): void {
  statusEl.textContent = message;
  footerStatusEl.textContent = message;
}

setStatus('editor ready');

/* -------------------------------------------------------------------------
 * Fixture picker — simulates a real Word paste for anyone without Word open.
 *
 * Dispatches a genuine `paste` ClipboardEvent carrying a `text/html`
 * flavour, the same event a real Ctrl+V produces, through the exact code
 * path a real paste does — ProseMirror's own `handlePaste` prop, which
 * `WordClipboardExtension` implements.
 * ---------------------------------------------------------------------- */

const fixtures = import.meta.glob('../../fixtures/word/**/input.html', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>;

const fixtureList = Object.entries(fixtures)
  .map(([path, html]) => ({
    id: path.replace(/^.*fixtures\/word\//, '').replace(/\/input\.html$/, ''),
    html,
  }))
  .sort((a, b) => a.id.localeCompare(b.id));

for (const fixture of fixtureList) {
  const option = document.createElement('option');
  option.value = fixture.id;
  option.textContent = fixture.id;
  fixturePicker.append(option);
}

fixturePicker.addEventListener('change', () => {
  const fixture = fixtureList.find((f) => f.id === fixturePicker.value);
  fixturePicker.value = '';
  if (!fixture) return;

  editor.commands.focus('end');
  const dataTransfer = new DataTransfer();
  dataTransfer.setData('text/html', fixture.html);
  dataTransfer.setData('text/plain', stripTags(fixture.html));

  const event = new ClipboardEvent('paste', {
    clipboardData: dataTransfer,
    bubbles: true,
    cancelable: true,
  });
  editor.view.dom.dispatchEvent(event);
  setStatus(`pasted fixture: ${fixture.id}`);
});

function stripTags(html: string): string {
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

document.getElementById('clear-editor')!.addEventListener('click', () => {
  editor.commands.clearContent(true);
  setStatus('editor cleared');
});

// Exposed for manual poking from the console, and for automated verification
// of the integration without relying on a real OS clipboard.
Object.assign(window as unknown as Record<string, unknown>, { tiptap: { editor } });
