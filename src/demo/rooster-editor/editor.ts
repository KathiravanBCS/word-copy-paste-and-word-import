import { createEditor } from 'roosterjs';
import { WordClipboardEnginePlugin } from './WordClipboardEnginePlugin.js';

/**
 * The RoosterJS integration demo.
 *
 * `createEditor` is RoosterJS's own, completely unmodified. The only thing
 * this file adds beyond what any RoosterJS consumer would write is passing
 * `WordClipboardEnginePlugin` in as an additional plugin — that one line is
 * the entire integration surface.
 */

const host = document.getElementById('rooster-editor') as HTMLDivElement;
const editor = createEditor(host, [new WordClipboardEnginePlugin()]);

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
 * This dispatches a genuine `paste` ClipboardEvent carrying a `text/html`
 * flavour, the same event a real Ctrl+V produces. It goes through the exact
 * code path a real paste does — RoosterJS's own paste listener,
 * `BeforePasteEvent`, and `WordClipboardEnginePlugin` — nothing here is a
 * shortcut around the integration being demonstrated.
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

  host.focus();
  const dataTransfer = new DataTransfer();
  dataTransfer.setData('text/html', fixture.html);
  dataTransfer.setData('text/plain', stripTags(fixture.html));

  const event = new ClipboardEvent('paste', {
    clipboardData: dataTransfer,
    bubbles: true,
    cancelable: true,
  });
  host.dispatchEvent(event);
  setStatus(`pasted fixture: ${fixture.id}`);
});

function stripTags(html: string): string {
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

document.getElementById('clear-editor')!.addEventListener('click', () => {
  // formatContentModel is RoosterJS's own documented way to mutate content;
  // emptying `blocks` and returning true (meaning "this changed something")
  // keeps RoosterJS's internal Content Model and the visible DOM in sync,
  // which a direct innerHTML clear would not.
  editor.formatContentModel((model) => {
    model.blocks = [];
    return true;
  });
  setStatus('editor cleared');
});

// Exposed for manual poking from the console, and for automated verification
// of the integration without relying on a real OS clipboard.
Object.assign(window as unknown as Record<string, unknown>, { rooster: { editor } });
