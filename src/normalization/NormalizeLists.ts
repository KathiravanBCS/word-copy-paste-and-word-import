import type { WordBlock } from '../model/Block.js';
import type { WordDocument } from '../model/Document.js';
import type { ListItemInfo } from '../model/List.js';
import type { WordNumberFormat } from '../model/ListLevel.js';
import type { DiagnosticCollector } from '../diagnostics/DiagnosticCollector.js';
import { DiagnosticCode } from '../diagnostics/UnsupportedFeature.js';
import { expandLevelText, formatNumber } from '../word/WordListStyleParser.js';

/**
 * List identity, continuation and numbering.
 *
 * The parser knows what each paragraph *says* about itself — its list id, its
 * level, the marker Word drew. It does not know how those paragraphs relate to
 * each other. That is this module's job, and it is where the four cases the
 * spec calls out are distinguished:
 *
 *   same list          consecutive items sharing (listId, lfo)
 *   nested list        a deeper level inside the current run
 *   continuation list  the same (listId, lfo) resuming after an interruption,
 *                      numbering carried forward — Word's own behaviour
 *   restarted list     a new (listId, lfo), or a level re-entered from above
 *
 * Numbering is computed from Word's own level definitions, never from an
 * application scheme. Where Word also told us the literal marker it drew, that
 * text wins on appearance and any disagreement with the computed value is
 * reported rather than quietly resolved — a mismatch means the payload is
 * telling us something the definitions do not.
 */

interface ListRunState {
  instanceId: string;
  /** Counter value per level; `undefined` means "not started at this level". */
  counters: Array<number | undefined>;
  /** True once the run has been interrupted and resumed. */
  continued: boolean;
}

export interface NormalizeListsResult {
  /** Number of distinct list instances found. */
  instances: number;
  /** Number of continuation resumptions. */
  continuations: number;
  /** Number of explicit restarts. */
  restarts: number;
}

export function normalizeLists(
  document: WordDocument,
  diagnostics: DiagnosticCollector,
): NormalizeListsResult {
  const result: NormalizeListsResult = { instances: 0, continuations: 0, restarts: 0 };
  // Runs are keyed by (listId, lfo) and survive across interruptions, which is
  // what makes continuation numbering work.
  const runs = new Map<string, ListRunState>();
  let sequence = 0;

  const nextInstanceId = (key: string): string => `list-${++sequence}-${key.replace(/\|/g, '-')}`;

  const walk = (blocks: WordBlock[], scope: string): void => {
    // `previousKey` tracks contiguity: a non-list block clears it.
    let previousKey: string | null = null;

    for (const block of blocks) {
      if (block.type === 'paragraph' && block.listItem) {
        const item = block.listItem;
        const key = `${scope}|${item.listId}|${item.lfo ?? ''}`;

        let run = runs.get(key);
        if (!run) {
          run = {
            instanceId: nextInstanceId(`${item.listId}${item.lfo ? `-${item.lfo}` : ''}`),
            counters: [],
            continued: false,
          };
          runs.set(key, run);
          result.instances++;
        } else if (previousKey !== key) {
          // The same list resuming after something else came between it: Word
          // continues the numbering, so the counters are deliberately kept.
          run.continued = true;
          result.continuations++;
        }

        applyCounters(item, run, diagnostics);
        item.listInstanceId = run.instanceId;
        previousKey = key;
        continue;
      }

      previousKey = null;

      if (block.type === 'container') {
        walk(block.blocks, `${scope}>c`);
      } else if (block.type === 'table') {
        let cellIndex = 0;
        for (const row of block.rows) {
          for (const cell of row.cells) {
            // Each cell is its own list scope: a list in one cell must not
            // continue the numbering of a list in the cell before it.
            walk(cell.blocks, `${scope}>t${cellIndex++}`);
          }
        }
      }
    }
  };

  walk(document.blocks, 'root');
  countRestarts(document, result);
  return result;
}

/**
 * Advance the counter state for one item and settle its marker text.
 */
function applyCounters(
  item: ListItemInfo,
  run: ListRunState,
  diagnostics: DiagnosticCollector,
): void {
  const level = Math.max(0, item.level);
  const definition = item.levelDefinition;
  const format: WordNumberFormat =
    definition?.numberFormat ?? item.marker.numberFormat ?? 'decimal';

  if (item.marker.type === 'bullet' || item.marker.type === 'none') {
    // Bullets do not count, but entering a bullet level still resets the
    // numbered levels below it, exactly as Word does.
    truncateBelow(run.counters, level);
    return;
  }

  const startAt = definition?.startAt ?? item.marker.startAt ?? 1;
  const current = run.counters[level];

  if (current === undefined) {
    run.counters[level] = startAt;
    // A level entered for the first time, or re-entered after the list came
    // back up a level, restarts at the definition's start value.
    item.restart = true;
    item.startAt = startAt;
  } else {
    run.counters[level] = current + 1;
  }
  truncateBelow(run.counters, level);

  const counters = run.counters.slice(0, level + 1).map((c) => c ?? 1);
  // A composite level text's *own* number format applies to every
  // placeholder it contains, including ones naming an ancestor level — not
  // each ancestor's own independent format. A roman-numbered level 0 ("I.")
  // followed by a decimal-formatted level 1 whose level text is "%1.%2"
  // renders "1.1" in Word, not "I.1": level 1's own `decimal` format governs
  // both digits, because `%1` here does not mean "show the ancestor the way
  // it shows itself", it means "show the ancestor's current count, in the
  // format *this* level declared". So the same format is used for every
  // position in `counters`, not a chain of each level's own — verified
  // directly against a real Word document, while building the TipTap
  // integration's own marker computation (`src/demo/tiptap-editor/
  // WordListNodes.ts`), which has no Word-rendered text to fall back on for
  // a newly typed item and so needed this to be actually correct, not just
  // a diagnostic-only fallback the way it is here (below).
  const formats: WordNumberFormat[] = new Array(counters.length).fill(format);

  const levelText = definition?.levelText ?? item.marker.levelText;
  const computed = levelText
    ? expandLevelText(levelText, counters, formats)
    : formatNumber(counters[level] ?? 1, format);

  if (!item.marker.text) {
    // Word Online and partial payloads omit the rendered marker; the level
    // definition plus the counter state reproduces it exactly.
    item.marker.text = computed;
  } else if (item.marker.text !== computed && computed) {
    // Word's own rendering is the authority on appearance. A disagreement
    // usually means the copy started mid-list, so the counters cannot know the
    // true start value — worth saying out loud, never worth silently "fixing".
    diagnostics.info(
      DiagnosticCode.WORD_LIST_NUMBER_FORMAT_APPROXIMATED,
      `List marker "${item.marker.text}" as rendered by Word differs from "${computed}" computed from the list definition. Word's rendered marker was kept. This is expected when the copied selection begins part-way through a list.`,
      {
        details: {
          rendered: item.marker.text,
          computed,
          listId: item.listId,
          level: level + 1,
        },
        fidelity: 'EXACT',
      },
    );
    // Re-seat the counter on what Word actually drew so the *rest* of the list
    // continues from the right place.
    const reseated = readTrailingNumber(item.marker.text, format);
    if (reseated !== undefined) {
      run.counters[level] = reseated;
      if (current === undefined) item.startAt = reseated;
    }
  }
}

function truncateBelow(counters: Array<number | undefined>, level: number): void {
  for (let i = level + 1; i < counters.length; i++) counters[i] = undefined;
}

/**
 * Read the counter value back out of a rendered marker.
 *
 * `1.4` at level 1 yields 4; `IV.` yields 4; `c)` yields 3. Used only to
 * re-seat a counter when Word's rendered marker disagrees with the computed
 * one, so the remaining items in the list stay consistent with what Word drew.
 */
export function readTrailingNumber(
  text: string,
  format: WordNumberFormat,
): number | undefined {
  const cleaned = text.trim().replace(/^[([{]+/, '').replace(/[.)\]}]+$/, '');
  const last = cleaned.split('.').pop() ?? cleaned;
  if (!last) return undefined;

  switch (format) {
    case 'lower-roman':
    case 'upper-roman': {
      const value = fromRoman(last.toUpperCase());
      return value > 0 ? value : undefined;
    }
    case 'lower-alpha':
    case 'upper-alpha': {
      const value = fromAlpha(last.toUpperCase());
      return value > 0 ? value : undefined;
    }
    default: {
      const value = Number.parseInt(last, 10);
      return Number.isFinite(value) ? value : undefined;
    }
  }
}

const ROMAN_VALUES: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };

function fromRoman(text: string): number {
  if (!/^[IVXLCDM]+$/.test(text)) return 0;
  let total = 0;
  for (let i = 0; i < text.length; i++) {
    const value = ROMAN_VALUES[text[i]!] ?? 0;
    const next = ROMAN_VALUES[text[i + 1] ?? ''] ?? 0;
    total += value < next ? -value : value;
  }
  return total;
}

function fromAlpha(text: string): number {
  if (!/^[A-Z]+$/.test(text)) return 0;
  let total = 0;
  for (const char of text) {
    total = total * 26 + (char.charCodeAt(0) - 64);
  }
  return total;
}

function countRestarts(document: WordDocument, result: NormalizeListsResult): void {
  const walk = (blocks: WordBlock[]): void => {
    for (const block of blocks) {
      if (block.type === 'paragraph' && block.listItem?.restart) result.restarts++;
      else if (block.type === 'container') walk(block.blocks);
      else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
      }
    }
  };
  walk(document.blocks);
}

/* -------------------------------------------------------------------------
 * List tree building (used by the renderer)
 * ---------------------------------------------------------------------- */

export interface ListTreeItem {
  /** The paragraph block this item came from, absent for a synthesised spacer. */
  block?: Extract<WordBlock, { type: 'paragraph' }>;
  item?: ListItemInfo;
  children: ListTreeNode[];
  /**
   * True for an item the engine synthesised to hold a deeper list when Word
   * jumped more than one level at once. It has no content of its own and is
   * rendered as an empty, marker-less `<li>` so the nesting depth survives
   * without inventing a visible item.
   */
  spacer?: boolean;
}

export interface ListTreeNode {
  /** Zero-based nesting level. */
  level: number;
  listInstanceId: string;
  listId: string;
  ordered: boolean;
  /** Explicit start value when this list restarts. */
  startAt?: number;
  items: ListTreeItem[];
}

/**
 * Turn a contiguous run of list paragraphs into a nesting tree.
 *
 * Word's output is flat — every item is a sibling paragraph carrying its own
 * level number — so the tree has to be rebuilt. Levels can jump by more than
 * one (a document can go from level 1 straight to level 3), so intermediate
 * levels are synthesised rather than the jump being flattened, which is what
 * keeps `I. / 1.1 / a. / i. / • / §` structurally valid at every depth.
 */
export function buildListTree(
  blocks: WordBlock[],
  startIndex: number,
): { node: ListTreeNode; nextIndex: number } | null {
  const first = blocks[startIndex];
  if (!first || first.type !== 'paragraph' || !first.listItem) return null;

  const rootLevel = first.listItem.level;
  const root = createNode(first.listItem, rootLevel);
  const stack: ListTreeNode[] = [root];
  let index = startIndex;

  while (index < blocks.length) {
    const block = blocks[index];
    if (!block || block.type !== 'paragraph' || !block.listItem) break;
    const item = block.listItem;
    if (item.listInstanceId !== root.listInstanceId) break;
    if (item.level < rootLevel) break;

    while (stack.length - 1 > item.level - rootLevel) stack.pop();

    while (stack.length - 1 < item.level - rootLevel) {
      // Level jumped by more than one. Create the intermediate list so the
      // nesting depth in the output matches the depth Word declared.
      const parent = stack[stack.length - 1]!;
      const owner = parent.items[parent.items.length - 1];
      const child = createNode(item, parent.level + 1);
      if (owner) owner.children.push(child);
      else parent.items.push({ children: [child], spacer: true });
      stack.push(child);
    }

    const target = stack[stack.length - 1]!;
    if (target.items.length === 0 && item.startAt !== undefined && item.restart) {
      target.startAt = item.startAt;
    }
    target.items.push({ block, item, children: [] });
    index++;
  }

  return { node: root, nextIndex: index };
}

function createNode(item: ListItemInfo, level: number): ListTreeNode {
  const node: ListTreeNode = {
    level,
    listInstanceId: item.listInstanceId ?? item.listId,
    listId: item.listId,
    ordered: item.marker.type === 'number',
    items: [],
  };
  if (item.startAt !== undefined && item.restart) node.startAt = item.startAt;
  return node;
}

