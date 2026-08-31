import type { ClipboardPayload } from '../clipboard/ClipboardPayload.js';
import { detectSignals, type DetectedSignal } from './WordSignalDetector.js';

/** Which Office application (or lookalike) produced the payload. */
export type WordSource =
  | 'word-desktop'
  | 'word-online'
  | 'outlook'
  | 'excel'
  | 'powerpoint'
  | 'office-generic'
  | 'unknown';

export interface WordDetectionResult {
  isWord: boolean;
  /** 0..1. Reaches 1.0 quickly once a decisive signal is present. */
  confidence: number;
  /** Human-readable signal descriptions, for the lab UI and diagnostics. */
  signals: string[];
  /** The full signal records, for programmatic use. */
  detail: DetectedSignal[];
  source: WordSource;
  /** Word major version parsed from the Generator meta, when present. */
  wordVersion?: number;
}

export interface DetectionOptions {
  /**
   * Confidence at which a payload is treated as Word. Default 0.5, which one
   * decisive signal or two strong signals will clear, but a lone `MsoNormal`
   * class will not.
   */
  threshold?: number;
  /**
   * Minimum number of distinct signals required regardless of weight.
   * Default 2 — deliberately refusing to classify on a single weak signal.
   * A decisive signal overrides this (see below).
   */
  minSignals?: number;
}

const DEFAULT_THRESHOLD = 0.5;
const DEFAULT_MIN_SIGNALS = 2;

/**
 * Classify a raw HTML string as Word clipboard content.
 *
 * Confidence is the saturating sum of matched signal weights. A decisive
 * signal (ProgId meta, `@list`, `<w:WordDocument>`, `<![if !supportLists]>`)
 * is on its own sufficient: those strings do not occur outside Office output.
 */
export function detectWordHtml(html: string, options: DetectionOptions = {}): WordDetectionResult {
  const threshold = options.threshold ?? DEFAULT_THRESHOLD;
  const minSignals = options.minSignals ?? DEFAULT_MIN_SIGNALS;

  if (!html || html.length === 0) {
    return { isWord: false, confidence: 0, signals: [], detail: [], source: 'unknown' };
  }

  const detail = detectSignals(html);
  const hasDecisive = detail.some((s) => s.strength === 'decisive');

  let confidence = 0;
  for (const signal of detail) {
    // Saturating accumulation: each signal closes part of the remaining gap to
    // 1.0, so many weak signals can corroborate without any single one lying.
    confidence += (1 - confidence) * signal.weight;
  }
  confidence = Math.min(1, Number(confidence.toFixed(4)));

  const isWord = hasDecisive || (confidence >= threshold && detail.length >= minSignals);

  const result: WordDetectionResult = {
    isWord,
    confidence,
    signals: detail.map((s) => s.description),
    detail,
    source: identifySource(html, detail),
  };
  const version = parseWordVersion(html);
  if (version !== undefined) result.wordVersion = version;
  return result;
}

/** Detect against a captured clipboard payload. */
export function detectWordPayload(
  payload: ClipboardPayload,
  options: DetectionOptions = {},
): WordDetectionResult {
  return detectWordHtml(payload.html ?? '', options);
}

function identifySource(html: string, detail: DetectedSignal[]): WordSource {
  const progId = /<meta[^>]+name=["']?ProgId["']?[^>]*content=["']?([\w.]+)/i.exec(html);
  const generator = /<meta[^>]+name=["']?Generator["']?[^>]*content=["']?([^"'>]+)/i.exec(html);
  const gen = generator?.[1] ?? '';

  if (/Microsoft\s+Word/i.test(gen)) {
    // Word Online writes a distinctly leaner payload without the filtered-medium
    // marker and without the WordSection wrapper.
    if (/filtered\s+medium/i.test(gen)) return 'word-desktop';
    return /WordSection\d/i.test(html) ? 'word-desktop' : 'word-online';
  }
  if (/Microsoft\s+Excel/i.test(gen)) return 'excel';
  if (/Microsoft\s+PowerPoint/i.test(gen)) return 'powerpoint';
  if (/Microsoft\s+Outlook/i.test(gen) || /class=["']?MsoPlainText/i.test(html)) return 'outlook';

  const prog = progId?.[1] ?? '';
  if (/^Word\./i.test(prog)) return 'word-desktop';
  if (/^Excel\./i.test(prog)) return 'excel';
  if (/^PowerPoint\./i.test(prog)) return 'powerpoint';

  if (detail.some((s) => s.id === 'css-at-list' || s.id === 'element-word-document-xml')) {
    return 'word-desktop';
  }
  if (detail.length > 0) return 'office-generic';
  return 'unknown';
}

function parseWordVersion(html: string): number | undefined {
  const match = /<meta[^>]+name=["']?Generator["']?[^>]*content=["']?Microsoft\s+Word\s+(\d+)/i.exec(
    html,
  );
  if (!match) return undefined;
  const version = Number.parseInt(match[1] ?? '', 10);
  return Number.isFinite(version) ? version : undefined;
}
