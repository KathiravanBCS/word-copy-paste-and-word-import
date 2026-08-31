import type {
  DiagnosticCodeValue,
  DiagnosticLocation,
  DiagnosticSeverity,
  FidelityClass,
  WordDiagnostic,
} from './UnsupportedFeature.js';

export interface AddDiagnosticInput {
  code: DiagnosticCodeValue | string;
  message: string;
  severity?: DiagnosticSeverity;
  fidelity?: FidelityClass;
  location?: DiagnosticLocation;
  details?: Record<string, string | number | boolean>;
}

/**
 * Accumulates diagnostics during a parse.
 *
 * Identical code+message pairs are folded together with a `count` so that a
 * 400-page paste with 900 unresolved images produces one readable line, not
 * 900. The collector is deliberately cheap: parsing a large document must not
 * be dominated by diagnostic bookkeeping.
 */
export class DiagnosticCollector {
  private readonly items: WordDiagnostic[] = [];
  private readonly index = new Map<string, WordDiagnostic>();
  private readonly maxItems: number;
  private overflowed = false;

  constructor(maxItems = 2000) {
    this.maxItems = maxItems;
  }

  add(input: AddDiagnosticInput): void {
    const severity = input.severity ?? 'warning';
    const fidelity = input.fidelity ?? severityToFidelity(severity);
    const key = `${input.code} ${input.message}`;
    const existing = this.index.get(key);
    if (existing) {
      existing.count = (existing.count ?? 1) + 1;
      return;
    }
    if (this.items.length >= this.maxItems) {
      if (!this.overflowed) {
        this.overflowed = true;
        this.items.push({
          code: 'DIAGNOSTIC_OVERFLOW',
          severity: 'info',
          fidelity: 'EQUIVALENT',
          message: `Diagnostic limit of ${this.maxItems} reached; further diagnostics were suppressed.`,
        });
      }
      return;
    }
    const diagnostic: WordDiagnostic = {
      code: input.code,
      severity,
      fidelity,
      message: input.message,
      count: 1,
    };
    if (input.location) diagnostic.location = input.location;
    if (input.details) diagnostic.details = input.details;
    this.items.push(diagnostic);
    this.index.set(key, diagnostic);
  }

  info(
    code: DiagnosticCodeValue | string,
    message: string,
    rest: Partial<AddDiagnosticInput> = {},
  ): void {
    this.add({ ...rest, code, message, severity: 'info', fidelity: rest.fidelity ?? 'EQUIVALENT' });
  }

  warn(
    code: DiagnosticCodeValue | string,
    message: string,
    rest: Partial<AddDiagnosticInput> = {},
  ): void {
    this.add({
      ...rest,
      code,
      message,
      severity: 'warning',
      fidelity: rest.fidelity ?? 'APPROXIMATED',
    });
  }

  error(
    code: DiagnosticCodeValue | string,
    message: string,
    rest: Partial<AddDiagnosticInput> = {},
  ): void {
    this.add({
      ...rest,
      code,
      message,
      severity: 'error',
      fidelity: rest.fidelity ?? 'UNSUPPORTED',
    });
  }

  has(code: string): boolean {
    return this.items.some((d) => d.code === code);
  }

  all(): WordDiagnostic[] {
    return this.items;
  }

  /** Merge another collector's items into this one (used by sub-parsers). */
  merge(other: DiagnosticCollector): void {
    for (const item of other.all()) {
      const key = `${item.code} ${item.message}`;
      const existing = this.index.get(key);
      if (existing) {
        existing.count = (existing.count ?? 1) + (item.count ?? 1);
      } else {
        this.items.push(item);
        this.index.set(key, item);
      }
    }
  }
}

function severityToFidelity(severity: DiagnosticSeverity): FidelityClass {
  switch (severity) {
    case 'error':
      return 'UNSUPPORTED';
    case 'warning':
      return 'APPROXIMATED';
    default:
      return 'EQUIVALENT';
  }
}
