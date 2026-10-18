/**
 * Thrown by {@link WorkbookHandle.clearContents} when the clear's collaboration
 * update could exceed the limit it was given; the workbook did not change.
 */
export class ClearTooLargeError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ClearTooLargeError';
  }
}

/**
 * Thrown when a range read (`rangeCells`, a copy) or a range format spans more
 * cells than the engine takes at once; nothing changed.
 */
export class RangeTooLargeError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'RangeTooLargeError';
  }
}

/**
 * Thrown by a typed {@link WorkbookHandle.editCell} whose formula does not
 * parse, even with its open parentheses closed; nothing changed.
 */
export class UnparsedFormulaError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'UnparsedFormulaError';
  }
}

/** The engine refusal an error message names, as its typed error, else null. */
export function refusalOf(message: string): Error | null {
  if (message.startsWith('clear too large')) return new ClearTooLargeError(message);
  if (message === "there's a problem with this formula") return new UnparsedFormulaError(message);
  if (/^range \d+x\d+ exceeds the \d+-cell copy cap$/.test(message))
    return new RangeTooLargeError(message);
  return null;
}
