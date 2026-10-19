/**
 * Thrown when an edit's collaboration update could exceed the engine's cap
 * (64 MiB, one action's most): a clear, a paste or another batch of cell
 * inputs. The workbook did not change.
 */
export class EditTooLargeError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'EditTooLargeError';
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

/** Thrown by an edit whose cell text exceeds Excel's 32,767-character limit; nothing changed. */
export class CellTextTooLongError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'CellTextTooLongError';
  }
}

/** The engine refusal an error message names, as its typed error, else null. */
export function refusalOf(message: string): Error | null {
  if (message.startsWith('edit too large')) return new EditTooLargeError(message);
  if (message === "there's a problem with this formula") return new UnparsedFormulaError(message);
  if (message === "cell text exceeds Excel's length limit") return new CellTextTooLongError(message);
  if (/^range \d+x\d+ exceeds the \d+-cell copy cap$/.test(message))
    return new RangeTooLargeError(message);
  return null;
}
