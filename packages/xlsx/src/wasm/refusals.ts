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

/** The engine refusal an error message names, as its typed error, else null. */
export function refusalOf(message: string): Error | null {
  if (message.startsWith('edit too large')) return new EditTooLargeError(message);
  if (/^range \d+x\d+ exceeds the \d+-cell copy cap$/.test(message))
    return new RangeTooLargeError(message);
  return null;
}
