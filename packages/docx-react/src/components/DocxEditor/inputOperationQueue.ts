/** What must land in a session before it is destroyed: input behind a held press. */
const heldInput = new WeakMap<object, () => Promise<void>>();

export function setHeldInput(session: object, drain: (() => Promise<void>) | null): void {
  if (drain) heldInput.set(session, drain);
  else heldInput.delete(session);
}

/** Lands a session's held input; null when it holds none. */
export function drainHeldInput(session: object): Promise<void> | null {
  const drain = heldInput.get(session);
  heldInput.delete(session);
  return drain ? drain() : null;
}

export class InputOperationQueue {
  private pending: Promise<void> = Promise.resolve();
  private interactionEpoch = 0;
  private failure: { error: unknown } | null = null;
  private pendingCount = 0;

  constructor(
    private readonly reportError: (error: unknown) => void,
    private readonly reportPending?: (pending: boolean) => void,
  ) {}

  enqueue(operation: () => void | Promise<void>): void {
    this.pendingCount += 1;
    if (this.pendingCount === 1) this.reportPending?.(true);
    const pending = this.pending.then(operation, operation);
    this.pending = pending.catch((error) => {
      this.failure ??= { error };
      this.reportError(error);
    }).finally(() => {
      this.pendingCount -= 1;
      if (!this.pendingCount && !this.failure) this.reportPending?.(false);
    });
  }

  /**
   * Later operations wait until `released` settles, which must not reject. A hold is
   * no input: it reports nothing pending, though a flush waits for it.
   */
  hold(released: Promise<void>): void {
    this.pending = this.pending.then(() => released);
  }

  hasPending(): boolean {
    return this.pendingCount > 0;
  }

  idle(): Promise<void> {
    return this.pending;
  }

  /** Waits for accepted operations and rejects if this queue has lost input. */
  flush(): Promise<void> {
    return this.pending.then(() => {
      if (this.failure) throw this.failure.error;
    });
  }

  captureInteractionEpoch(): number {
    return this.interactionEpoch;
  }

  advanceInteractionEpoch(): void {
    this.interactionEpoch += 1;
  }

  isInteractionEpochCurrent(epoch: number): boolean {
    return this.interactionEpoch === epoch;
  }
}
