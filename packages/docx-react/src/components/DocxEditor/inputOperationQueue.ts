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
