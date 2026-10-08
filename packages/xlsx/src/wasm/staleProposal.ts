/**
 * Thrown by {@link WorkbookHandle.acceptProposal} when the workbook changed
 * under a proposal since it was staged (an edit touched one of its base cells)
 * and `force` was not set. `cells` are the a1 addresses that moved, so the UI
 * can name them and offer a force-apply.
 */
export class StaleProposalError extends Error {
  readonly cells: string[];
  constructor(cells: string[]) {
    super(`stale: ${cells.join(', ')}`);
    this.name = 'StaleProposalError';
    this.cells = cells;
  }
}
