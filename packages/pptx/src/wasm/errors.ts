import { StaleProposalError } from '../proposals';

/** A text edit on an equation, whose math PowerPoint keeps apart from the picture the editor shows. */
export class EquationTextError extends Error {
  constructor() {
    super('An equation can only be changed in PowerPoint.');
    this.name = 'EquationTextError';
  }
}

/** The error a WASM call threw: a coded refusal as its class, engine text as an Error. */
export function toError(error: unknown): Error {
  if (error instanceof Error) return error;
  if (typeof error === 'string') {
    try {
      const parsed = JSON.parse(error);
      if (parsed.code === 'staleProposal' && Array.isArray(parsed.targets)
          && parsed.targets.every((target: unknown) => typeof target === 'string')) {
        return new StaleProposalError(parsed.targets);
      }
      if (parsed.code === 'equationText') return new EquationTextError();
    } catch {}
  }
  return new Error(typeof error === 'string' ? error : String(error));
}
