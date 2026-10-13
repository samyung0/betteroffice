import { describe, expect, it } from 'bun:test';

import { StaleProposalError } from '../proposals';
import { EquationTextError, toError } from './errors';

describe('WASM errors', () => {
  it('turns coded refusals into their classes and keeps engine text', () => {
    const equation = toError('{"code":"equationText"}');
    expect(equation).toBeInstanceOf(EquationTextError);
    expect(equation.message).toBe('An equation can only be changed in PowerPoint.');
    expect(toError('{"code":"staleProposal","targets":["a"]}')).toBeInstanceOf(StaleProposalError);
    const engine = toError('story "s" was not found');
    expect(engine).not.toBeInstanceOf(EquationTextError);
    expect(engine.message).toBe('story "s" was not found');
  });
});
