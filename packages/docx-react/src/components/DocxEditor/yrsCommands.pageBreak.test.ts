import { expect, test } from 'bun:test';
import { pageBreakOffered } from './yrsCommands';

// As in Word: the body and its content controls, never a cell, header, footer or note.
test.each([
  ['body', true],
  ['body:sdt0', true],
  ['body:t0:r0c0', false],
  ['body:sdt0:t1:r2c3:sdt0', false],
  ['hf:rId20', false],
  ['fn:1', false],
  ['en:2', false],
] as const)('a page break in %s is offered: %s', (story, offered) => {
  expect(pageBreakOffered(story)).toBe(offered);
});
