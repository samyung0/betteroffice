import { expect, test } from 'bun:test';
import { yrsCellLocFromStory } from './yrsCommands';

// An edit-born cell carries its replica's client (`~{client}`, #1454).
test.each([
  ['body:t0:r1c2', { story: 'body', tableIndex: 0, row: 1, column: 2 }],
  ['body:t0:r1c2~7', { story: 'body', tableIndex: 0, row: 1, column: 2 }],
  ['body:t0:r0c0~7:t1:r2c3~7', { story: 'body:t0:r0c0~7', tableIndex: 1, row: 2, column: 3 }],
] as const)('%s is a table cell', (story, cell) => {
  expect(yrsCellLocFromStory(story)).toEqual(cell);
});
