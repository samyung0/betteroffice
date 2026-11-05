import { expect, test } from 'bun:test';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { setYrsSelectionInCell, yrsCellLocFromStory, yrsCellStory } from './yrsCommands';

// An edit-born cell carries its replica's client (`~{client}`, #1454).
test.each([
  ['body:t0:r1c2', { story: 'body', tableIndex: 0, row: 1, column: 2 }],
  ['body:t0:r1c2~7', { story: 'body', tableIndex: 0, row: 1, column: 2 }],
  ['body:t0:r0c0~7:t1:r2c3~7', { story: 'body:t0:r0c0~7', tableIndex: 1, row: 2, column: 3 }],
] as const)('%s is a table cell', (story, cell) => {
  expect(yrsCellLocFromStory(story)).toEqual(cell);
});

// A table's `rows` can name a story an Undo removed (Q624): an empty cell.
test('a cell naming a removed story resolves to none and takes no caret', () => {
  const session = {
    hasStory: (story: string) => story === 'body' || story === 'kept',
    paragraphs: (story: string) => {
      if (story !== 'kept') throw new Error(`story ${story} was not found`);
      return [{ paraId: 'p', text: '', properties: {} }];
    },
    setSelection: () => {},
    storySegments: () => [
      {
        kind: 'embed',
        embedKind: 'table',
        payload: {
          grid: [100, 100],
          rows: [{ cells: [{ story: 'kept', tcPr: {} }, { story: 'removed', tcPr: {} }] }],
        },
        attributes: {},
      },
    ],
  } as unknown as YrsSession;
  const cell = (column: number) => ({ story: 'body', tableIndex: 0, row: 0, column });
  expect(yrsCellStory(session, cell(0))).toBe('kept');
  expect(setYrsSelectionInCell(session, cell(0))).toBe(true);
  expect(yrsCellStory(session, cell(1))).toBeNull();
  expect(setYrsSelectionInCell(session, cell(1))).toBe(false);
});
