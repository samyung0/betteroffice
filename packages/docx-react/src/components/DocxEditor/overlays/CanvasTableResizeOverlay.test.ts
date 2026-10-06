import { expect, test } from 'bun:test';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { drawnColumnsTwips } from './CanvasTableResizeOverlay';

const layout = (columnWidths?: number[]) =>
  ({
    pages: [
      { fragments: [{ kind: 'paragraph', pmStart: 0 }] },
      { fragments: [{ kind: 'table', pmStart: 12, columnWidths }] },
    ],
  }) as unknown as Layout;

test('a drag starts from the columns the layout drew, not the stored grid', () => {
  // An auto-fitted table draws 23 / 38 / 105px over a 3120-twip grid.
  expect(drawnColumnsTwips(layout([23, 38.2, 105]), 12, 3)).toEqual([345, 573, 1575]);
  // Another table, a column count that no longer matches, or no layout: the grid.
  expect(drawnColumnsTwips(layout([23, 38, 105]), 40, 3)).toBeNull();
  expect(drawnColumnsTwips(layout([23, 38]), 12, 3)).toBeNull();
  expect(drawnColumnsTwips(layout(), 12, 3)).toBeNull();
  expect(drawnColumnsTwips(null, 12, 3)).toBeNull();
});
