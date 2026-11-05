import { describe, expect, test } from 'bun:test';
import {
  createYrsInputPositionMap,
  displayPositionToYrsLoc,
  type YrsSession,
  type YrsStorySegment,
} from '@betteroffice/docx/yrs';
import { createYrsPositionProjection, YrsPositionProjection } from './yrsPositionProjection';

const segments: Record<string, YrsStorySegment[]> = {
  body: [
    { kind: 'text', text: 'before', attributes: {} },
    { kind: 'pilcrow', paraId: 'body:p0', properties: {}, attributes: {} },
    {
      kind: 'embed',
      embedKind: 'table',
      payload: {
        grid: [100],
        rows: [{ cells: [{ story: 'body:t0:r0c0', tcPr: {} }] }],
      },
      attributes: {},
    },
    { kind: 'text', text: 'after', attributes: {} },
    { kind: 'pilcrow', paraId: 'body:p1', properties: {}, attributes: {} },
  ],
  'body:t0:r0c0': [
    { kind: 'text', text: 'cell', attributes: {} },
    { kind: 'pilcrow', paraId: 'cell:p0', properties: {}, attributes: {} },
  ],
};

const session = {
  docEpoch: () => 0,
  hasStory: (story: string) => story in segments,
  storySegments: (story: string) => segments[story] ?? [],
} as unknown as YrsSession;

describe('YrsPositionProjection', () => {
  test('returns null without reading a missing story', () => {
    let readStory = false;
    const missingStorySession = {
      hasStory: (story: string) => story === 'body',
      storySegments: () => {
        readStory = true;
        throw new Error('missing story');
      },
    } as unknown as YrsSession;
    let projection: YrsPositionProjection | null = null;

    expect(() => {
      projection = createYrsPositionProjection(missingStorySession, 'fn:2');
    }).not.toThrow();
    expect(projection).toBeNull();
    expect(readStory).toBe(false);
  });

  test('a cell naming a removed story takes no positions and no caret, as the engine lowers it', () => {
    const removed: Record<string, YrsStorySegment[]> = {
      body: [
        {
          kind: 'embed',
          embedKind: 'table',
          payload: {
            grid: [100, 100],
            rows: [
              { cells: [{ story: 'removed', tcPr: {} }, { story: 'cell', tcPr: {} }] },
            ],
          },
          attributes: {},
        },
        { kind: 'text', text: 'after', attributes: {} },
        { kind: 'pilcrow', paraId: 'body:p0', properties: {}, attributes: {} },
      ],
      cell: [{ kind: 'pilcrow', paraId: 'cell:p0', properties: {}, attributes: {} }],
    };
    const projection = new YrsPositionProjection(
      {
        docEpoch: () => 0,
        hasStory: (story: string) => story in removed,
        storySegments: (story: string) => {
          if (!(story in removed)) throw new Error(`story ${story} was not found`);
          return removed[story];
        },
      } as unknown as YrsSession,
      'body'
    );

    const table = projection.tableAtStart(0)!;
    expect(table.cells.map(({ story, start, nodeSize }) => ({ story, start, nodeSize }))).toEqual([
      { story: 'removed', start: 2, nodeSize: 2 },
      { story: 'cell', start: 4, nodeSize: 4 },
    ]);
    // A position in the empty cell belongs to the story around the table.
    expect(projection.targetAt(3).story).toBe('body');
    expect(projection.targetAt(5).story).toBe('cell');
  });

  test('maps post-table positions back to the root story input map', () => {
    const projection = new YrsPositionProjection(session, 'body');
    // `paragraph_spans` measures a paragraph from the previous pilcrow, so the
    // table embed ahead of `body:p1` counts as one of its six units.
    const map = createYrsInputPositionMap('body', [
      { paraId: 'body:p0', length: 6 },
      { paraId: 'body:p1', length: 6 },
    ]);
    const target = projection.targetAt(23);
    const loc = displayPositionToYrsLoc(map, target.displayPosition);

    expect(target).toEqual({ story: 'body', displayPosition: 12 });
    expect(loc).toEqual({
      story: 'body',
      paraId: 'body:p1',
      offset: 3,
    });
    expect(projection.positionForLoc(loc!)).toBe(23);
  });

  test('keeps table cell positions scoped to the cell input map', () => {
    const projection = new YrsPositionProjection(session, 'body');
    const map = createYrsInputPositionMap('body:t0:r0c0', [
      { paraId: 'cell:p0', length: 4 },
    ]);
    const target = projection.targetAt(12);

    expect(target).toMatchObject({
      story: 'body:t0:r0c0',
      displayPosition: 1,
    });
    const loc = displayPositionToYrsLoc(map, target.displayPosition);
    expect(loc).toEqual({
      story: 'body:t0:r0c0',
      paraId: 'cell:p0',
      offset: 0,
    });
    expect(projection.positionForLoc(loc!)).toBe(12);
  });
});
