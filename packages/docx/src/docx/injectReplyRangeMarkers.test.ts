import { describe, expect, it } from 'bun:test';

import type { BlockContent, Comment, Insertion, Paragraph, Run } from '../types/content';
import { injectReplyRangeMarkers, injectTCReplyRangeMarkers } from './injectReplyRangeMarkers';

const comment = (id: number, parentId?: number): Comment =>
  ({
    id,
    author: 'Reviewer',
    content: [],
    ...(parentId === undefined ? {} : { parentId }),
  } as Comment);
const markers = (blocks: BlockContent[]) =>
  (blocks[0] as Paragraph).content.flatMap((item) =>
    item.type === 'commentRangeStart'
      ? [`S${item.id}`]
      : item.type === 'commentRangeEnd'
      ? [`E${item.id}`]
      : []
  );

describe('reply range markers', () => {
  it("gives a reply its parent's range and leaves a reply with its own markers alone", () => {
    const run: Run = { type: 'run', content: [{ type: 'text', text: 'alpha' }] };
    const blocks: BlockContent[] = [
      {
        type: 'paragraph',
        content: [
          { type: 'commentRangeStart', id: 0 },
          { type: 'commentRangeStart', id: 1 },
          run,
          { type: 'commentRangeEnd', id: 0 },
          { type: 'commentRangeEnd', id: 1 },
        ],
      },
    ];
    injectReplyRangeMarkers(blocks, [comment(0), comment(1, 0), comment(2, 0)]);
    expect(markers(blocks)).toEqual(['S0', 'S2', 'S1', 'E0', 'E2', 'E1']);
  });

  it('leaves a tracked-change reply with its own markers alone', () => {
    const insertion: Insertion = {
      type: 'insertion',
      info: { id: 9, author: 'A' },
      content: [{ type: 'run', content: [{ type: 'text', text: 'new' }] }],
    };
    const blocks: BlockContent[] = [
      {
        type: 'paragraph',
        content: [
          { type: 'commentRangeStart', id: 1 },
          insertion,
          { type: 'commentRangeEnd', id: 1 },
        ],
      },
    ];
    injectTCReplyRangeMarkers(blocks, [comment(1, 9)]);
    expect(markers(blocks)).toEqual(['S1', 'E1']);
  });
});
