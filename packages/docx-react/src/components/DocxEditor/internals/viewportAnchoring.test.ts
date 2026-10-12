import { describe, expect, test } from 'bun:test';
import {
  computeViewportAnchoredScrollTop,
  mergeLayoutUpdateOrigin,
  rescaleScrollTop,
  shouldScrollCaretIntoView,
  type ViewportAnchorSnapshot,
} from './viewportAnchoring';

const anchor: ViewportAnchorSnapshot = {
  viewportOffset: 140,
  scrollTopSnapshot: 600,
};

// One page plus its gap, the step a page-count change adds or removes.
const PAGE_STEP = 824;

describe('computeViewportAnchoredScrollTop', () => {
  test('compensates for content height added above the viewport', () => {
    expect(computeViewportAnchoredScrollTop(anchor, 840, 2_000)).toBe(700);
  });

  test('does not move for an edit below the viewport', () => {
    expect(computeViewportAnchoredScrollTop(anchor, 740, 2_000)).toBe(600);
  });

  test('keeps the prior position when an edit spanning the viewport removes the anchor', () => {
    expect(computeViewportAnchoredScrollTop(anchor, null, 2_000)).toBe(600);
  });

  test('compensates a full page added above the viewport', () => {
    // maxScroll grows with the new page, so the target must not be clamped back.
    expect(computeViewportAnchoredScrollTop(anchor, 740 + PAGE_STEP, 2_000 + PAGE_STEP)).toBe(
      600 + PAGE_STEP
    );
  });

  test('compensates a full page removed above the viewport', () => {
    expect(
      computeViewportAnchoredScrollTop(
        { viewportOffset: 140, scrollTopSnapshot: 600 + PAGE_STEP },
        740,
        2_000
      )
    ).toBe(600);
  });

  test('clamps to the shortened document when pages are removed below the viewport', () => {
    expect(computeViewportAnchoredScrollTop(anchor, 740, 500)).toBe(500);
  });
});

describe('shouldScrollCaretIntoView', () => {
  test('preserves local caret scrolling', () => {
    expect(shouldScrollCaretIntoView('local', false)).toBe(true);
  });

  test('does not scroll to a sticky caret after a remote relayout', () => {
    expect(shouldScrollCaretIntoView('remote', false)).toBe(false);
  });

  test('allows a local selection action after a remote relayout', () => {
    expect(shouldScrollCaretIntoView('remote', true)).toBe(true);
  });
});

describe('mergeLayoutUpdateOrigin', () => {
  test('preserves a remote origin through coalescing', () => {
    expect(mergeLayoutUpdateOrigin(null, 'remote')).toBe('remote');
    expect(mergeLayoutUpdateOrigin('remote', 'remote')).toBe('remote');
  });

  test('lets a newer local update supersede a pending remote restore', () => {
    expect(mergeLayoutUpdateOrigin('remote', 'local')).toBe('local');
  });
});

describe('rescaleScrollTop', () => {
  // Three 1000 px pages under a 54 px column top (ruler and padding), 16 px apart.
  const heights = [1000, 1000, 1000];
  test('keeps the point on its page across a zoom change', () => {
    // At 50%, 300 px into page 2 (its top at 54 + 500 + 16 = 570): 60% of the page.
    const at = rescaleScrollTop(570 + 300, 54, heights, 16, 0.5, 1);
    // At 100% page 2's top is 54 + 1000 + 16 = 1070, and 60% of it 600 px down.
    expect(at).toBe(1070 + 600);
    expect(rescaleScrollTop(at, 54, heights, 16, 1, 0.5)).toBe(870);
  });
  test('leaves what does not scale: the column top and the gaps', () => {
    expect(rescaleScrollTop(30, 54, heights, 16, 0.5, 1)).toBe(30);
    // 4 px into the gap under page 1 at 50% stays 4 px into it at 100%.
    expect(rescaleScrollTop(54 + 500 + 4, 54, heights, 16, 0.5, 1)).toBe(54 + 1000 + 4);
  });
});
