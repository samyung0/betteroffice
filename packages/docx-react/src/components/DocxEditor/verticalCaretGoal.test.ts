import { describe, expect, test } from 'bun:test';
import type { DisplayListVisualLine } from '@betteroffice/docx/layout/render';
import { pageMoveTarget, paragraphVerticalMove, VerticalCaretGoal } from './verticalCaretGoal';

describe('VerticalCaretGoal', () => {
  test('retains goal X across consecutive vertical moves', () => {
    const goal = new VerticalCaretGoal();
    expect(goal.current()).toBeUndefined();
    goal.retain(172.5);
    expect(goal.current()).toBe(172.5);
    goal.retain(172.5);
    expect(goal.current()).toBe(172.5);
  });

  test('resets after a non-vertical caret action', () => {
    const goal = new VerticalCaretGoal();
    goal.retain(84);
    goal.reset();
    expect(goal.current()).toBeUndefined();
  });
});

describe('paragraphVerticalMove', () => {
  const paragraphs = [
    { paraId: 'first', length: 8 },
    { paraId: 'second', length: 3 },
  ];

  test('keeps header and footer navigation available without a display-list facade', () => {
    expect(
      paragraphVerticalMove(
        paragraphs,
        { story: 'hf:rIdHeader', paraId: 'first', offset: 6 },
        'down'
      )
    ).toEqual({ story: 'hf:rIdHeader', paraId: 'second', offset: 3 });
    expect(
      paragraphVerticalMove(
        paragraphs,
        { story: 'hf:rIdFooter', paraId: 'second', offset: 2 },
        'up'
      )
    ).toEqual({ story: 'hf:rIdFooter', paraId: 'first', offset: 2 });
  });

  // The facade answers body positions only, so a note navigates the same way a
  // band does — and the caret never leaves the note it is in.
  test('keeps footnote and endnote navigation inside their own story', () => {
    expect(
      paragraphVerticalMove(paragraphs, { story: 'fn:2', paraId: 'first', offset: 6 }, 'down')
    ).toEqual({ story: 'fn:2', paraId: 'second', offset: 3 });
    expect(
      paragraphVerticalMove(paragraphs, { story: 'en:2', paraId: 'second', offset: 2 }, 'up')
    ).toEqual({ story: 'en:2', paraId: 'first', offset: 2 });
  });

  test('the ends of a note story hold the caret', () => {
    const top = { story: 'fn:2', paraId: 'first', offset: 4 };
    const bottom = { story: 'fn:2', paraId: 'second', offset: 1 };
    expect(paragraphVerticalMove(paragraphs, top, 'up')).toEqual(top);
    expect(paragraphVerticalMove(paragraphs, bottom, 'down')).toEqual(bottom);
  });
});

describe('pageMoveTarget', () => {
  // Two 100 x 200 pages of ten 20px lines, 20px apart on screen. Line i of page p holds
  // positions p * 1000 + i * 10 to that + 9, one per 1px from x = 5.
  const lines: DisplayListVisualLine[] = [0, 1].flatMap((pageIndex) =>
    Array.from({ length: 10 }, (_, index) => ({
      pageIndex,
      x: 5,
      y: index * 20,
      width: 10,
      height: 20,
      baseline: index * 20 + 15,
      from: pageIndex * 1000 + index * 10,
      to: pageIndex * 1000 + index * 10 + 9,
    }))
  );
  const lineAt = (pos: number) => lines.find((line) => line.from <= pos && pos <= line.to) ?? null;
  const queries = {
    visualLines: () => lines,
    visualLineAtPosition: lineAt,
    caretRect: (pos: number) => {
      const line = lineAt(pos);
      return (
        line && {
          pageIndex: line.pageIndex,
          x: 5 + pos - line.from,
          y: line.y,
          width: 0,
          height: 20,
        }
      );
    },
    hitTestRegions: (pageIndex: number, x: number, y: number) => {
      const line = lines.find((l) => l.pageIndex === pageIndex && l.y <= y && y < l.y + l.height);
      return line
        ? { region: 'body' as const, pos: line.from + Math.min(9, Math.max(0, Math.round(x - 5))) }
        : null;
    },
    pageSize: () => ({ width: 100, height: 200 }),
  };
  const clientY = (pageIndex: number, y: number) => pageIndex * 220 + y;

  test('PageDown lands one viewport below at the caret x, across a page, and the view follows', () => {
    expect(pageMoveTarget(queries, clientY, 23, 'down', 200, undefined)).toEqual({
      position: 1013,
      goalX: 8,
      scrollBy: 200,
    });
  });

  test('PageUp lands one viewport above at the retained goal x', () => {
    expect(pageMoveTarget(queries, clientY, 1013, 'up', 200, 6)).toEqual({
      position: 21,
      goalX: 6,
      scrollBy: -200,
    });
  });

  test('near the end it takes the last line, and on it the line end; on the first line, its start', () => {
    expect(pageMoveTarget(queries, clientY, 1083, 'down', 200, undefined)).toMatchObject({
      position: 1093,
      scrollBy: 20,
    });
    expect(pageMoveTarget(queries, clientY, 1093, 'down', 200, undefined)).toMatchObject({
      position: 1099,
      scrollBy: 0,
    });
    expect(pageMoveTarget(queries, clientY, 4, 'up', 200, undefined)).toMatchObject({
      position: 0,
      scrollBy: 0,
    });
  });

  test('a caret off the body lines or on a page without geometry stays put', () => {
    expect(pageMoveTarget(queries, clientY, 5000, 'down', 200, undefined)).toBeNull();
    expect(pageMoveTarget(queries, () => null, 23, 'down', 200, undefined)).toBeNull();
  });
});
