import { describe, expect, it } from 'bun:test';
import type { ShapeSnapshot, StorySnapshot, TextStyleSnapshot } from '@betteroffice/pptx';
import { shapeClipboard, storyClipboard } from './clipboard';

type Run = [text: string, style?: Partial<TextStyleSnapshot>];

function story(id: string, ...paragraphs: Run[][]): StorySnapshot {
  const plain = { bold: null, italic: null, fontSizePt: null, color: null, fontFamily: null, underline: null };
  return {
    id,
    length: paragraphs.reduce((length, runs) => length + runs.reduce((sum, [text]) => sum + text.length, 0) + 1, 0),
    paragraphs: paragraphs.map((runs, index) => ({
      id: `${id}:${index}`,
      alignment: null,
      level: 0,
      bulletJson: null,
      runs: runs.map(([text, style]) => ({ text, style: { ...plain, ...style } })),
    })),
  };
}

function shape(id: string, textStories: StorySnapshot[], extra: Partial<ShapeSnapshot> = {}): ShapeSnapshot {
  return {
    id, sourceId: 0, kind: 'shape', name: id, x: 0, y: 0, width: 1, height: 1, rotationDeg: 0,
    flipH: false, flipV: false, geometry: 'rect', adjustValues: {}, placeholder: null, fill: null,
    resolvedFillColor: null, outline: null, resolvedOutlineColor: null, mediaPartPath: null,
    graphic: null, textStories, children: [], ...extra,
  };
}

describe('storyClipboard', () => {
  const formatted = story(
    's',
    [['Hello '], ['bold', { bold: true }]],
    [['<i&u>', { italic: true, underline: 'sng' }], [' tail', { underline: 'none' }]]
  );

  it('copies a range across runs and paragraphs as text and escaped HTML', () => {
    expect(storyClipboard(formatted, 6, 16)).toEqual({
      text: 'bold\n<i&u>',
      html: '<p><b>bold</b></p><p><i><u>&lt;i&amp;u&gt;</u></i></p>',
    });
  });

  it('copies the whole story without a trailing paragraph break', () => {
    expect(storyClipboard(formatted, 0, formatted.length)).toEqual({
      text: 'Hello bold\n<i&u> tail',
      html: '<p>Hello <b>bold</b></p><p><i><u>&lt;i&amp;u&gt;</u></i> tail</p>',
    });
  });

  it('keeps soft line breaks and the empty paragraph a range ends in', () => {
    const lines = story('l', [['one\ntwo']], [['three']]);
    expect(storyClipboard(lines, 0, 7)).toEqual({ text: 'one\ntwo', html: '<p>one<br>two</p>' });
    expect(storyClipboard(lines, 4, 8)).toEqual({ text: 'two\n', html: '<p>two</p><p></p>' });
  });
});

describe('shapeClipboard', () => {
  it('copies stories in tree order, one per line, without hidden or blank ones', () => {
    const group = shape('group', [], {
      kind: 'group',
      children: [
        shape('title', [story('t', [['Title']])]),
        shape('hidden', [story('h', [['Hidden']])], { hidden: true }),
        shape('table', [story('c1', [['A1']]), story('c2', [[' ']]), story('c3', [['B1', { bold: true }]], [['B1 second']])], {
          kind: 'graphicFrame',
        }),
      ],
    });
    expect(shapeClipboard(group)).toEqual({
      text: 'Title\nA1\nB1\nB1 second',
      html: '<p>Title</p><p>A1</p><p><b>B1</b></p><p>B1 second</p>',
    });
  });
});
