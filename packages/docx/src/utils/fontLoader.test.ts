import { describe, expect, it } from 'bun:test';

import { extractFontsFromDocument } from './fontLoader';

const run = (text: string, hint?: string) => ({
  type: 'run',
  formatting: { fontFamily: { ascii: 'Calibri', eastAsia: 'ＭＳ 明朝', ...(hint ? { hint } : {}) } },
  content: [{ type: 'text', text }],
});

const document = (...runs: unknown[]) => ({
  package: {
    document: { content: [{ type: 'paragraph', content: runs }] },
    styles: { docDefaults: { rPr: { fontFamily: { ascii: 'Cambria', eastAsia: 'ＭＳ ゴシック' } } } },
    fontTable: {
      fonts: [
        { name: 'Calibri', panose1: '020F0502020204030204', charset: '00' },
        { name: 'ＭＳ 明朝', altName: 'MS Mincho', charset: '80' },
      ],
    },
  },
});

describe('extractFontsFromDocument', () => {
  it('leaves East Asian faces out of a document whose text uses none', () => {
    expect([...extractFontsFromDocument(document(run('Lesson plan')))].sort()).toEqual([
      'Calibri',
      'Cambria',
    ]);
  });

  it('loads them once a run holds East Asian characters', () => {
    expect([...extractFontsFromDocument(document(run('Plan'), run('授業計画')))].sort()).toEqual([
      'Calibri',
      'Cambria',
      'MS Mincho',
      'ＭＳ ゴシック',
      'ＭＳ 明朝',
    ]);
  });

  it('loads them for other non-ASCII text under the eastAsia hint', () => {
    expect(extractFontsFromDocument(document(run('“Plan”', 'eastAsia')))).toContain('ＭＳ 明朝');
  });
});
