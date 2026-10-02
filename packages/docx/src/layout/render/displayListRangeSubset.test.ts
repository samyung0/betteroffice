/**
 * A superseded query facade (no session handle) answers body range queries
 * through the JSON-arg path from only the pages the range can touch. The
 * answer must equal the whole-list query against the real layout wasm.
 */

import { beforeAll, describe, expect, test } from 'bun:test';

import { preloadLayoutWasm } from '../../wasm/layout';
import { createDisplayListQueries } from './displayListQueries';
import type { DisplayList, DisplayPage } from './displayList';
import {
  loadRustDisplayListQueryEngine,
  type RustDisplayListQueryEngine,
} from './rustDisplayList';

// Page n holds a 10-character run at 20n+1..20n+11 and a blank-line marker at
// 20n+12, so ranges cross runs, markers, page ends and the gaps between pages.
function page(pageIndex: number): DisplayPage {
  const start = pageIndex * 20 + 1;
  const run = (docStart: number, docEnd: number, text: string, lineIndex: number) => ({
    kind: 'text',
    text,
    x: 100,
    baselineY: 100 + lineIndex * 20,
    width: text.length * 8,
    font: '400 16px Calibri',
    color: '#000000',
    docStart,
    docEnd,
    blockId: pageIndex * 2 + lineIndex,
    lineIndex: 0,
  });
  return {
    pageIndex,
    width: 500,
    height: 500,
    primitives: [run(start, start + 10, 'abcdefghij', 0), run(start + 11, start + 11, '', 1)],
  } as unknown as DisplayPage;
}

let engine: RustDisplayListQueryEngine;
const list: DisplayList = { pages: Array.from({ length: 40 }, (_, index) => page(index)) };
const whole = JSON.stringify(list);

beforeAll(async () => {
  await preloadLayoutWasm();
  engine = await loadRustDisplayListQueryEngine();
});

describe('superseded range queries', () => {
  test('match the whole-list query and serialize only the touched pages', () => {
    const inputs: number[] = [];
    const recording: RustDisplayListQueryEngine = {
      ...engine,
      rangeRectsJson: (displayList, from, to) => {
        inputs.push(displayList.length);
        return engine.rangeRectsJson(displayList, from, to);
      },
    };
    const first = createDisplayListQueries(list, recording);
    first.rangeRects(1, 2);
    // the successor adopts the handle; `first` keeps answering without one
    createDisplayListQueries({ pages: [...list.pages] }, recording, first).rangeRects(1, 2);

    const ranges: Array<[number, number]> = [];
    for (let from = 0; from < 820; from += 7) ranges.push([from, from + 1], [from, from + 33]);
    ranges.push([12, 12], [500, 480], [0, 900]);
    for (const [from, to] of ranges) {
      expect(first.rangeRects(from, to)).toEqual(
        JSON.parse(engine.rangeRectsJson(whole, from, to))
      );
    }
    expect(inputs.length).toBe(ranges.length);
    // a caret-sized range ships a page or two, not the 40-page list
    expect(Math.max(...inputs.slice(0, 2))).toBeLessThan(whole.length / 10);
  });
});
