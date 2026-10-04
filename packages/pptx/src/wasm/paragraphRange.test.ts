import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type { PresentationHandle, StorySnapshot } from '../index';
import { initWasm, openPresentation } from '../index';

const root = resolve(import.meta.dir, '../../../..');
// The demo deck's first slide title: "Office files," then "without the office.".
const STORY = 'story:slide:0:256:shape:2:0';
const FIRST_END = 'Office files,'.length;
// "fice files," + the paragraph break + "with".
const START = 2;
const END = FIRST_END + 5;
let fixture: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
  ]);
  await initWasm(wasm);
  fixture = pptx;
});

const plain = (story: StorySnapshot) =>
  story.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n');

/** The deck with the second paragraph centred, so the two paragraphs differ. */
function open(): { handle: PresentationHandle; before: StorySnapshot } {
  const handle = openPresentation(fixture);
  expect(plain(handle.story(STORY))).toBe('Office files,\nwithout the office.');
  handle.setParagraphAlignment(STORY, FIRST_END + 1, FIRST_END + 1, 'ctr');
  const before = handle.story(STORY);
  expect(before.paragraphs.map((paragraph) => paragraph.alignment)).toEqual(['l', 'ctr']);
  return { handle, before };
}

describe('text ranges across paragraphs', () => {
  const cases: Array<[string, (handle: PresentationHandle) => void, string]> = [
    ['delete', (handle) => {
      expect(() => handle.deleteText(STORY, START, END)).toThrow('crosses a paragraph boundary');
      handle.replaceText(STORY, START, END, '');
    }, 'Ofout the office.'],
    ['type a character', (handle) => { handle.replaceText(STORY, START, END, 'x'); }, 'Ofxout the office.'],
    ['paste one line', (handle) => { handle.replaceText(STORY, START, END, 'pasted '); }, 'Ofpasted out the office.'],
    ['paste lines', (handle) => {
      expect(handle.replaceText(STORY, START, END, 'one\ntwo ').end).toBe(START + 'one\ntwo '.length);
    }, 'Ofone\ntwo out the office.'],
    ['Enter', (handle) => { handle.replaceText(STORY, START, END, '\n'); }, 'Of\nout the office.'],
  ];

  for (const [name, edit, expected] of cases) {
    test(`${name} joins the paragraphs under the first one's properties, undoes in one step and saves`, () => {
      const { handle, before } = open();
      try {
        edit(handle);
        const after = handle.story(STORY);
        expect(plain(after)).toBe(expected);
        // A typed newline splits the joined paragraph; its first half keeps the id.
        expect(after.paragraphs[0].id).toBe(before.paragraphs[0].id);
        expect(after.paragraphs.map((paragraph) => paragraph.alignment)).toEqual(after.paragraphs.map(() => 'l'));

        expect(handle.undo().applied).toBe(true);
        expect(handle.story(STORY)).toEqual(before);
        expect(handle.redo().applied).toBe(true);
        expect(handle.story(STORY)).toEqual(after);
        // Reopening renumbers paragraph ids, so compare what they carry.
        const reopened = openPresentation(handle.save());
        try {
          const saved = reopened.story(STORY);
          expect(plain(saved)).toBe(expected);
          expect(saved.paragraphs.map((paragraph) => paragraph.alignment)).toEqual(saved.paragraphs.map(() => 'l'));
        } finally {
          reopened.dispose();
        }
      } finally {
        handle.dispose();
      }
    });
  }

  test('a refused edit changes nothing and leaves saving working', () => {
    const { handle, before } = open();
    try {
      expect(() => handle.replaceText(STORY, START, END, 'one\ntwo\u0001')).toThrow('cannot be stored');
      expect(() => handle.replaceText(STORY, 0, before.length, '')).toThrow('outside length');
      expect(handle.story(STORY)).toEqual(before);
      const saved = openPresentation(handle.save());
      try {
        expect(saved.story(STORY)).toEqual(before);
      } finally {
        saved.dispose();
      }
      handle.replaceText(STORY, START, END, 'x');
      expect(plain(handle.story(STORY))).toBe('Ofxout the office.');
      expect(handle.save().byteLength).toBeGreaterThan(0);
    } finally {
      handle.dispose();
    }
  });
});
