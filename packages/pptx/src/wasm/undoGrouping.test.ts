import { beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type { PresentationHandle } from '../index';
import { initWasm, openPresentation } from '../index';

// Typing goes through replaceText one keystroke at a time (PptxEditor insertSlideText),
// so Undo must take back every keystroke typed inside the 500 ms capture window together.
const root = resolve(import.meta.dir, '../../../..');
let fixture: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
  ]);
  await initWasm(wasm);
  fixture = pptx;
});

const firstStory = (deck: PresentationHandle) =>
  deck.snapshot().slides[0].shapes.find((shape) => shape.textStories.length > 0)!.textStories[0].id;

const type = (deck: PresentationHandle, storyId: string, at: number, text: string) => {
  for (const [index, character] of [...text].entries()) {
    deck.replaceText(storyId, at + index, at + index, character);
  }
  return at + text.length;
};

test('one Undo takes back keystrokes typed inside the capture window', () => {
  const deck = openPresentation(fixture, { clientId: 9991 });
  try {
    const storyId = firstStory(deck);
    const before = deck.story(storyId);
    type(deck, storyId, 0, 'Hello');
    expect(deck.undo().applied).toBe(true);
    expect(deck.story(storyId)).toEqual(before);
    expect(deck.canUndo()).toBe(false);
  } finally {
    deck.dispose();
  }
});

test('a pause longer than the capture window starts a new Undo step', async () => {
  const deck = openPresentation(fixture, { clientId: 9992 });
  try {
    const storyId = firstStory(deck);
    const before = deck.story(storyId);
    const at = type(deck, storyId, 0, 'Hello');
    const typed = deck.story(storyId);
    await Bun.sleep(650);
    type(deck, storyId, at, ' world');
    expect(deck.undo().applied).toBe(true);
    expect(deck.story(storyId)).toEqual(typed);
    expect(deck.undo().applied).toBe(true);
    expect(deck.story(storyId)).toEqual(before);
  } finally {
    deck.dispose();
  }
});
