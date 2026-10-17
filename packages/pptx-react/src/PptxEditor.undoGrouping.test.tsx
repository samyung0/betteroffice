import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, it } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import * as pptx from '@betteroffice/pptx';
import type { StorySnapshot } from '@betteroffice/pptx';
import type { PptxEditorApi } from './PptxEditor';
import { PptxEditor } from './PptxEditor';

const root = resolve(import.meta.dir, '../../..');
// The demo deck's first slide title: "Office files," then "without the office.".
const STORY = 'story:slide:0:256:shape:2:0';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

let fixture: Uint8Array;
let fonts: pptx.PptxFontFace[];

beforeAll(async () => {
  const [wasm, deck, font] = await Promise.all([
    readFile(resolve(root, 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await pptx.initWasm(wasm);
  fixture = deck;
  fonts = [{ family: 'Liberation Sans', bytes: font }];
});

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

const plain = (story: StorySnapshot) =>
  story.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n');

it('one Ctrl+Z takes back a word typed in one burst', async () => {
  const opened: PptxEditorApi[] = [];
  const errors: Error[] = [];
  const view = render(<PptxEditor file={fixture} fonts={fonts} onReady={(api) => opened.push(api)} onError={(error) => errors.push(error)} />);
  await waitFor(() => expect(opened.length).toBe(1), { timeout: 15_000 });
  const api = opened[0];
  const shape = api.handle.snapshot().slides[0].shapes.find((item) => item.textStories.some((story) => story.id === STORY))!;
  const before = api.handle.story(STORY);
  const input = view.getByTestId('pptx-text-input') as HTMLTextAreaElement;
  const caret = 'Office'.length;
  await act(async () => {
    expect(api.selectText({ slide: 1, shapeId: shape.id, storyId: STORY, start: caret, end: caret })).toBe(true);
  });

  for (const key of ' new') fireEvent.input(input, { target: { value: key } });
  expect(plain(api.handle.story(STORY))).toBe('Office new files,\nwithout the office.');

  fireEvent.keyDown(input, { key: 'z', ctrlKey: true });
  expect(plain(api.handle.story(STORY))).toBe(plain(before));
  expect(api.handle.story(STORY)).toEqual(before);
  expect(errors).toEqual([]);
}, 60_000);
