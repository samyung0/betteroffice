import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, it, spyOn } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import * as pptx from '@betteroffice/pptx';
import type { StorySnapshot } from '@betteroffice/pptx';
import type { PptxEditorApi } from './PptxEditor';
import { PptxEditor } from './PptxEditor';
import { pointerTargetAtPoint } from './interactions';

const root = resolve(import.meta.dir, '../../..');
// The demo deck's first slide title: "Office files," then "without the office.".
const STORY = 'story:slide:0:256:shape:2:0';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, createEvent, fireEvent, render, waitFor } = await import('@testing-library/react');

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

/** Opens the demo deck with the caret after "Office" in the first slide's title. */
async function editTitle() {
  const opened: PptxEditorApi[] = [];
  const errors: Error[] = [];
  const view = render(<PptxEditor file={fixture} fonts={fonts} onReady={(api) => opened.push(api)} onError={(error) => errors.push(error)} />);
  await waitFor(() => expect(opened.length).toBe(1), { timeout: 15_000 });
  const api = opened[0];
  const shape = api.handle.snapshot().slides[0].shapes.find((item) => item.textStories.some((story) => story.id === STORY))!;
  const input = view.getByTestId('pptx-text-input') as HTMLTextAreaElement;
  const caret = 'Office'.length;
  await act(async () => {
    expect(api.selectText({ slide: 1, shapeId: shape.id, storyId: STORY, start: caret, end: caret })).toBe(true);
  });
  return { api, view, input, errors };
}

it('one Ctrl+Z takes back a word typed in one burst', async () => {
  const { api, input, errors } = await editTitle();
  const before = api.handle.story(STORY);

  for (const key of ' new') fireEvent.input(input, { target: { value: key } });
  expect(plain(api.handle.story(STORY))).toBe('Office new files,\nwithout the office.');

  fireEvent.keyDown(input, { key: 'z', ctrlKey: true });
  expect(plain(api.handle.story(STORY))).toBe(plain(before));
  expect(api.handle.story(STORY)).toEqual(before);
  expect(errors).toEqual([]);
}, 60_000);

it('a click or an arrow key between keystrokes starts a new undo step', async () => {
  const { api, view, input, errors } = await editTitle();
  const before = api.handle.story(STORY);
  // The clock stands still, so only the caret moves can split the typing.
  const clock = spyOn(Date, 'now').mockReturnValue(Date.now());
  try {
    for (const key of ' new') fireEvent.input(input, { target: { value: key } });
    const typed = api.handle.story(STORY);
    const text = plain(typed);

    // Click into the second line, "without the office.".
    const frame = api.handle.layoutSlide(0);
    let point: { x: number; y: number; position: number } | undefined;
    for (let y = 0; y < frame.height && !point; y += 4) {
      for (let x = 0; x < frame.width; x += 4) {
        const hit = api.handle.hitTest(x, y);
        if (hit?.kind === 'text' && hit.storyId === STORY && hit.position > text.indexOf('\n') + 1 &&
          pointerTargetAtPoint(frame, { x, y }, 1) === 'text') { point = { x, y, position: hit.position }; break; }
      }
    }
    expect(point).toBeDefined();
    const { x, y, position } = point!;
    const canvas = view.getByTestId('pptx-slide-canvas');
    canvas.getBoundingClientRect = () => new DOMRect(0, 0, frame.width, frame.height);
    for (const [type, timeStamp] of [['pointerDown', 1_000], ['pointerUp', 1_050]] as const) {
      const event = createEvent[type](canvas, { pointerId: 1, isPrimary: true, button: 0, clientX: x, clientY: y });
      Object.defineProperty(event, 'timeStamp', { value: timeStamp });
      fireEvent(canvas, event);
    }
    fireEvent.input(input, { target: { value: 'x' } });
    const clicked = api.handle.story(STORY);
    expect(plain(clicked)).toBe(`${text.slice(0, position)}x${text.slice(position)}`);

    fireEvent.keyDown(input, { key: 'ArrowLeft' });
    fireEvent.input(input, { target: { value: 'y' } });
    expect(plain(api.handle.story(STORY))).toBe(`${text.slice(0, position)}yx${text.slice(position)}`);

    fireEvent.keyDown(input, { key: 'z', ctrlKey: true });
    expect(api.handle.story(STORY)).toEqual(clicked);
    fireEvent.keyDown(input, { key: 'z', ctrlKey: true });
    expect(api.handle.story(STORY)).toEqual(typed);
    fireEvent.keyDown(input, { key: 'z', ctrlKey: true });
    expect(api.handle.story(STORY)).toEqual(before);
  } finally {
    clock.mockRestore();
  }
  expect(errors).toEqual([]);
}, 60_000);

it('a paste between keystrokes is its own undo step', async () => {
  const { api, input, errors } = await editTitle();
  const before = api.handle.story(STORY);
  // The clock stands still, so only the paste can split the typing.
  const clock = spyOn(Date, 'now').mockReturnValue(Date.now());
  try {
    for (const key of ' a') fireEvent.input(input, { target: { value: key } });
    const typed = api.handle.story(STORY);
    fireEvent.paste(input, { clipboardData: { getData: () => 'XY' } });
    const pasted = api.handle.story(STORY);
    fireEvent.input(input, { target: { value: 'b' } });
    expect(plain(api.handle.story(STORY))).toBe('Office aXYb files,\nwithout the office.');

    fireEvent.keyDown(input, { key: 'z', ctrlKey: true });
    expect(api.handle.story(STORY)).toEqual(pasted);
    fireEvent.keyDown(input, { key: 'z', ctrlKey: true });
    expect(api.handle.story(STORY)).toEqual(typed);
    fireEvent.keyDown(input, { key: 'z', ctrlKey: true });
    expect(api.handle.story(STORY)).toEqual(before);
  } finally {
    clock.mockRestore();
  }
  expect(errors).toEqual([]);
}, 60_000);
