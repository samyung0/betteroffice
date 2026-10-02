import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { initWasm } from '@betteroffice/pptx';
import type { PptxCommandState } from './commands';
import type { PptxEditorApi } from './PptxEditor';
import { PptxEditor } from './PptxEditor';

const root = resolve(import.meta.dir, '../../..');
const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

let fixture: Uint8Array;
let fonts: { family: string; bytes: Uint8Array }[];

beforeAll(async () => {
  const [wasm, pptx, font] = await Promise.all([
    readFile(resolve(root, 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  fixture = pptx;
  fonts = [{ family: 'Liberation Sans', bytes: font }];
});

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

async function open(readOnly = false) {
  const opened: PptxEditorApi[] = [];
  const states: PptxCommandState[] = [];
  const errors: Error[] = [];
  const view = render(
    <PptxEditor
      file={fixture}
      fonts={fonts}
      readOnly={readOnly}
      onReady={(api) => opened.push(api)}
      onCommandState={(state) => states.push(state)}
      onError={(error) => errors.push(error)}
    />
  );
  await waitFor(() => expect(opened.length).toBe(1), { timeout: 15_000 });
  const api = opened[0];
  const slideIds = () => api.handle.snapshot().slides.map((slide) => slide.id);
  const run = (id: Parameters<PptxEditorApi['runCommand']>[0], value?: string) => {
    let ran = false;
    act(() => {
      ran = api.runCommand(id, value);
    });
    return ran;
  };
  return { api, errors, run, slideIds, state: () => states[states.length - 1], view };
}

describe('PptxEditor host commands', () => {
  it('deletes and moves the current slide, which the strip follows', async () => {
    const { api, errors, run, slideIds, state } = await open();
    const ids = slideIds();
    expect(ids.length).toBeGreaterThan(2);
    expect(state().enabled['slide.moveUp']).toBe(false);
    expect(run('slide.moveUp')).toBe(false);

    expect(run('slide.moveDown')).toBe(true);
    expect(slideIds().slice(0, 2)).toEqual([ids[1], ids[0]]);
    expect(state().enabled['slide.moveUp']).toBe(true);

    expect(run('slide.moveToEnd')).toBe(true);
    expect(slideIds()[ids.length - 1]).toBe(ids[0]);
    expect(state().enabled['slide.moveDown']).toBe(false);

    expect(run('slide.moveToStart')).toBe(true);
    expect(slideIds()[0]).toBe(ids[0]);

    expect(run('slide.delete')).toBe(true);
    expect(slideIds()).toEqual(ids.slice(1));
    // The next slide becomes current, so another delete removes it.
    expect(run('slide.delete')).toBe(true);
    expect(slideIds()).toEqual(ids.slice(2));
    expect(run('edit.undo')).toBe(true);
    expect(slideIds()).toEqual(ids.slice(1));
    expect(api.handle.save().length).toBeGreaterThan(0);
    expect(errors).toEqual([]);
  }, 60_000);

  it('deletes the selected object only', async () => {
    const originalImage = globalThis.Image;
    class FakeImage {
      onload: (() => void) | null = null;
      naturalWidth = 400;
      naturalHeight = 200;
      set src(_value: string) {
        queueMicrotask(() => this.onload?.());
      }
    }
    globalThis.Image = FakeImage as unknown as typeof Image;
    try {
      const { api, errors, run, state, view } = await open();
      const shapes = () => api.handle.snapshot().slides[0].shapes.length;
      const before = shapes();
      expect(state().enabled['edit.delete']).toBe(false);
      expect(run('edit.delete')).toBe(false);

      const file = new File([Uint8Array.from([0x89, 0x50, 0x4e, 0x47])], 'logo.png', {
        type: 'image/png',
      });
      await act(async () => {
        fireEvent.change(view.getByTestId('pptx-insert-image-input'), {
          target: { files: [file] },
        });
        await new Promise((done) => setTimeout(done, 0));
      });
      await waitFor(() => expect(shapes()).toBe(before + 1));
      await waitFor(() => expect(state().enabled['edit.delete']).toBe(true));
      expect(state().enabled['arrange.sendToBack']).toBe(true);

      expect(run('edit.delete')).toBe(true);
      expect(shapes()).toBe(before);
      expect(state().enabled['edit.delete']).toBe(false);
      expect(errors).toEqual([]);
    } finally {
      globalThis.Image = originalImage;
    }
  }, 60_000);

  it('validates values and refuses edits when read-only', async () => {
    const { run, slideIds, state } = await open(true);
    const ids = slideIds();
    expect(state().enabled['view.present']).toBe(true);
    expect(run('view.zoom', '1.5')).toBe(true);
    expect(state().zoom).toBe('1.5');
    expect(run('view.zoom', 'huge')).toBe(false);
    expect(run('slide.delete')).toBe(false);
    expect(run('slide.new')).toBe(false);
    expect(run('insert.shape', 'ellipse')).toBe(false);
    expect(slideIds()).toEqual(ids);
  }, 60_000);
});
