import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import {
  applyFrameDeltaOwned,
  decodeFrameDelta,
  type RetainedFrame,
} from '@betteroffice/docx/layout/render';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { CanvasPagesView } from './CanvasPagesView';

const { act, cleanup, render } = await import('@testing-library/react');

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'))
    )
  )
);
afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const request = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
  renderEnv: {},
});

/** Every page mirror's document positions, page by page. */
function mirrorPositions(container: HTMLElement): string[][] {
  return [...container.querySelectorAll('.canvas-page')].map((page) =>
    [...page.querySelectorAll<HTMLElement>('[data-doc-start], [data-doc-end]')].map(
      (el) => `${el.dataset.docStart ?? ''}-${el.dataset.docEnd ?? ''}`
    )
  );
}

const settle = () => act(() => new Promise((resolve) => setTimeout(resolve, 400)));

// An edit before later pages only shifts their positions, which the owned
// apply does in place (same page object): their mirrors must still carry the
// positions a fresh layout gives them (RenderedDomContext reads them).
for (const edit of ['peer', 'local'] as const) {
  test(`after a ${edit} insert before later pages, every mirror holds a fresh layout's positions`, async () => {
    const native = createEditSession(9310);
    native.load_json(
      JSON.stringify([
        {
          storyId: 'body',
          paragraphs: Array.from({ length: 150 }, (_, index) => ({ text: `Paragraph ${index} of the body` })),
        },
      ])
    );
    try {
      let frame: RetainedFrame = applyFrameDeltaOwned(
        null,
        decodeFrameDelta(native.relayout_frame(request, '{}', 0))
      );
      const shown = render(<CanvasPagesView displayList={frame.displayList} frame={frame} glyphOutlineProvider={false} />);
      await settle();
      const later = frame.pages[1]!.page;

      const [opening] = JSON.parse(native.paragraphs('body')) as Array<{ paraId: string }>;
      let delta: Uint8Array;
      if (edit === 'peer') {
        native.insert_text('body', opening!.paraId, 0, 'x', undefined, undefined);
        delta = native.relayout_frame(request, '{}', frame.frameEpoch);
      } else {
        native.set_selection('body', opening!.paraId, 0, opening!.paraId, 0);
        delta = native.apply_input('x', frame.frameEpoch);
      }
      frame = applyFrameDeltaOwned(frame, decodeFrameDelta(delta));
      expect(frame.pages[1]!.page).toBe(later);
      shown.rerender(<CanvasPagesView displayList={frame.displayList} frame={frame} glyphOutlineProvider={false} />);
      await settle();

      const fresh = createEditSession(9311);
      try {
        fresh.load(native.encode_state());
        const full = applyFrameDeltaOwned(null, decodeFrameDelta(fresh.relayout_frame(request, '{}', 0)));
        const expected = render(<CanvasPagesView displayList={full.displayList} frame={full} glyphOutlineProvider={false} />);
        await settle();
        const want = mirrorPositions(expected.container);
        expect(want[1]!.length).toBeGreaterThan(0);
        expect(mirrorPositions(shown.container)).toEqual(want);
      } finally {
        fresh.free();
      }
    } finally {
      native.free();
    }
  });
}
