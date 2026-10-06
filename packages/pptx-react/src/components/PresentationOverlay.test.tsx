import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, it, spyOn } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { initWasm, openPresentation, type SlideDisplayList } from '@betteroffice/pptx/viewer';
import { Window as HappyWindow } from 'happy-dom';
import { NotesWindow } from '../notesWindow';
import { PresentationOverlay, type PresentationSource } from './PresentationOverlay';
import { DRAWN_ICON_NAMES, IconSetContext, type IconSet } from './ui/ToolbarIcon';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import(
  '@testing-library/react'
);
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

const frame = (assetId: string): SlideDisplayList => ({
  contractVersion: 1,
  width: 320,
  height: 180,
  primitives: [
    {
      kind: 'image',
      name: 'Test image',
      assetId,
      objectId: 1,
      x: 0,
      y: 0,
      w: 320,
      h: 180,
    },
  ],
});

it('only presents the latest slide when an earlier image resolves late', async () => {
  const images = [{}, {}] as CanvasImageSource[];
  let finishFirst: (image: CanvasImageSource) => void = () => {};
  const pending = new Promise<CanvasImageSource>((resolve) => {
    finishFirst = resolve;
  });
  const draws = new Map<HTMLCanvasElement, unknown[][]>();
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
    function (this: HTMLCanvasElement) {
      const calls: unknown[][] = [];
      draws.set(this, calls);
      return new Proxy(
        {},
        {
          get: (_target, key) =>
            key === 'drawImage' ? (...args: unknown[]) => calls.push(args) : () => {},
        }
      ) as CanvasRenderingContext2D;
    } as unknown as typeof HTMLCanvasElement.prototype.getContext
  );
  const handle = {
    layoutSlide: (index: number) => frame(String(index)),
  } satisfies PresentationSource;
  const errors: unknown[] = [];
  try {
    const view = render(
      <PresentationOverlay
        handle={handle}
        slideCount={2}
        startIndex={0}
        resolveImage={(id) => (id === '0' ? pending : images[1])}
        onExit={() => {}}
        onError={(error) => errors.push(error)}
      />
    );
    const displayed = view.container.querySelector('canvas')!;
    fireEvent.click(view.getByRole('button', { name: 'Next slide' }));
    await waitFor(() => expect(draws.get(displayed)?.length).toBe(1));
    const painted = draws.get(displayed)![0][0] as HTMLCanvasElement;
    expect(draws.get(painted)?.[0][0]).toBe(images[1]);
    await act(async () => {
      finishFirst(images[0]);
      await pending;
    });
    expect(draws.get(displayed)?.length).toBe(1);
    expect(view.getByText('2 / 2')).toBeTruthy();
    expect(errors).toEqual([]);
  } finally {
    context.mockRestore();
  }
});

it('clamps navigation after slides are deleted and restores keyboard focus on exit', async () => {
  const focusTarget = document.createElement('button');
  document.body.append(focusTarget);
  focusTarget.focus();
  const layouts: number[] = [];
  const handle = {
    layoutSlide: (index: number) => {
      layouts.push(index);
      return { ...frame(''), primitives: [] };
    },
  } satisfies PresentationSource;
  let exits = 0;
  const props = {
    handle,
    slideCount: 3,
    startIndex: 2,
    resolveImage: () => null,
    onExit: () => {
      exits++;
    },
    onError: () => {},
  };
  const view = render(<PresentationOverlay {...props} />);
  expect(view.getByText('3 / 3')).toBeTruthy();
  view.rerender(<PresentationOverlay {...props} slideCount={1} />);
  expect(view.getByText('1 / 1')).toBeTruthy();
  expect(
    (view.getByRole('button', { name: 'Next slide' }) as HTMLButtonElement).disabled
  ).toBe(true);
  const dialog = view.getByRole('dialog');
  fireEvent.keyDown(dialog, { key: 'Escape' });
  expect(exits).toBe(1);
  view.unmount();
  expect(document.activeElement).toBe(focusTarget);
  focusTarget.remove();
});

it("draws the host's icons for its controls, in both windows", () => {
  const icons = Object.fromEntries(
    DRAWN_ICON_NAMES.map((name) => [
      name,
      function HostIcon() {
        return <i data-host-icon={name} />;
      },
    ])
  ) as unknown as IconSet;
  const notes = new NotesWindow({ open: () => {}, close: () => {} });
  const view = render(
    <IconSetContext.Provider value={icons}>
      <PresentationOverlay
        handle={{ layoutSlide: () => ({ ...frame(''), primitives: [] }) }}
        slideCount={2}
        startIndex={0}
        resolveImage={() => null}
        notesWindow={notes}
        onExit={() => {}}
        onError={() => {}}
      />
    </IconSetContext.Provider>
  );
  const drawn = (root: ParentNode) =>
    [...root.querySelectorAll('[data-host-icon]')].map((icon) => icon.getAttribute('data-host-icon'));
  expect(view.container.querySelector('svg')).toBeNull();
  // Windowed (no full screen here), with the notes window to open from ⋮.
  const seen = new Set(drawn(view.container));
  expect(new Set(drawn(view.container))).toEqual(
    new Set([
      'presentationExit',
      'presentationPrevious',
      'presentationNext',
      'presentationFullscreen',
      'presentationMore',
    ])
  );
  const popup = new HappyWindow({ url: 'about:blank' }) as unknown as Window;
  act(() => notes.set(popup));
  expect(popup.document.body.querySelector('svg')).toBeNull();
  expect(new Set(drawn(popup.document.body))).toEqual(
    new Set([
      'presenterPause',
      'presenterReset',
      'presenterNotesSize',
      'presentationPrevious',
      'presentationNext',
    ])
  );
  for (const name of drawn(popup.document.body)) seen.add(name);
  fireEvent.click(popup.document.body.querySelector('[aria-label="Pause timer"]')!);
  for (const name of [...drawn(view.container), ...drawn(popup.document.body)]) seen.add(name);
  expect(seen).toEqual(new Set(DRAWN_ICON_NAMES));
});

it('presents a deck through the viewer handle', async () => {
  const root = resolve(import.meta.dir, '../../../..');
  const [wasm, deck, font] = await Promise.all([
    readFile(resolve(root, 'packages/pptx/src/wasm/generated/viewer/pptx_view_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  const handle = openPresentation(deck, { fonts: [{ family: 'Liberation Sans', bytes: font }] });
  const slideCount = handle.snapshot().slides.length;
  const draws = new Map<HTMLCanvasElement, number>();
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
    function (this: HTMLCanvasElement) {
      return new Proxy(
        {},
        {
          get: (_target, key) =>
            key === 'drawImage'
              ? () => draws.set(this, (draws.get(this) ?? 0) + 1)
              : key === 'measureText'
                ? () => ({ width: 0 })
                : () => {},
        }
      ) as CanvasRenderingContext2D;
    } as unknown as typeof HTMLCanvasElement.prototype.getContext
  );
  const errors: unknown[] = [];
  try {
    const view = render(
      <PresentationOverlay
        handle={handle}
        slideCount={slideCount}
        startIndex={0}
        resolveImage={() => null}
        onExit={() => {}}
        onError={(error) => errors.push(error)}
      />
    );
    const displayed = view.container.querySelector('canvas')!;
    await waitFor(() => expect(draws.get(displayed)).toBe(1));
    fireEvent.click(view.getByRole('button', { name: 'Next slide' }));
    await waitFor(() => expect(draws.get(displayed)).toBe(2));
    expect(view.getByText(`2 / ${slideCount}`)).toBeTruthy();
    expect(errors).toEqual([]);
  } finally {
    context.mockRestore();
    handle.dispose();
  }
});
