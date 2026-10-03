import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, it, spyOn } from 'bun:test';
import { paintSlide, type SlideDisplayList } from '@betteroffice/pptx';
import { SlideThumbnail } from './PptxEditor';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

const frame: SlideDisplayList = {
  contractVersion: 1,
  width: 960,
  height: 540,
  primitives: [
    { kind: 'image', name: 'Slide picture', assetId: 'media1', objectId: 1, x: 40, y: 40, w: 400, h: 300 },
  ],
};

it("shows one paint of the slide when the strip re-renders while that slide's images load", async () => {
  // Every canvas gets a context that logs its calls; a canvas drawn onto
  // another is logged as that canvas, so the shown pixels can be traced.
  const calls = new Map<HTMLCanvasElement, unknown[][]>();
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
    function (this: HTMLCanvasElement) {
      const log = calls.get(this) ?? [];
      calls.set(this, log);
      return new Proxy({}, {
        get: (_target, key) => (...args: unknown[]) => log.push([key, ...args]),
      }) as CanvasRenderingContext2D;
    } as unknown as typeof HTMLCanvasElement.prototype.getContext
  );
  try {
    const image = {} as CanvasImageSource;
    let finishFirst: (image: CanvasImageSource) => void = () => {};
    const late = new Promise<CanvasImageSource>((resolve) => (finishFirst = resolve));
    // The editor passes a new resolver when it re-renders; the first paint is
    // still waiting for its image then.
    const view = render(<SlideThumbnail frame={frame} resolveImage={() => late} />);
    view.rerender(<SlideThumbnail frame={frame} resolveImage={() => image} />);
    await act(async () => {
      await Promise.resolve();
      finishFirst(image);
      await late;
      await new Promise((settle) => setTimeout(settle, 0));
    });

    const fresh = document.createElement('canvas');
    await paintSlide(fresh.getContext('2d')!, frame, window.devicePixelRatio || 1, 126 / frame.width, {
      resolveImage: () => image,
    });
    const shown = view.container.querySelector('canvas')!;
    const traced = (canvas: HTMLCanvasElement): unknown[][] => {
      const log = calls.get(canvas) ?? [];
      const copy = log.length === 1 && log[0][0] === 'drawImage' ? log[0][1] : null;
      return copy instanceof HTMLCanvasElement ? traced(copy) : log;
    };
    expect(traced(shown)).toEqual(calls.get(fresh)!);
  } finally {
    context.mockRestore();
  }
});
