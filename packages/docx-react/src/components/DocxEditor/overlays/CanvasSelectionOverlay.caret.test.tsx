import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { DisplayList, DisplayPrimitive } from '@betteroffice/docx/layout/render';
import { CanvasSelectionOverlay } from './CanvasSelectionOverlay';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

// White text in a dark-shaded cell: the canvas caret takes the text's colour
// rather than the page token (upstream #2205).
test('the canvas caret takes the colour of the text under it', () => {
  const shading = { kind: 'rect', x: 90, y: 90, w: 200, h: 40, fill: '#1f1f1f' } as DisplayPrimitive;
  const run = { kind: 'glyphRun', fontId: 1, size: 16, color: '#ffffff', text: 'ab',
    glyphs: [{ id: 1, x: 100, y: 115, cluster: 0, advance: 30 }, { id: 2, x: 130, y: 115, cluster: 1, advance: 30 }] } as DisplayPrimitive;
  const displayList = { pages: [{ pageIndex: 0, width: 816, height: 1056, primitives: [shading, run] }] } as unknown as DisplayList;
  const target = document.createElement('div');
  document.body.append(target);
  try {
    const view = render(
      <CanvasSelectionOverlay selectionRects={[]} caretPosition={{ pageIndex: 0, x: 130, y: 100, height: 20 }}
        isFocused overlayTarget={target} canvasHostRef={{ current: null }} displayList={displayList}
        directProjection sidebarOpen={false} zoom={1} />
    );
    const caret = view.baseElement.querySelector<HTMLElement>('[data-testid="caret"]');
    expect(caret?.style.backgroundColor).toBe('#ffffff');
  } finally {
    target.remove();
  }
});
