import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';

import type { DisplayPage } from './displayList';
import { buildInteractiveOverlayPage, syncInteractiveOverlayPositions } from './interactiveOverlay';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('an inline widget keeps its button and takes its shifted position', () => {
  const widget = { kind: 'checkbox', groupId: 'sdt:1:7', pos: 12 };
  const page = {
    pageIndex: 0,
    width: 500,
    height: 500,
    primitives: [
      {
        kind: 'text',
        text: 'x',
        x: 10,
        baselineY: 20,
        width: 8,
        font: '400 16px Calibri',
        color: '#000000',
        docStart: 13,
        docEnd: 14,
        inlineSdtWidget: widget,
      },
    ],
  } as unknown as DisplayPage;
  const overlay = buildInteractiveOverlayPage(page);
  const button = overlay.querySelector<HTMLElement>('[data-sdt-pos]')!;
  expect(button.dataset.sdtPos).toBe('12');
  // the owned frame apply moves the widget's pos on the same attrs object
  widget.pos += 3;
  expect(syncInteractiveOverlayPositions(overlay)).toBe(true);
  expect(overlay.querySelector('[data-sdt-pos]') === button).toBe(true);
  expect(button.dataset.sdtPos).toBe('15');
});
