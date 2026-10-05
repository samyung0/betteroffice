import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const { watchTextDrag } = await import('./CanvasPagesView');

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const press = (target: EventTarget, type: string, pointerType: string) =>
  target.dispatchEvent(new PointerEvent(type, { bubbles: true, button: 0, pointerType }));

test('a mouse or pen drag in the host holds; touch pans and a cancelled press ends it', () => {
  const host = document.createElement('div');
  const outside = document.createElement('div');
  document.body.append(host, outside);
  const watch = watchTextDrag(host);

  press(host, 'pointerdown', 'mouse');
  expect(watch.held()).toBe(true);
  press(host, 'pointerup', 'mouse');
  expect(watch.held()).toBe(false);

  press(host, 'pointerdown', 'pen');
  expect(watch.held()).toBe(true);
  press(host, 'pointercancel', 'pen');
  expect(watch.held()).toBe(false);

  press(host, 'pointerdown', 'touch');
  expect(watch.held()).toBe(false);
  press(outside, 'pointerdown', 'mouse');
  expect(watch.held()).toBe(false);

  watch.stop();
  press(host, 'pointerdown', 'mouse');
  expect(watch.held()).toBe(false);
  host.remove();
  outside.remove();
});
