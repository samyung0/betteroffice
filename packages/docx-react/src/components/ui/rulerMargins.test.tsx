import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { HorizontalRuler } from './HorizontalRuler';
import { VerticalRuler } from './VerticalRuler';

const { cleanup, render } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const draggable = (container: HTMLElement) =>
  Object.fromEntries(
    [...container.querySelectorAll('[data-ruler-margin]')].map((zone) => [
      zone.getAttribute('data-ruler-margin'),
      zone.getAttribute('data-draggable'),
    ])
  );

test('a margin drags only when the host takes its change', () => {
  const change = () => {};
  const withHandlers = render(
    <>
      <HorizontalRuler editable onLeftMarginChange={change} onRightMarginChange={change} />
      <VerticalRuler editable onTopMarginChange={change} onBottomMarginChange={change} />
    </>
  );
  expect(draggable(withHandlers.container)).toEqual({
    left: 'true',
    right: 'true',
    top: 'true',
    bottom: 'true',
  });
  cleanup();
  const without = render(
    <>
      <HorizontalRuler editable />
      <VerticalRuler editable />
    </>
  );
  expect(draggable(without.container)).toEqual({
    left: 'false',
    right: 'false',
    top: 'false',
    bottom: 'false',
  });
});
