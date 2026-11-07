import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { LocaleProvider, useTranslation } from './LocaleContext';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

function Lang() {
  return <span data-testid="lang">{useTranslation().lang}</span>;
}

test('a malformed host locale formats with en instead of throwing', () => {
  for (const _lang of ['en_US', '', 'not a locale']) {
    const { getByTestId, unmount } = render(<LocaleProvider i18n={{ _lang }}><Lang /></LocaleProvider>);
    expect(getByTestId('lang').textContent).toBe('en');
    unmount();
  }
});

test('a valid host locale is canonicalized', () => {
  const { getByTestId } = render(<LocaleProvider i18n={{ _lang: 'de-de' }}><Lang /></LocaleProvider>);
  expect(getByTestId('lang').textContent).toBe('de-DE');
});
