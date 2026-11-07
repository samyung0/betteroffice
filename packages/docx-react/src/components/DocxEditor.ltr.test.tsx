import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { DocxEditor } from './DocxEditor';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

// A host page with <html dir="rtl"> must not mirror the editor (upstream #1884).
test('pins the empty editor root to ltr', () => {
  document.documentElement.dir = 'rtl';
  try {
    const view = render(<DocxEditor />);
    expect(view.container.querySelector('[data-testid="docx-editor"]')?.getAttribute('dir')).toBe('ltr');
  } finally {
    document.documentElement.dir = '';
  }
});
