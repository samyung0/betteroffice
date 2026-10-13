import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { buildInteractiveOverlayPage, type DisplayList } from '@betteroffice/docx/layout/render';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { ContentControlWidgets } from './ContentControlWidgets';
import {
  destroySessions,
  docxPackage,
  loadEditor,
  openSession,
  renderPagedEditor,
  savedParts,
} from './pagedEditor.testing';

const { act, cleanup, fireEvent, render } = await import('@testing-library/react');

beforeAll(loadEditor);
afterEach(async () => {
  await act(async () => {
    await new Promise((done) => setTimeout(done, 50));
  });
  cleanup();
  destroySessions();
  document.body.innerHTML = '';
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const CHECK_BOX = `<w:sdt><w:sdtPr><w:alias w:val="Done"/><w:tag w:val="done"/><w:id w:val="201"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>`;

/** The first page's display list, as the canvas pages lay it out. */
function firstPage(session: YrsSession): DisplayList['pages'][number] {
  const layout = session.layoutDocumentWithRegionsJson(
    JSON.stringify({
      bodyStory: 'body',
      regions: { sections: [{ sectionId: 'main', properties: {} }] },
      notes: { contents: [] },
      measurement: {
        fontChains: {},
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        authoritativeShaping: false,
      },
      renderEnv: {},
    })
  );
  return (JSON.parse(session.buildDisplayListJson(layout)) as DisplayList).pages[0];
}

test('a check box control toggles from its widget and the save writes its state', async () => {
  const session = await openSession(
    docxPackage(
      `<w:p w14:paraId="10000001"><w:r><w:t xml:space="preserve">Done: </w:t></w:r>${CHECK_BOX}</w:p>`
    ),
    7701
  );
  const editor = await renderPagedEditor(session);
  // The canvas page's interactive overlay, with the widgets DocxEditor mounts over it.
  const container = document.createElement('div');
  document.body.appendChild(container);
  const overlay = () => {
    container.replaceChildren(buildInteractiveOverlayPage(firstPage(session)));
    return container.querySelector<HTMLElement>('.layout-sdt-widget[data-sdt-tag="done"]')!;
  };
  render(
    <ContentControlWidgets
      containerRef={{ current: container }}
      applyYrsValue={(pmPos, value, embedId) =>
        editor.current?.applyYrsCommand({
          type: 'contentControlValue',
          pmPos,
          ...(embedId ? { embedId } : {}),
          value,
        }) ?? false
      }
    />
  );
  const saved = async () => (await savedParts(session))['word/document.xml'];

  let widget = overlay();
  expect(widget.getAttribute('aria-checked')).toBe('false');
  act(() => {
    fireEvent.click(widget);
  });
  let xml = await saved();
  expect(xml).toContain('<w14:checked w14:val="1"/>');
  expect(xml).toContain('<w:t>☒</w:t>');

  widget = overlay();
  expect(widget.getAttribute('aria-checked')).toBe('true');
  act(() => {
    fireEvent.click(widget);
  });
  xml = await saved();
  expect(xml).toContain('<w14:checked w14:val="0"/>');
  expect(xml).toContain('<w:t>☐</w:t>');
});
