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
      applyYrsValue={(target, value, embedId) =>
        editor.current?.applyYrsCommand({
          type: 'contentControlValue',
          ...(target.story !== undefined ? { story: target.story } : { pmPos: target.pos }),
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

const blockControl = (properties: string, paraId: string, text: string) =>
  `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent><w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p></w:sdtContent></w:sdt>`;

// Block controls are named by their child story, not their position, so text
// typed before them leaves their widgets pointing at them. (The test editor's
// position map is a stub, so the controls carry a w:id to be written by.)
test('block check box, date and drop-down controls set their values after typing before them', async () => {
  const session = await openSession(
    docxPackage(
      `<w:p w14:paraId="10000001"><w:r><w:t>Lead</w:t></w:r></w:p>` +
        blockControl(
          `<w:tag w:val="agreed"/><w:id w:val="301"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>`,
          '10000002',
          '☐'
        ) +
        blockControl(
          `<w:tag w:val="when"/><w:id w:val="302"/><w:date w:fullDate="2026-01-02T00:00:00Z"><w:dateFormat w:val="yyyy-MM-dd"/><w:lid w:val="en-US"/></w:date>`,
          '10000003',
          '2026-01-02'
        ) +
        blockControl(
          `<w:tag w:val="choice"/><w:id w:val="303"/><w:dropDownList><w:listItem w:displayText="Alpha" w:value="a"/><w:listItem w:displayText="Beta" w:value="b"/></w:dropDownList>`,
          '10000004',
          'Alpha'
        )
    ),
    7702
  );
  const editor = await renderPagedEditor(session);
  const container = document.createElement('div');
  document.body.appendChild(container);
  const widget = (tag: string) => {
    container.replaceChildren(buildInteractiveOverlayPage(firstPage(session)));
    return container.querySelector<HTMLElement>(`.layout-sdt-widget[data-sdt-tag="${tag}"]`)!;
  };
  render(
    <ContentControlWidgets
      containerRef={{ current: container }}
      applyYrsValue={(target, value, embedId) =>
        editor.current?.applyYrsCommand({
          type: 'contentControlValue',
          ...(target.story !== undefined ? { story: target.story } : { pmPos: target.pos }),
          ...(embedId ? { embedId } : {}),
          value,
        }) ?? false
      }
    />
  );
  const before = widget('agreed').dataset.sdtGroupId;
  await act(async () => {
    session.insertText({ story: 'body', paraId: '10000001', offset: 0 }, 'Typed first. ');
    await new Promise((done) => setTimeout(done, 100));
  });
  expect(widget('agreed').dataset.sdtGroupId).toBe(before!);

  act(() => {
    fireEvent.click(widget('agreed'));
  });
  act(() => {
    fireEvent.click(widget('when'));
  });
  const date = document.querySelector<HTMLInputElement>('.layout-sdt-widget-date')!;
  act(() => {
    fireEvent.change(date, { target: { value: '2026-03-04' } });
  });
  // A block drop-down's popup lists no items (the display list carries none),
  // so its value goes through the command the popup would send.
  const choice = /^sdt:(.+)$/.exec(widget('choice').dataset.sdtGroupId ?? '')![1]!;
  expect(
    editor.current!.applyYrsCommand({
      type: 'contentControlValue',
      story: choice,
      value: { kind: 'dropdown', value: 'b' },
    })
  ).toBe(true);

  const xml = (await savedParts(session))['word/document.xml'];
  expect(xml).toContain('Typed first. Lead');
  expect(xml).toContain('<w14:checked w14:val="1"/>');
  expect(xml).toContain('<w:t>☒</w:t>');
  expect(xml).toContain('w:fullDate="2026-03-04');
  expect(xml).toContain('<w:t>2026-03-04</w:t>');
  expect(xml).toContain('<w:t>Beta</w:t>');
});
