import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { toolbarValueToLayoutTarget } from '@betteroffice/docx/docx';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type { PagedEditorRef } from './PagedEditor';
import { createYrsPositionProjection } from './internals/yrsPositionProjection';
import {
  destroySessions,
  docxPackage,
  exchange,
  joinSession,
  loadEditor,
  openSession,
  renderPagedEditor,
  savedParts,
} from './pagedEditor.testing';

const { act, cleanup } = await import('@testing-library/react');

beforeAll(loadEditor);
afterEach(async () => {
  await act(async () => {
    await new Promise((done) => setTimeout(done, 50));
  });
  cleanup();
  destroySessions();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
  ),
  (char) => char.charCodeAt(0)
);
const INCH = 914400;
const PICTURE = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${INCH}" cy="${INCH}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="dot.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${INCH}" cy="${INCH}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const pictureDocx = () =>
  docxPackage(
    `<w:p w14:paraId="10000001"><w:r><w:t xml:space="preserve">Figure: </w:t></w:r>${PICTURE}<w:r><w:t xml:space="preserve"> done.</w:t></w:r></w:p>`,
    {
      parts: { 'word/media/image1.png': PNG },
      rels: '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>',
      contentTypes: '<Default Extension="png" ContentType="image/png"/>',
    }
  );

/** The picture's display position, as the editor's pointer and menus address it. */
function picturePosition(session: YrsSession): number {
  const projection = createYrsPositionProjection(session, 'body')!;
  for (let position = 0; position <= projection.size; position += 1)
    if (projection.nodeAt(position)?.kind === 'image') return position;
  throw new Error('no picture');
}

const command = (
  editor: React.RefObject<PagedEditorRef | null>,
  input: Parameters<PagedEditorRef['applyYrsCommand']>[0]
) => {
  let applied = false;
  act(() => {
    applied = editor.current?.applyYrsCommand(input) ?? false;
  });
  return applied;
};

test('resize, square wrap, rotate, flip and alt text from the editor are saved', async () => {
  const session = await openSession(pictureDocx(), 7501);
  const editor = await renderPagedEditor(session);
  const pmPos = picturePosition(session);

  // The resize handles report pixels; 192 px is two inches.
  expect(command(editor, { type: 'imageGeometry', pmPos, patch: { width: 192, height: 96 } })).toBe(
    true
  );
  expect(
    command(editor, { type: 'imageWrap', pmPos, target: toolbarValueToLayoutTarget('square')! })
  ).toBe(true);
  expect(command(editor, { type: 'imageTransform', pmPos, action: 'rotateCW' })).toBe(true);
  expect(command(editor, { type: 'imageTransform', pmPos, action: 'flipH' })).toBe(true);
  // Picture options: alt text.
  expect(command(editor, { type: 'imageGeometry', pmPos, patch: { alt: 'A red dot' } })).toBe(
    true
  );

  const parts = await savedParts(session);
  const xml = parts['word/document.xml'];
  const anchor = /<wp:anchor\b[\s\S]*?<\/wp:anchor>/.exec(xml)?.[0];
  expect(anchor).toBeDefined();
  expect(xml).not.toContain('<wp:inline');
  expect(anchor).toContain(`<wp:extent cx="${2 * INCH}" cy="${INCH}"/>`);
  // Square wrap puts the picture at the left, text on its right.
  expect(anchor).toContain('<wp:wrapSquare wrapText="right"/>');
  expect(anchor).toMatch(/<a:xfrm rot="5400000" flipH="1">/);
  expect(anchor).toContain('r:embed="rIdImg"');
  expect(anchor).toMatch(/<wp:docPr [^>]*descr="A red dot"/);
  expect(parts['word/media/image1.png']).toBeDefined();
  expect(xml).toContain('Figure: ');
  expect(xml).toContain(' done.');
});

test('Insert › Image puts the picked picture at the caret and the save adds its part', async () => {
  const session = await openSession(pictureDocx(), 7521);
  const editor = await renderPagedEditor(session);
  const [{ paraId }] = session.paragraphs('body');
  session.setSelection({ story: 'body', paraId, offset: 0 });
  // The payload Insert › Image (useFileIO) builds from the picked file.
  const inserted = command(editor, {
    type: 'insertImage',
    image: {
      src: `data:image/png;base64,${btoa(String.fromCharCode(...PNG))}`,
      alt: 'picked.png',
      width: 48,
      height: 24,
      rId: 'rId_img_1',
      wrapType: 'inline',
      displayMode: 'inline',
    },
  });
  expect(inserted).toBe(true);
  expect(session.selection()?.head).toMatchObject({ paraId, offset: 1 });

  const parts = await savedParts(session);
  const xml = parts['word/document.xml'];
  const drawings = [...xml.matchAll(/<wp:inline\b[\s\S]*?<\/wp:inline>/g)].map((m) => m[0]);
  expect(drawings).toHaveLength(2);
  expect(drawings[0]).toContain(`<wp:extent cx="${INCH / 2}" cy="${INCH / 4}"/>`);
  const id = /r:embed="([^"]+)"/.exec(drawings[0])![1];
  expect(id).not.toBe('rIdImg');
  const target = new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`).exec(
    parts['word/_rels/document.xml.rels']
  )?.[1];
  expect(target).toMatch(/^media\//);
  expect(parts[`word/${target}`]).toBeDefined();
  expect(xml.indexOf('<wp:inline')).toBeLessThan(xml.indexOf('Figure: '));
});

test('two people resizing and wrapping or rotating one picture at once keep every change', async () => {
  // Both orders, so each person wins the entries both commands write.
  for (const [alice, bob] of [
    [7511, 7512],
    [7514, 7513],
  ]) {
    const bytes = pictureDocx();
    const first = await openSession(bytes, alice);
    const second = await joinSession(bytes, first.encodeState(), bob);
    const firstEditor = await renderPagedEditor(first);
    const secondEditor = await renderPagedEditor(second);
    const pmPos = picturePosition(first);

    command(firstEditor, { type: 'imageGeometry', pmPos, patch: { width: 192, height: 96 } });
    command(secondEditor, {
      type: 'imageWrap',
      pmPos,
      target: toolbarValueToLayoutTarget('square')!,
    });
    command(secondEditor, { type: 'imageTransform', pmPos, action: 'rotateCW' });
    exchange(first, second);

    expect(second.storySegments('body')).toEqual(first.storySegments('body'));
    for (const session of [first, second]) {
      const anchor = /<wp:anchor\b[\s\S]*?<\/wp:anchor>/.exec(
        (await savedParts(session))['word/document.xml']
      )?.[0];
      expect(anchor).toContain(`<wp:extent cx="${2 * INCH}" cy="${INCH}"/>`);
      expect(anchor).toContain('<wp:wrapSquare wrapText="right"/>');
      expect(anchor).toMatch(/<a:xfrm rot="5400000">/);
    }
    cleanup();
  }
});
