/**
 * YrsInput's selection emit reads `onStateChange` through a ref: a parent
 * passing a new callback on every frame (a long document's page batches,
 * Q815 C) re-renders the input once, and the selection-emit effect does not
 * run again (its setSelectionEpoch was a second commit per frame, which
 * added up to React's nested-update limit).
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { Profiler } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { YrsInput } from './YrsInput';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render } = await import('@testing-library/react');
const sessions: YrsSession[] = [];

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'))
    )
  )
);
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a new onStateChange with the same selection commits once', async () => {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  zip.file(
    'word/document.xml',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>One line</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'
  );
  const session = await createYrsSession();
  sessions.push(session);
  session.openDocx(await zip.generateAsync({ type: 'uint8array' }), true);
  const map = (story = 'body') => createYrsInputPositionMap(story, session.paragraphSpans(story));
  const toLoc = (position: number, story: string) => displayPositionToYrsLoc(map(story), position);
  const toPosition = (loc: Parameters<typeof yrsLocToDisplayPosition>[1]) =>
    yrsLocToDisplayPosition(map(loc.story), loc);
  const noop = () => {};
  let commits = 0;
  const view = (onStateChange: () => void) => (
    <Profiler id="input" onRender={() => (commits += 1)}>
      <YrsInput
        enabled
        readOnly={false}
        session={session}
        inputPositionMap={map}
        displayPositionToLoc={toLoc}
        locToDisplayPosition={toPosition}
        displayListQueries={null}
        onStateChange={onStateChange}
        onDirectInput={noop}
      />
    </Profiler>
  );
  const shown = render(view(() => {}));
  await act(async () => {});
  commits = 0;
  shown.rerender(view(() => {}));
  await act(async () => {});
  expect(commits).toBe(1);
});
