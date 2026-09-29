import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { useCallback, useState } from 'react';
import { writeDocumentWithRust } from '@betteroffice/docx/docx/rustSaveFacade';
import {
  createYrsSidebarProjection,
  extractTrackedChangesFromYrs,
} from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { preloadOpcWasm, rezipContainer, unzipContainer } from '@betteroffice/docx/wasm/opc';
import { preloadParseWasm } from '@betteroffice/docx/wasm/parse';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import { yrsToDocument } from '@betteroffice/docx/yrs/yrsToDocument';
import { useReviewAllItems } from './useReviewAllItems';

const { act, cleanup, fireEvent, render, screen } = await import('@testing-library/react');

const GENERATED = resolve(import.meta.dir, '../../../docx/src/wasm/generated');
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const TYPES = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const RELS = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const ins = (xml: string) =>
  `<w:ins w:id="90" w:author="A" w:date="2026-09-01T00:00:00Z">${xml}</w:ins>`;
const del = (text: string) =>
  `<w:del w:id="91" w:author="A" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>${text}</w:delText></w:r></w:del>`;
const field = (result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
/** A change Word keeps inside a field result, which the revision list leaves out. */
const fieldChange = field(`${run('20')}${ins(run('26'))}${del('19')}`);

function docx(paragraph: string): Uint8Array {
  const encode = (xml: string) => new TextEncoder().encode(xml);
  return rezipContainer({
    '[Content_Types].xml': encode(TYPES),
    '_rels/.rels': encode(RELS),
    'word/document.xml': encode(
      `<w:document ${W}><w:body><w:p>${paragraph}</w:p><w:sectPr/></w:body></w:document>`
    ),
  });
}

const wasm = (path: string) => new Uint8Array(readFileSync(resolve(GENERATED, path)));
beforeAll(async () => {
  await preloadEditWasm(wasm('edit/docx_edit_bg.wasm'));
  await preloadParseWasm(wasm('parse/docx_parse_bg.wasm'));
  await preloadOpcWasm(wasm('opc/ooxml_opc_bg.wasm'));
});
afterEach(() => cleanup());
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

let clientId = 51900;
const sessions: YrsSession[] = [];
async function open(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: (clientId += 1) });
  session.openDocx(bytes, true);
  sessions.push(session);
  return session;
}
afterAll(() => sessions.forEach((session) => session.destroy()));

const listed = (session: YrsSession) =>
  extractTrackedChangesFromYrs(
    session.listRevisions(),
    createYrsSidebarProjection(session),
    session.hasFieldChanges()
  );

/** The review-all items as the editor's sidebar renders them, re-read after each resolve. */
function Sidebar({ session, readOnly = false }: { session: YrsSession; readOnly?: boolean }) {
  const [trackedChanges, setTrackedChanges] = useState(() => listed(session));
  const getSession = useCallback(() => session, [session]);
  const onResolved = useCallback((resolved: YrsSession) => setTrackedChanges(listed(resolved)), []);
  const items = useReviewAllItems({ getSession, trackedChanges, readOnly, onResolved });
  return (
    <>
      {items.map((item) => (
        <div key={item.id}>
          {item.render({ isExpanded: false, onToggleExpand: () => {}, measureRef: () => {} })}
        </div>
      ))}
    </>
  );
}

const buttons = () => screen.queryAllByRole('button').map((button) => button.textContent);

async function exported(bytes: Uint8Array, session: YrsSession): Promise<string> {
  const document = yrsToDocument(session, session.materializeDocx()!);
  const { buffer } = await writeDocumentWithRust(document, bytes.slice().buffer as ArrayBuffer);
  return new TextDecoder().decode(unzipContainer(new Uint8Array(buffer))['word/document.xml']);
}

test('Accept all and Reject all show only while the document holds changes it can resolve', async () => {
  render(<Sidebar session={await open(docx(run('plain')))} />);
  expect(buttons()).toEqual([]);
  cleanup();

  const inField = await open(docx(`${run('a ')}${fieldChange}`));
  expect(inField.listRevisions()).toEqual([]);
  render(<Sidebar session={inField} />);
  expect(buttons()).toEqual(['Accept all', 'Reject all']);
  // Only changes inside fields: the page marks none, so a line says where they are.
  expect(screen.queryByText(/inside fields/)).not.toBeNull();
  cleanup();

  render(<Sidebar session={await open(docx(`${run('a ')}${ins(run('new'))}${fieldChange}`))} />);
  expect(buttons()).toEqual(['Accept all', 'Reject all']);
  expect(screen.queryByText(/inside fields/)).toBeNull();
  cleanup();

  render(<Sidebar session={inField} readOnly />);
  expect(buttons()).toEqual([]);
});

test.each([
  ['Accept all', 'a new 2026 b'],
  ['Reject all', 'a  2019 b'],
])('%s resolves listed and in-field changes, and the export holds none', async (label, text) => {
  const bytes = docx(`${run('a ')}${ins(run('new'))}${run(' ')}${fieldChange}${run(' b')}`);
  const session = await open(bytes);
  expect(session.listRevisions()).not.toEqual([]);
  expect(session.hasFieldChanges()).toBe(true);
  render(<Sidebar session={session} />);
  await act(async () => screen.getByText(label).click());

  expect(session.listRevisions()).toEqual([]);
  expect(session.hasFieldChanges()).toBe(false);
  expect(buttons()).toEqual([]);
  const xml = await exported(bytes, session);
  expect(xml).not.toMatch(/<w:(ins|del)\b/);
  expect([...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(([, t]) => t).join('')).toBe(text);
});

test('Accept all is its own undo step and leaves the editor focused', async () => {
  const session = await open(docx(`${run('a ')}${fieldChange}`));
  render(<Sidebar session={session} />);
  const paraId = session.paragraphs('body')[0]!.paraId;
  session.insertText({ story: 'body', paraId, offset: 0 }, 'typed ');
  const accept = screen.getByText('Accept all');
  // A prevented mousedown keeps focus where it was, so Ctrl+Z reaches the editor.
  expect(fireEvent.mouseDown(accept)).toBe(false);
  await act(async () => accept.click());
  expect(session.hasFieldChanges()).toBe(false);
  session.undo();
  expect(session.hasFieldChanges()).toBe(true);
  expect(session.paragraphs('body')[0]!.text).toStartWith('typed a ');
});
