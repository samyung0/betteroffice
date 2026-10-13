import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { repackDocx } from '@betteroffice/docx/docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { unzipContainer } from '@betteroffice/docx/wasm/opc';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import { yrsToDocument } from '@betteroffice/docx/yrs/yrsToDocument';
import { PagedEditor, type PagedEditorRef } from './PagedEditor';
import type { YrsCoreSession } from './hooks/useYrsCoreSession';

// A PagedEditor over a real engine session, for tests that drive the editor's
// own commands and read what the save writes. Needs a registered DOM.

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const NAMESPACES = `xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;

let fontBytes: ArrayBuffer | null = null;
const sessions: YrsSession[] = [];

export async function loadEditor(): Promise<void> {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {} },
      configurable: true,
    });
  }
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  fontBytes = readFileSync(FONT).buffer as ArrayBuffer;
}

export function destroySessions(): void {
  for (const session of sessions.splice(0)) session.destroy();
}

/** A package with `body` (paragraph XML), plus extra parts and document relationships. */
export function docxPackage(
  body: string,
  extra: { parts?: Record<string, Uint8Array>; rels?: string; contentTypes?: string } = {}
): Uint8Array {
  const parts = new Map<string, Uint8Array>([
    [
      '[Content_Types].xml',
      toBytes(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${extra.contentTypes ?? ''}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
      ),
    ],
    [
      '_rels/.rels',
      toBytes(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
      ),
    ],
    [
      'word/_rels/document.xml.rels',
      toBytes(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${extra.rels ?? ''}</Relationships>`
      ),
    ],
    [
      'word/document.xml',
      toBytes(
        `<w:document ${NAMESPACES}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`
      ),
    ],
    ...Object.entries(extra.parts ?? {}),
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

export async function openSession(bytes: Uint8Array, clientId: number): Promise<YrsSession> {
  const session = await createYrsSession({ clientId });
  sessions.push(session);
  session.openDocx(bytes, true);
  return session;
}

/** A second person joining `state`, the room seeded from `bytes`. */
export async function joinSession(
  bytes: Uint8Array,
  state: Uint8Array,
  clientId: number
): Promise<YrsSession> {
  const session = await createYrsSession({ clientId });
  sessions.push(session);
  session.openDocx(bytes, false);
  session.loadState(state);
  return session;
}

/** Exchanges both sessions' updates. */
export function exchange(a: YrsSession, b: YrsSession): void {
  a.applyUpdate(b.encodeStateAsUpdate(a.encodeStateVector()));
  b.applyUpdate(a.encodeStateAsUpdate(b.encodeStateVector()));
}

function yrsCore(session: YrsSession): YrsCoreSession {
  return {
    session,
    storyBlocks: () => null,
    bodyBlocks: () => null,
    inputPositionMap: () => null,
    displayPositionToLoc: () => null,
    locToDisplayPosition: () => null,
    documentFromYrs: () => null,
    publishDirectInput: () => {},
    sourceNumbering: () => undefined,
    sourceDocument: () => undefined,
  };
}

/** Renders a PagedEditor on `session` and waits for its first layout. */
export async function renderPagedEditor(
  session: YrsSession
): Promise<React.RefObject<PagedEditorRef | null>> {
  const { act, render } = await import('@testing-library/react');
  const ref = createRef<PagedEditorRef>();
  const layouts: Layout[] = [];
  render(
    <PagedEditor
      ref={ref}
      document={null}
      yrsCore={yrsCore(session)}
      measurementFontProvider={{ resolve: () => () => Promise.resolve(fontBytes!) }}
      onLayoutComputed={(layout) => {
        if (layout) layouts.push(layout);
      }}
    />
  );
  const deadline = Date.now() + 2000;
  while (layouts.length === 0 && Date.now() < deadline)
    await act(async () => {
      await new Promise((done) => setTimeout(done, 10));
    });
  if (layouts.length === 0) throw new Error('the editor never laid out');
  return ref;
}

/** The package the save writes for `session`. */
export async function savedBytes(session: YrsSession): Promise<Uint8Array> {
  const base = session.materializeDocx();
  if (!base) throw new Error('no source package');
  return new Uint8Array(await repackDocx(yrsToDocument(session, base)));
}

/** The parts of the package the save writes for `session`, as text. */
export async function savedParts(session: YrsSession): Promise<Record<string, string>> {
  return Object.fromEntries(
    Object.entries(unzipContainer(await savedBytes(session))).map(([path, bytes]) => [
      path,
      new TextDecoder().decode(bytes),
    ])
  );
}
