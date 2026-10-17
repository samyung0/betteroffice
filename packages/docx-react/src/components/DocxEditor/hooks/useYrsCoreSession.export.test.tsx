import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Document } from '@betteroffice/docx/types/document';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession } from '@betteroffice/docx/yrs';
import { useYrsCoreSession } from './useYrsCoreSession';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');

const FIXTURE = new Uint8Array(
  readFileSync(resolve(import.meta.dir, '__fixtures__/probe-linked-header.docx'))
);

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function seed(): Promise<Uint8Array> {
  const seeder = await createYrsSession({ clientId: 11 });
  try {
    seeder.openDocx(FIXTURE.slice(), true);
    return seeder.encodeState();
  } finally {
    seeder.destroy();
  }
}

/** A peer's edit of the header, as the updates it sends. */
async function peerHeaderEdit(state: Uint8Array, text: string): Promise<Uint8Array[]> {
  const peer = await createYrsSession({ clientId: 22 });
  try {
    peer.loadState(state.slice());
    const story = peer.storyIds().find((id) => id.startsWith('hf:'))!;
    const updates: Uint8Array[] = [];
    const stop = peer.onUpdate((update) => updates.push(update));
    peer.insertText({ story, paraId: peer.paragraphs(story)[0]!.paraId, offset: 0 }, text);
    stop();
    return updates;
  } finally {
    peer.destroy();
  }
}

/** The editor's session as DocxEditor holds it, with PagedEditor's peer-update handling. */
async function openEditor(initialUpdate: Uint8Array) {
  const collaboration = { clientId: 33, initialUpdate };
  let host: Document | null = null;
  const hook = renderHook(
    ({ document }: { document: Document | null }) =>
      useYrsCoreSession(true, document, null, FIXTURE, 1, collaboration, {
        onHostDocument: (next) => {
          host = next.document;
        },
      }),
    { initialProps: { document: null as Document | null } }
  );
  await waitFor(() => expect(hook.result.current.session).not.toBeNull());
  hook.rerender({ document: host });
  const session = hook.result.current.session!;
  session.onUpdate((_update, origin) => {
    if (origin === 'remote') hook.result.current.publishDirectInput();
  });
  // The idle warm materializes the source base before anything is edited.
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 250));
  });
  const body = session.paragraphs('body').find((paragraph) => paragraph.text.length > 0)!;
  session.setSelection({ story: 'body', paraId: body.paraId, offset: 0 });
  return { hook, session, body };
}

const headerText = (document: Document | null) =>
  JSON.stringify([...(document?.package.headers?.values() ?? [])].map((part) => part.content));

describe('useYrsCoreSession export projection', () => {
  test("a peer's header edit reaches the export while the caret is in the body", async () => {
    const state = await seed();
    const { hook, session } = await openEditor(state);
    for (const update of await peerHeaderEdit(state, 'PEER ')) session.applyUpdate(update);

    expect(headerText(hook.result.current.documentFromYrs())).toContain('PEER ');
  });

  test('a header edit saved before the editor opened stays in the export after a body edit', async () => {
    const base = await seed();
    const room = await createYrsSession({ clientId: 44 });
    room.loadState(base);
    for (const update of await peerHeaderEdit(base, 'SAVED ')) room.applyUpdate(update);
    const stored = room.encodeState();
    room.destroy();
    const { hook, session, body } = await openEditor(stored);
    session.insertText({ story: 'body', paraId: body.paraId, offset: 0 }, 'LOCAL ');
    hook.result.current.publishDirectInput();

    const document = hook.result.current.documentFromYrs();
    expect(JSON.stringify(document?.package.document.content)).toContain('LOCAL ');
    expect(headerText(document)).toContain('SAVED ');
  });
});
