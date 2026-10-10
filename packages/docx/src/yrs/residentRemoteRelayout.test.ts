import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import {
  applyFrameDeltaOwned,
  decodeFrameDelta,
  type RetainedFrame,
} from '../layout/render/frameDelta';
import { createEditSession, preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';

const FONT = readFileSync(
  resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')
);

beforeAll(async () => {
  await preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  );
});

const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (id: string, xml: string) => `<w:p w14:paraId="${id}">${xml}</w:p>`;
const char = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = (code: string) => `<w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r>`;
const FIELD_PARA = '44444444';

/** Several small pages under a header, with a bookmark, an internal link and
 * (as in shared/docx-field-containers.test.ts) a REF field continued over two
 * paragraphs whose result holds a nested PAGE field before a link. */
function fixture(): Uint8Array {
  const paragraphs: string[] = [];
  for (let index = 0; index < 60; index += 1) {
    const id = (0x10000000 + index).toString(16).toUpperCase();
    if (index === 2) {
      paragraphs.push(p(id, `${run('Before ')}<w:bookmarkStart w:id="5" w:name="mark"/>${run('Target words')}<w:bookmarkEnd w:id="5"/>`));
    } else if (index === 5) {
      paragraphs.push(
        p(
          FIELD_PARA,
          `${run('a')}${char('begin')}${instr(' REF a \\h ')}${char('separate')}${char('begin')}${instr(' PAGE ')}${char('separate')}${run('7')}${char('end')}<w:hyperlink w:anchor="target">${run('AA')}</w:hyperlink>`
        ),
        p('45454545', `${run('0')}${char('end')}${run('z')}`)
      );
    } else if (index === 8) {
      paragraphs.push(p(id, `${run('See ')}<w:hyperlink w:anchor="mark">${run('the mark')}</w:hyperlink>`));
    } else {
      paragraphs.push(p(id, run(`Paragraph ${index} carries a line of ordinary words`)));
    }
  }
  const parts: PartsMap = new Map();
  const put = (name: string, xml: string) => parts.set(name, toBytes(xml));
  put(
    '[Content_Types].xml',
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`
  );
  put(
    '_rels/.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  put(
    'word/_rels/document.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId20" Type="${REL}/header" Target="header1.xml"/></Relationships>`
  );
  put('word/header1.xml', `<w:hdr ${W}>${p('66666666', run('Running head'))}</w:hdr>`);
  put(
    'word/document.xml',
    `<w:document ${W}><w:body>${paragraphs.join('')}<w:sectPr><w:headerReference w:type="default" r:id="rId20"/><w:pgSz w:w="6000" w:h="4000"/><w:pgMar w:top="600" w:right="500" w:bottom="500" w:left="500" w:header="200" w:footer="200"/></w:sectPr></w:body></w:document>`
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

/** Deterministic PRNG (mulberry32). */
function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function recorder(session: YrsSession): Uint8Array[] {
  const updates: Uint8Array[] = [];
  session.onUpdate((update, origin) => {
    if (origin === 'local') updates.push(update);
  });
  return updates;
}

function sync(from: YrsSession, to: YrsSession): void {
  to.applyUpdate(from.encodeStateAsUpdate(to.encodeStateVector()));
}

test('the worker paints what a fresh layout paints after random peer streams', async () => {
  const bytes = fixture();
  const seeder = await createYrsSession({ clientId: 0 });
  seeder.seedFromDocx(bytes);
  const seeded = seeder.encodeState();
  seeder.destroy();
  const open = async (clientId: number) => {
    const session = await createYrsSession({ clientId });
    const host = session.openDocx(bytes, false);
    session.loadState(seeded);
    return { session, host };
  };
  const { session: editor, host } = await open(103);
  const { session: a } = await open(101);
  const { session: b } = await open(102);
  a.beginUndoCapture();
  const fromA = recorder(a);
  const fromB = recorder(b);

  editor.clearFonts();
  const font = editor.registerFont(new Uint8Array(FONT));
  const request = JSON.stringify({
    ...buildResidentRegionLayoutRequest(host.document, 20, {}),
    measurement: {
      fontChains: { 'calibri|0|0': [font] },
      defaults: { fontSize: 11, fontFamily: 'Calibri' },
      authoritativeShaping: true,
    },
  });
  const extras = JSON.stringify({ fontChains: { 'calibri|0|0': [font] } });

  // The worker replica: every editor update reaches it, as the hook forwards
  // them; the editor's own system edits (renames, field re-reads) also go to
  // the peers, as the provider sends them to the room.
  const worker: ResidentEngineSession = await createResidentEngineSession();
  worker.loadState(editor.encodeState());
  const systemEdits: Uint8Array[] = [];
  let systemCount = 0;
  editor.onUpdate((update, origin) => {
    worker.applyUpdate(update);
    if (origin === 'local') {
      systemEdits.push(update);
      systemCount += 1;
    }
  });
  let painted: RetainedFrame = applyFrameDeltaOwned(
    null,
    decodeFrameDelta(worker.relayoutFrame(request, extras, 0))
  );
  expect(painted.pages.length).toBeGreaterThan(4);

  const fresh = (): RetainedFrame => {
    const session = createEditSession(900);
    try {
      session.open_docx(bytes, false);
      session.load(editor.encodeState());
      session.layout_document_with_regions_retained_json(request);
      return applyFrameDeltaOwned(null, decodeFrameDelta(session.build_display_list_frame(extras, 0)));
    } finally {
      session.free();
    }
  };

  const next = random(0x5eed);
  const pick = <T>(items: readonly T[]): T => items[Math.floor(next() * items.length)]!;
  // Random edits leave the field's two paragraphs to `fieldDelete`.
  const bodyParagraphs = (session: YrsSession) =>
    session
      .paragraphs('body')
      .filter(
        (paragraph) =>
          paragraph.text.length > 2 && paragraph.paraId !== FIELD_PARA && paragraph.paraId !== '45454545'
      );
  const type = (session: YrsSession) => {
    const paragraph = pick(bodyParagraphs(session));
    const offset = Math.floor(next() * paragraph.text.length);
    session.insertText({ story: 'body', paraId: paragraph.paraId, offset }, pick(['x', 'yz', ' w ']));
  };
  const remove = (session: YrsSession) => {
    const paragraphs = bodyParagraphs(session);
    const index = Math.floor(next() * (paragraphs.length - 1));
    const paragraph = paragraphs[index]!;
    const start = Math.floor(next() * (paragraph.text.length - 2));
    // One in three deletes runs into the next paragraph and merges it.
    const end =
      next() < 1 / 3
        ? { paraId: paragraphs[index + 1]!.paraId, offset: 1 }
        : { paraId: paragraph.paraId, offset: start + 2 };
    session.deleteRange({ story: 'body', start: { paraId: paragraph.paraId, offset: start }, end });
  };
  const split = (session: YrsSession, paraId: string, offset: number) =>
    session.splitParagraph({ story: 'body', paraId, offset });

  const ops = [
    'typeA',
    'typeB',
    'deleteA',
    'splitBoth',
    'fieldDelete',
    'undoA',
    'header',
    'pending',
  ] as const;
  const seen = new Set<string>();
  const systemByOp: Record<string, number> = {};
  let fieldDeleted = false;
  let pendingHeld = 0;
  let frames = 0;
  for (let round = 0; round < 32; round += 1) {
    const op = round < ops.length ? ops[round]! : pick(ops);
    seen.add(op);
    // Updates reach the editor in this order (pending: a dependant first).
    let order: Uint8Array[] = [];
    const take = (log: Uint8Array[]) => log.splice(0, log.length);
    if (op === 'typeA') {
      type(a);
      order = take(fromA);
    } else if (op === 'typeB') {
      type(b);
      order = take(fromB);
    } else if (op === 'deleteA') {
      remove(a);
      order = take(fromA);
    } else if (op === 'splitBoth') {
      // Two peers split one paragraph mid-text before seeing each other: the
      // editor renames the duplicate id (D8) and sends the rename.
      const paragraph = pick(bodyParagraphs(a));
      split(a, paragraph.paraId, 2);
      split(b, paragraph.paraId, 4);
      order = [...take(fromA), ...take(fromB)];
    } else if (op === 'fieldDelete') {
      // Two peers each delete half the field's last link: the editor re-reads
      // what the field shows and sends it (once).
      const drop = (session: YrsSession, from: number, to: number) =>
        session.deleteRange({
          story: 'body',
          start: { paraId: FIELD_PARA, offset: from },
          end: { paraId: FIELD_PARA, offset: to },
        });
      if (fieldDeleted) type(a);
      else {
        drop(a, 1, 2);
        drop(b, 2, 3);
        fieldDeleted = true;
      }
      order = [...take(fromA), ...take(fromB)];
    } else if (op === 'undoA') {
      if (!a.undo()) type(a);
      order = take(fromA);
    } else if (op === 'header') {
      a.insertText({ story: 'hf:rId20', paraId: '66666666', offset: 0 }, 'H');
      order = take(fromA);
    } else {
      // b types inside a's text before the editor has it, and b's update
      // arrives first: the editor holds it as pending structs until a's does.
      const paragraph = bodyParagraphs(a)[0]!;
      a.insertText({ story: 'body', paraId: paragraph.paraId, offset: 1 }, 'QQ');
      const first = take(fromA);
      for (const update of first) b.applyUpdate(update);
      take(fromB);
      b.insertText({ story: 'body', paraId: paragraph.paraId, offset: 2 }, 'P');
      const dependant = take(fromB);
      const shown = editor.paragraphs('body').map((paragraph) => paragraph.text);
      for (const update of dependant) editor.applyUpdate(update);
      if (JSON.stringify(editor.paragraphs('body').map((paragraph) => paragraph.text)) === JSON.stringify(shown))
        pendingHeld += 1;
      order = first;
    }
    const systemBefore = systemCount;
    for (const update of order) editor.applyUpdate(update);
    systemByOp[op] = (systemByOp[op] ?? 0) + systemCount - systemBefore;
    const frame = decodeFrameDelta(worker.relayoutFrame(request, extras, painted.frameEpoch));
    expect(frame.full).toBe(false);
    painted = applyFrameDeltaOwned(painted, frame);
    frames += 1;
    expect(painted.displayList).toEqual(fresh().displayList);

    // The room catches every replica up, the editor's system edits included.
    sync(a, b);
    sync(b, a);
    for (const update of take(systemEdits)) {
      a.applyUpdate(update);
      b.applyUpdate(update);
    }
    take(fromA);
    take(fromB);
    // The peers' own renames and re-reads reach the editor too.
    sync(a, editor);
    sync(b, editor);
    const caughtUp = decodeFrameDelta(worker.relayoutFrame(request, extras, painted.frameEpoch));
    expect(caughtUp.full).toBe(false);
    painted = applyFrameDeltaOwned(painted, caughtUp);
    expect(painted.displayList).toEqual(fresh().displayList);
  }
  expect([...seen].sort()).toEqual([...ops].sort());
  expect(frames).toBe(32);
  // The stream reached the editor's own system edits and held pending structs.
  expect(systemByOp.splitBoth).toBeGreaterThan(0);
  expect(systemByOp.fieldDelete).toBeGreaterThan(0);
  expect(pendingHeld).toBeGreaterThan(0);
  // Every replica holds the same document.
  const text = (session: YrsSession) =>
    session.storyIds().map((story) => session.paragraphs(story).map((paragraph) => `${paragraph.paraId}:${paragraph.text}`));
  expect(text(a)).toEqual(text(editor));
  expect(text(b)).toEqual(text(editor));
  // Two splits of one paragraph left no duplicate id behind.
  const ids = editor.paragraphs('body').map((paragraph) => paragraph.paraId);
  expect(new Set(ids).size).toBe(ids.length);
  worker.destroy();
  for (const session of [editor, a, b]) session.destroy();
}, 120_000);
