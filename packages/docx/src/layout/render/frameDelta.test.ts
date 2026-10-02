import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { createEditSession, preloadEditWasm } from '../../wasm/edit';
import {
  applyFrameDelta,
  applyFrameDeltaOwned,
  decodeFrameDelta,
  displayPageRevision,
  displayPageShiftsSince,
  FRAME_DELTA_VERSION,
  type DecodedFrameDelta,
  type RetainedFrame,
} from './frameDelta';
import { createDisplayListQueries } from './displayListQueries';
import type { RustDisplayListQueryEngine } from './rustDisplayList';
import type { DisplayList, DisplayPage } from './displayList';

const WASM = resolve(import.meta.dir, '../../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);

describe('FrameDelta wire round-trip', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('decodes wasm-encoded full and delta frames to the equivalent JSON list', () => {
    const session = createEditSession(11);
    const { paraId } = JSON.parse(session.create_story('body', 'Hello frame', 'Normal', 'left'));
    const fontId = session.register_measure_font(new Uint8Array(readFileSync(FONT)));
    const request = JSON.stringify({
      bodyStory: 'body',
      regions: { sections: [{ sectionId: 'main', properties: {} }] },
      measurement: {
        fontChains: { 'calibri|0|0': [fontId] },
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        authoritativeShaping: true,
      },
      renderEnv: {},
    });

    const envelopeFor = (): string => {
      const output = JSON.parse(session.layout_document_with_regions_json(request)) as {
        measured: unknown;
        options: unknown;
        layout: unknown;
      };
      return JSON.stringify({
        measured: output.measured,
        options: output.options,
        layout: output.layout,
        fontChains: { 'calibri|0|0': [fontId] },
      });
    };

    const first = envelopeFor();
    const jsonList = JSON.parse(session.build_display_list_json(first)) as DisplayList;
    const fullFrame = session.build_display_list_frame(first, 0);
    const retained = applyFrameDelta(null, decodeFrameDelta(fullFrame));
    expect(retained.displayList).toEqual(jsonList);

    session.insert_text('body', paraId, 5, ' typed', undefined, undefined);
    const second = envelopeFor();
    const nextJsonList = JSON.parse(session.build_display_list_json(second)) as DisplayList;
    const deltaFrame = session.build_display_list_frame(second, retained.frameEpoch);
    const next = applyFrameDelta(retained, decodeFrameDelta(deltaFrame));
    expect(next.displayList).toEqual(nextJsonList);
    const pageText = next.displayList.pages[0].primitives
      .map((primitive) => ('text' in primitive ? (primitive.text ?? '') : ''))
      .join('');
    expect(pageText).toContain('typed');
  });

  it('records owned position shifts and ships them as query-store shift ops', () => {
    const session = createEditSession(13);
    const sentence = 'shift the following pages with enough text to fill several tiny pages. ';
    const { paraId } = JSON.parse(
      session.create_story('body', `start me. ${sentence.repeat(12)}`, 'Normal', 'left')
    );
    // split into paragraphs so trailing pages hold untouched blocks whose doc
    // positions merely shift when the first paragraph grows
    let splitId = paraId as string;
    for (let i = 0; i < 10; i++) {
      const receipt = JSON.parse(session.split_paragraph('body', splitId, 60, undefined, undefined)) as {
        secondParaId: string;
      };
      splitId = receipt.secondParaId;
    }
    const fontId = session.register_measure_font(new Uint8Array(readFileSync(FONT)));
    const request = JSON.stringify({
      bodyStory: 'body',
      regions: {
        sections: [
          {
            sectionId: 'main',
            properties: {
              pageWidth: 4320,
              pageHeight: 2880,
              marginTop: 300,
              marginRight: 300,
              marginBottom: 300,
              marginLeft: 300,
            },
          },
        ],
      },
      measurement: {
        fontChains: { 'calibri|0|0': [fontId] },
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        authoritativeShaping: true,
      },
      renderEnv: {},
    });
    const envelopeFor = (): string => {
      const output = JSON.parse(session.layout_document_with_regions_json(request)) as {
        measured: unknown;
        options: unknown;
        layout: unknown;
      };
      return JSON.stringify({
        measured: output.measured,
        options: output.options,
        layout: output.layout,
        fontChains: { 'calibri|0|0': [fontId] },
      });
    };

    const first = envelopeFor();
    const retained = applyFrameDeltaOwned(null, decodeFrameDelta(session.build_display_list_frame(first, 0)));
    expect(retained.displayList.pages.length).toBeGreaterThan(1);
    const trailingPage = retained.displayList.pages.at(-1)!;
    const revisionBefore = displayPageRevision(trailingPage);

    const updates: string[] = [];
    let nextHandle = 1;
    const engine: RustDisplayListQueryEngine = {
      hitTestRegionsJson: () => 'null',
      verticalMoveJson: () => 'null',
      rangeRectsJson: () => '[]',
      hasDisplayListSession: () => true,
      openDisplayList: () => nextHandle++,
      closeDisplayList: () => {},
      updateDisplayList: (_handle, update) => {
        updates.push(update);
      },
      hasDisplayListUpdate: () => true,
      rangeRectsByHandle: () => '[]',
      verticalMoveByHandle: () => 'null',
    };
    const firstQueries = createDisplayListQueries(retained.displayList, engine);
    firstQueries.prime();

    session.insert_text('body', paraId, 5, 'x', undefined, undefined);
    const second = envelopeFor();
    const next = applyFrameDeltaOwned(
      retained,
      decodeFrameDelta(session.build_display_list_frame(second, retained.frameEpoch))
    );

    // trailing pages absorb the insert as an in-place position shift with a
    // recorded, replayable run log
    expect(next.displayList.pages.at(-1)).toBe(trailingPage);
    expect(displayPageRevision(trailingPage)).toBe(revisionBefore + 1);
    const runLists = displayPageShiftsSince(trailingPage, revisionBefore);
    expect(runLists).not.toBeNull();
    expect(runLists!.length).toBe(1);
    expect(runLists![0].length).toBeGreaterThan(0);
    expect(displayPageShiftsSince(trailingPage, revisionBefore + 1)).toEqual([]);

    // handle adoption ships those shifts as compact ops instead of replacing
    // the page's serialized payload
    const secondQueries = createDisplayListQueries(next.displayList, engine, firstQueries);
    secondQueries.prime();
    expect(updates.length).toBe(1);
    const update = JSON.parse(updates[0]!) as {
      total: number;
      replace?: Array<[number, unknown]>;
      shift?: Array<[number, number, number[][][]]>;
    };
    const trailingIndex = next.displayList.pages.length - 1;
    expect(update.shift?.some(([to]) => to === trailingIndex)).toBe(true);
    expect(update.replace?.some(([to]) => to === trailingIndex)).toBeFalsy();
  });

  it('records owned shifts only after every run applies', () => {
    const page: DisplayPage = {
      pageIndex: 0,
      width: 100,
      height: 100,
      primitives: [
        {
          kind: 'text',
          text: 'x',
          x: 10,
          baselineY: 20,
          width: 10,
          font: '400 16px Calibri',
          color: '#000000',
        },
      ],
    };
    const previous: RetainedFrame = {
      protocolVersion: FRAME_DELTA_VERSION,
      docEpoch: 1,
      layoutEpoch: 1,
      frameEpoch: 1,
      pages: [
        {
          pageIndex: 0,
          pageId: 1n,
          fingerprint: 1n,
          primitiveIds: new BigUint64Array([1n]),
          page,
        },
      ],
      damagedPageIds: new Set(),
      removedPageIds: new Set(),
      displayList: { pages: [page] },
    };
    const delta: DecodedFrameDelta = {
      protocolVersion: FRAME_DELTA_VERSION,
      full: false,
      docEpoch: 1,
      layoutEpoch: 2,
      frameEpoch: 2,
      baseFrameEpoch: 1,
      pageCount: 1,
      operations: [
        {
          kind: 'shift-positions',
          pageIndex: 0,
          pageId: 1n,
          fingerprint: 2n,
          runs: [{ start: 0, count: 1, changedMask: 1, delta: 1 }],
        },
      ],
      bytes: new Uint8Array(),
    };

    expect(() => applyFrameDeltaOwned(previous, delta)).toThrow('requires retained docStart');
    expect(displayPageRevision(page)).toBe(0);
    expect(displayPageShiftsSince(page, 0)).toEqual([]);
  });
});

describe('FrameDelta typed value decoding', () => {
  // One full frame whose single page payload is an object with `extra` fields
  // appended after the required page fields.
  const frameWith = (extra: Array<[string, Uint8Array]>): Uint8Array => {
    const keys = ['pageIndex', 'width', 'height', 'primitives', ...extra.map(([key]) => key)];
    const strings = [...new Set(keys)];
    const u32 = (value: number) => new Uint8Array(new Uint32Array([value]).buffer);
    const f64 = (value: number) => new Uint8Array(new Float64Array([value]).buffer);
    const join = (parts: Uint8Array[]) => {
      const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
      let at = 0;
      for (const part of parts) (out.set(part, at), (at += part.length));
      return out;
    };
    const fields: Array<[string, Uint8Array]> = [
      ['pageIndex', join([Uint8Array.of(4), new Uint8Array(8)])],
      ['width', join([Uint8Array.of(5), f64(10)])],
      ['height', join([Uint8Array.of(5), f64(10)])],
      ['primitives', join([Uint8Array.of(7), u32(0), u32(0)])],
      ...extra,
    ];
    const body = join(fields.map(([key, value]) => join([u32(strings.indexOf(key)), value])));
    const payload = join([Uint8Array.of(8), u32(body.length), u32(fields.length), body]);
    const encoder = new TextEncoder();
    const table = join([
      u32(strings.length),
      ...strings.map((value) => join([u32(encoder.encode(value).length), encoder.encode(value)])),
    ]);
    const dataOffset = Math.ceil((128 + table.length) / 8) * 8;
    const total = dataOffset + payload.length;
    const frame = new Uint8Array(total);
    const view = new DataView(frame.buffer);
    frame.set([0x46, 0x44, 0x56, 0x31]);
    view.setUint16(4, FRAME_DELTA_VERSION, true);
    view.setUint16(6, 80, true);
    view.setUint32(8, total, true);
    view.setUint32(12, 1, true);
    view.setBigUint64(32, 1n, true);
    view.setUint32(48, 1, true);
    view.setUint32(52, 1, true);
    view.setUint32(56, 80, true);
    view.setUint32(60, 128, true);
    view.setUint32(64, table.length, true);
    view.setUint32(68, dataOffset, true);
    frame[80] = 1;
    view.setBigUint64(88, 1n, true);
    view.setUint32(108, dataOffset, true);
    view.setUint32(112, dataOffset, true);
    view.setUint32(116, payload.length, true);
    frame.set(table, 128);
    frame.set(payload, dataOffset);
    return frame;
  };
  const i64 = (value: bigint) => {
    const out = new Uint8Array(9);
    out[0] = 3;
    new DataView(out.buffer).setBigInt64(1, value, true);
    return out;
  };
  const u64 = (value: bigint) => {
    const out = new Uint8Array(9);
    out[0] = 4;
    new DataView(out.buffer).setBigUint64(1, value, true);
    return out;
  };
  const page = (frame: Uint8Array) => {
    const operation = decodeFrameDelta(frame).operations[0];
    if (operation.kind !== 'upsert') throw new Error('expected an upsert');
    return operation.page as unknown as Record<string, unknown>;
  };

  it('keeps a __proto__ key an own field and the prototype untouched', () => {
    const decoded = page(frameWith([['__proto__', i64(7n)]]));
    expect(Object.getPrototypeOf(decoded)).toBe(Object.prototype);
    expect(Object.getOwnPropertyDescriptor(decoded, '__proto__')?.value).toBe(7);
    expect(Object.keys(decoded)).toContain('__proto__');
  });

  it('decodes inherited names as fields and rejects repeated keys', () => {
    expect(page(frameWith([['constructor', i64(1n)]])).constructor).toBe(1);
    expect(() => page(frameWith([['note', i64(1n)], ['note', i64(2n)]]))).toThrow(
      'duplicate object key note'
    );
  });

  it('decodes integers up to the safe range and rejects beyond it', () => {
    const max = BigInt(Number.MAX_SAFE_INTEGER);
    const decoded = page(frameWith([['low', i64(-max)], ['high', u64(max)], ['neg', i64(-5n)]]));
    expect([decoded.low, decoded.high, decoded.neg]).toEqual([
      Number.MIN_SAFE_INTEGER,
      Number.MAX_SAFE_INTEGER,
      -5,
    ]);
    expect(() => page(frameWith([['x', i64(-max - 1n)]]))).toThrow(
      'signed value exceeds JavaScript safe-integer range'
    );
    expect(() => page(frameWith([['x', u64(max + 1n)]]))).toThrow(
      'unsigned value exceeds JavaScript safe-integer range'
    );
  });
});
