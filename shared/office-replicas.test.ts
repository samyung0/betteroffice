import { afterEach, expect, spyOn, test } from "bun:test";
import { readFile } from "node:fs/promises";
import * as Y from "yjs";
import {
  applyOfficeCommands,
  configureOfficeReplicas,
  dropOfficeReplica,
  exportOffice,
  officeReplicaStats,
  seedOffice,
  xlsxPendingEffects,
  type OfficeCheckpoint,
} from "./office-checkpoint";
import { XlsxDocument } from "../packages/xlsx/src/wasm/generated/xlsx_wasm.js";

const fixture = (name: string) =>
  readFile(new URL(`../apps/demo/public/${name}`, import.meta.url));
const ROOM = "source:file-1:epoch:1";

afterEach(() => configureOfficeReplicas(0));

/** Counters since the call before `run`. */
async function counted(run: () => Promise<unknown>) {
  const before = officeReplicaStats();
  await run();
  const after = officeReplicaStats();
  return {
    hits: after.hits - before.hits,
    misses: after.misses - before.misses,
    evictions: after.evictions - before.evictions,
  };
}

/** An editor replica's successive states: cell edits, a row insert, formatting and an Undo. */
async function xlsxStates(bytes: Uint8Array) {
  const seeded = await seedOffice("xlsx", bytes);
  const doc = XlsxDocument.openCollaborative(bytes, 7001);
  const states: OfficeCheckpoint[] = [];
  const take = () => states.push({ ...seeded, state: doc.encodeStateAsUpdate() });
  try {
    doc.applyUpdateJson(seeded.state);
    doc.editCellJson(JSON.stringify({ sheet: 0, row: 0, col: 0, input: "first" }));
    take();
    doc.applyOpsJson(
      JSON.stringify({ ops: [{ type: "insertRows", sheet: 0, at: 1, count: 2 }] })
    );
    doc.editCellJson(JSON.stringify({ sheet: 0, row: 4, col: 1, input: "=1+2" }));
    take();
    doc.patchRangeStyleJson(
      JSON.stringify({ sheet: 0, range: "A1:B2", patch: { bold: true } })
    );
    take();
    doc.undoJson();
    take();
  } finally {
    doc.free();
  }
  return { seeded, states };
}

/** What a save read before readers: a fresh replica applies the state. */
function applied(bytes: Uint8Array, state: Uint8Array) {
  const doc = XlsxDocument.openCollaborative(bytes, 7002);
  try {
    doc.applyUpdateJson(state);
    return JSON.parse(doc.pendingEffectsJson());
  } finally {
    doc.free();
  }
}

test("XLSX saves after the first reuse the room's replica, with the effects an apply reads", async () => {
  const bytes = await fixture("sample.xlsx");
  const { states } = await xlsxStates(bytes);
  const expected = states.map((state) => applied(bytes, state.state));
  expect(expected[1].length).toBeGreaterThan(expected[0].length);
  const uncached: unknown[] = [];
  for (const state of states) uncached.push(await xlsxPendingEffects(bytes, state));
  expect(uncached).toEqual(expected);
  configureOfficeReplicas(1 << 30);
  const cached: unknown[] = [];
  const counts = await counted(async () => {
    for (const state of states)
      cached.push(await xlsxPendingEffects(bytes, state, ROOM));
  });
  expect(cached).toEqual(expected);
  expect(counts).toEqual({ hits: 3, misses: 1, evictions: 0 });
  expect(await xlsxPendingEffects(bytes, states[3], ROOM)).toEqual(expected[3]);
});

test("a replica never holds a state, so an older one reads its own effects", async () => {
  const bytes = await fixture("sample.xlsx");
  const { states } = await xlsxStates(bytes);
  // Undoing the formatting only deletes: equal state vectors, other effects.
  expect(Y.encodeStateVectorFromUpdate(states[3].state)).toEqual(
    Y.encodeStateVectorFromUpdate(states[2].state)
  );
  configureOfficeReplicas(1 << 30);
  await xlsxPendingEffects(bytes, states[3], ROOM);
  const back = await counted(async () => {
    // The room reloaded from older checkpoints: the later edits are gone.
    for (const index of [2, 0])
      expect(await xlsxPendingEffects(bytes, states[index], ROOM)).toEqual(
        applied(bytes, states[index].state)
      );
  });
  expect(back).toEqual({ hits: 2, misses: 0, evictions: 0 });
});

test("a state that is not a whole workbook is applied, with or without a room", async () => {
  const bytes = await fixture("sample.xlsx");
  const seeded = await seedOffice("xlsx", bytes);
  const doc = XlsxDocument.openCollaborative(bytes, 7003);
  let delta: Uint8Array;
  try {
    doc.applyUpdateJson(seeded.state);
    const before = doc.encodeStateVector();
    doc.editCellJson(JSON.stringify({ sheet: 0, row: 0, col: 0, input: "delta" }));
    delta = doc.encodeDiff(before);
  } finally {
    doc.free();
  }
  const expected = applied(bytes, delta);
  expect(expected.length).toBe(1);
  const tail = { ...seeded, state: delta };
  expect(await xlsxPendingEffects(bytes, tail)).toEqual(expected);
  configureOfficeReplicas(1 << 30);
  expect(await xlsxPendingEffects(bytes, tail, ROOM)).toEqual(expected);
});

test("edits and exports never touch a replica", async () => {
  const bytes = await fixture("sample.xlsx");
  const { states } = await xlsxStates(bytes);
  configureOfficeReplicas(1 << 30);
  const expected = await xlsxPendingEffects(bytes, states[1]);
  await xlsxPendingEffects(bytes, states[1], ROOM);
  const untouched = await counted(async () => {
    await applyOfficeCommands(bytes, states[1], [
      { type: "set_cell", sheet: "0", cell: "C3", expectedValue: "", value: "agent" },
    ]);
    await exportOffice(bytes, states[1], {
      seed: "0".repeat(64),
      now: "2026-10-05T00:00:00.000Z",
    });
  });
  expect(untouched).toEqual({ hits: 0, misses: 0, evictions: 0 });
  expect(await xlsxPendingEffects(bytes, states[1], ROOM)).toEqual(expected);
});

test("unload, a new base and the byte budget drop replicas, an idle one before one in use", async () => {
  const sample = await fixture("sample.xlsx");
  const showcase = await fixture("showcase.xlsx");
  const sampleSeed = await seedOffice("xlsx", sample);
  const showcaseSeed = await seedOffice("xlsx", showcase);
  configureOfficeReplicas(1 << 30);
  await xlsxPendingEffects(sample, sampleSeed, ROOM);
  dropOfficeReplica(ROOM);
  expect(officeReplicaStats()).toMatchObject({ replicas: 0, replicaBytes: 0 });
  expect(await counted(() => xlsxPendingEffects(sample, sampleSeed, ROOM))).toMatchObject({ misses: 1 });
  // A publication moved the room to another base: the old replica goes.
  expect(await counted(() => xlsxPendingEffects(showcase, showcaseSeed, ROOM))).toMatchObject({ misses: 1 });
  expect(officeReplicaStats().replicas).toBe(1);
  // A budget that holds one replica: another room's replica is not kept while
  // this one is in use, and pushes it out once it has been idle long enough.
  const one = officeReplicaStats().replicaBytes;
  configureOfficeReplicas(one);
  const other = "source:file-2:epoch:1";
  expect(await counted(() => xlsxPendingEffects(showcase, showcaseSeed, other))).toEqual({ hits: 0, misses: 1, evictions: 0 });
  expect(await counted(() => xlsxPendingEffects(showcase, showcaseSeed, ROOM))).toMatchObject({ hits: 1 });
  const now = performance.now();
  const later = spyOn(performance, "now").mockReturnValue(now + 120_000);
  try {
    expect(await counted(() => xlsxPendingEffects(showcase, showcaseSeed, other))).toEqual({ hits: 0, misses: 1, evictions: 1 });
    expect(officeReplicaStats()).toMatchObject({ replicas: 1, replicaBytes: one });
    expect(await counted(() => xlsxPendingEffects(showcase, showcaseSeed, other))).toMatchObject({ hits: 1 });
  } finally {
    later.mockRestore();
  }
  // A replica larger than the whole budget is not kept.
  configureOfficeReplicas(1);
  expect(officeReplicaStats().replicas).toBe(0);
  await xlsxPendingEffects(showcase, showcaseSeed, ROOM);
  expect(officeReplicaStats().replicas).toBe(0);
});
