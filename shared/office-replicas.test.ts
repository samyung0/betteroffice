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

test("XLSX saves after the first reuse the room's replica, with the uncached effects", async () => {
  const bytes = await fixture("sample.xlsx");
  const { states } = await xlsxStates(bytes);
  const uncached = [];
  for (const state of states) uncached.push(await xlsxPendingEffects(bytes, state));
  expect(uncached[1].length).toBeGreaterThan(uncached[0].length);
  configureOfficeReplicas(1 << 30);
  const cached = [];
  const counts = await counted(async () => {
    for (const state of states)
      cached.push(await xlsxPendingEffects(bytes, state, ROOM));
  });
  expect(cached).toEqual(uncached);
  expect(counts).toEqual({ hits: 3, misses: 1, evictions: 0 });
  // The same state again applies nothing and reads the same effects.
  expect(await xlsxPendingEffects(bytes, states[3], ROOM)).toEqual(uncached[3]);
});

test("a state that lacks what the replica applied reopens, so a discarded save never leaks", async () => {
  const bytes = await fixture("sample.xlsx");
  const { states } = await xlsxStates(bytes);
  configureOfficeReplicas(1 << 30);
  await xlsxPendingEffects(bytes, states[2], ROOM);
  // The room reloaded from an older checkpoint: the later edits are gone.
  const reloaded = await counted(async () =>
    expect(await xlsxPendingEffects(bytes, states[0], ROOM)).toEqual(
      await xlsxPendingEffects(bytes, states[0])
    )
  );
  expect(reloaded).toMatchObject({ hits: 0, misses: 1 });
});

test("an Undo that adds no struct still counts: the state before it does not hold the replica", async () => {
  const bytes = await fixture("sample.xlsx");
  const { states } = await xlsxStates(bytes);
  // Undoing the formatting only deletes, so only the delete sets tell the two states apart.
  expect(Y.encodeStateVectorFromUpdate(states[3].state)).toEqual(
    Y.encodeStateVectorFromUpdate(states[2].state)
  );
  configureOfficeReplicas(1 << 30);
  await xlsxPendingEffects(bytes, states[3], ROOM);
  const back = await counted(async () =>
    expect(await xlsxPendingEffects(bytes, states[2], ROOM)).toEqual(
      await xlsxPendingEffects(bytes, states[2])
    )
  );
  expect(back).toMatchObject({ hits: 0, misses: 1 });
});

test("edits and exports never touch a replica, and the replica stays the state it applied", async () => {
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
