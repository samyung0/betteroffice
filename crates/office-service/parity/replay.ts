// Replays calls recorded by record.ts through the native crate and compares
// the answers with the TS results.
//   bun crates/office-service/parity/replay.ts <recording dir> <office-service binary> [--report <file>] [--only <method>]
// Seeds and other bytes must be identical; JSON results equal ignoring key
// order; exports that differ as zips are compared part by part; rebased
// states that differ as bytes are compared as decoded Yjs content.
import { createHash } from "node:crypto";
import { readFileSync, readdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import * as Y from "yjs";
import { unzipContainer } from "../../../packages/docx/src/wasm/opc";

const [dir, binary] = process.argv.slice(2);
const option = (name: string) => {
  const at = process.argv.indexOf(name);
  return at < 0 ? undefined : process.argv[at + 1];
};
const reportPath = option("--report") ?? join(dir, "report.txt");
const only = option("--only");
const formats = option("--formats")?.split(",");
/** The call's Office format, when an argument names one. */
const formatOf = (call: Recorded): string | undefined => {
  if (call.method === "seedOffice") return String(call.args[0]);
  for (const arg of call.args)
    if (arg && typeof arg === "object" && !Array.isArray(arg) && typeof arg.format === "string") return arg.format;
  return undefined;
};
if (!dir || !binary) throw new Error("usage: replay.ts <dir> <binary> [--report file] [--only method]");

type Json = null | boolean | number | string | Json[] | { [key: string]: Json };
const blobs = option("--blobs") ?? join(dir, "blobs");
const blob = (sha: string) => new Uint8Array(readFileSync(join(blobs, sha)));
const sha = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex");

/** Recorded `{$blob}` to the wire's `{$bytes}`. */
function toWire(value: Json): Json {
  if (Array.isArray(value)) return value.map(toWire);
  if (value && typeof value === "object") {
    if (typeof value.$blob === "string") return { $bytes: Buffer.from(blob(value.$blob)).toString("base64") };
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, toWire(item)]));
  }
  return value;
}

/** Byte fields as their SHA-256, objects with sorted keys. */
function normal(value: Json): Json {
  if (Array.isArray(value)) return value.map(normal);
  if (value && typeof value === "object") {
    if (typeof value.$blob === "string") return `bytes:${value.$blob}`;
    if (typeof value.$bytes === "string") return `bytes:${sha(Buffer.from(value.$bytes, "base64"))}`;
    return Object.fromEntries(
      Object.keys(value)
        .sort()
        .map((key) => [key, normal(value[key]!)])
    );
  }
  return value;
}

const bytesOf = (value: Json): Uint8Array =>
  value && typeof value === "object" && !Array.isArray(value)
    ? typeof value.$blob === "string"
      ? blob(value.$blob)
      : new Uint8Array(Buffer.from(String(value.$bytes), "base64"))
    : new Uint8Array();

/** Why two exports differ part by part, or undefined when every part is equal. */
function partDifference(expected: Uint8Array, actual: Uint8Array): string | undefined {
  const a = unzipContainer(expected);
  const b = unzipContainer(actual);
  const names = [...new Set([...Object.keys(a), ...Object.keys(b)])].sort();
  const differing = names.filter((name) => !a[name] || !b[name] || sha(a[name]) !== sha(b[name]));
  if (!differing.length) {
    const order = JSON.stringify(Object.keys(a)) === JSON.stringify(Object.keys(b));
    return order ? undefined : "part order differs";
  }
  return differing
    .slice(0, 5)
    .map((name) => {
      if (!a[name] || !b[name]) return `${name}: only in ${a[name] ? "TS" : "native"}`;
      const [x, y] = [new TextDecoder().decode(a[name]), new TextDecoder().decode(b[name])];
      let at = 0;
      while (at < x.length && x[at] === y[at]) at++;
      return `${name}: at ${at}: TS ${JSON.stringify(x.slice(Math.max(0, at - 80), at + 120))}\n   native ${JSON.stringify(y.slice(Math.max(0, at - 80), at + 120))}`;
    })
    .join("\n  ");
}

/** A state's decoded Yjs content with ids, for states encoded differently. */
function yjsContent(state: Uint8Array): string {
  const { structs, ds } = Y.decodeUpdate(state);
  const items = structs.map((struct) => {
    const item = struct as Y.Item;
    return {
      id: [struct.id.client, struct.id.clock],
      length: struct.length,
      kind: struct.constructor.name,
      origin: item.origin ? [item.origin.client, item.origin.clock] : null,
      rightOrigin: item.rightOrigin ? [item.rightOrigin.client, item.rightOrigin.clock] : null,
      parent:
        typeof item.parent === "string" ? item.parent : item.parent ? JSON.stringify(item.parent) : null,
      parentSub: item.parentSub ?? null,
      content: item.content ? JSON.stringify(item.content.getContent?.() ?? null) : null,
    };
  });
  // Split runs at every unit so differently merged encodings compare equal.
  const units: string[] = [];
  for (const item of items) {
    for (let offset = 0; offset < item.length; offset++)
      units.push(JSON.stringify([item.id[0], item.id[1] + offset, item.kind]));
  }
  units.sort();
  const deleted: string[] = [];
  ds.clients.forEach((ranges, client) => {
    for (const range of ranges) for (let at = 0; at < range.len; at++) deleted.push(`${client}:${range.clock + at}`);
  });
  deleted.sort();
  const doc = new Y.Doc();
  Y.applyUpdate(doc, state);
  const roots: Record<string, unknown> = {};
  for (const [name, type] of doc.share) roots[name] = type.toJSON();
  return JSON.stringify({ units, deleted, roots: normal(roots as Json) });
}

interface Recorded {
  method: string;
  args: Json[];
  clients: number[];
  nowMs: number;
  perfMs?: number;
  value?: Json;
  error?: string;
}

const counts = new Map<string, Map<string, number>>();
const report: string[] = [];
const count = (method: string, outcome: string) => {
  const byOutcome = counts.get(method) ?? new Map<string, number>();
  byOutcome.set(outcome, (byOutcome.get(outcome) ?? 0) + 1);
  counts.set(method, byOutcome);
};

function compare(call: Recorded, answer: { value?: Json; error?: string; kind?: string }): [string, string?] {
  if (call.error !== undefined || answer.error !== undefined) {
    if (call.error === answer.error) return ["equal"];
    return ["differ", `TS error ${JSON.stringify(call.error)}\n  native ${answer.error !== undefined ? `${answer.kind} error ${JSON.stringify(answer.error)}` : `value ${JSON.stringify(normal(answer.value!)).slice(0, 400)}`}`];
  }
  let expected = call.value ?? null;
  let actual = answer.value ?? null;
  if (call.method === "officeReplicaStats" && expected && typeof expected === "object" && !Array.isArray(expected)) {
    const { wasmBytes: _wasm, ...rest } = expected;
    expected = rest;
  }
  if (call.method === "runtimeManifest") return ["skipped"];
  if (JSON.stringify(normal(expected)) === JSON.stringify(normal(actual))) return ["equal"];
  if (call.method === "exportOffice") {
    const difference = partDifference(bytesOf(expected), bytesOf(actual));
    return difference ? ["differ", difference] : ["parts-equal"];
  }
  if (call.method === "rebaseOffice" || call.method === "applyOfficeCommands") {
    const [e, a] = [expected as Record<string, Json>, actual as Record<string, Json>];
    const { state: es, ...erest } = e;
    const { state: as, ...arest } = a;
    if (JSON.stringify(normal(erest)) !== JSON.stringify(normal(arest)))
      return ["differ", `TS ${JSON.stringify(normal(erest)).slice(0, 2000)}\n  native ${JSON.stringify(normal(arest)).slice(0, 2000)}`];
    return yjsContent(bytesOf(es)) === yjsContent(bytesOf(as))
      ? ["content-equal"]
      : ["differ", "the states hold different Yjs content"];
  }
  return ["differ", `TS ${JSON.stringify(normal(expected)).slice(0, 3000)}\n  native ${JSON.stringify(normal(actual)).slice(0, 3000)}`];
}

const files = readdirSync(dir).filter((name) => name.startsWith("calls-") && name.endsWith(".jsonl")).sort();
const started = performance.now();
for (const file of files) {
  const calls = readFileSync(join(dir, file), "utf8")
    .split("\n")
    .filter(Boolean)
    .map((line) => JSON.parse(line) as Recorded);
  // One process per recorded process, so replica state follows the same calls.
  const child = Bun.spawn([binary, "serve"], { stdin: "pipe", stdout: "pipe", stderr: "inherit" });
  const reader = child.stdout.getReader();
  let buffered = "";
  const nextLine = async (): Promise<string> => {
    for (;;) {
      const newline = buffered.indexOf("\n");
      if (newline >= 0) {
        const line = buffered.slice(0, newline);
        buffered = buffered.slice(newline + 1);
        return line;
      }
      const { value, done } = await reader.read();
      if (done) throw new Error("office-service exited");
      buffered += new TextDecoder().decode(value);
    }
  };
  for (const call of calls) {
    if (only && call.method !== only && !["configureOfficeReplicas", "dropOfficeReplica"].includes(call.method)) continue;
    const format = formatOf(call);
    if (formats && format && !formats.includes(format)) continue;
    const request = {
      method: call.method,
      args: call.args.map(toWire),
      clients: call.clients,
      nowMs: call.nowMs,
      perfMs: call.perfMs,
    };
    child.stdin.write(`${JSON.stringify(request)}\n`);
    await child.stdin.flush();
    const answer = JSON.parse(await nextLine()) as { value?: Json; error?: string; kind?: string };
    const [outcome, detail] = compare(call, answer);
    count(call.method, outcome);
    if (detail) report.push(`${outcome} ${call.method} (${file})\n  ${detail}\n`);
  }
  child.stdin.end();
  await child.exited;
}
const lines = [...counts]
  .sort(([a], [b]) => a.localeCompare(b))
  .map(([method, outcomes]) => `${method}: ${[...outcomes].map(([outcome, n]) => `${outcome} ${n}`).join(", ")}`);
writeFileSync(reportPath, `${lines.join("\n")}\n\n${report.join("\n")}`);
console.log(lines.join("\n"));
console.log(`${files.length} recordings in ${Math.round((performance.now() - started) / 1000)} s; details: ${reportPath}`);
