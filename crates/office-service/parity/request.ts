// Prints the wire request for one recorded call (by method and index in a file).
//   bun crates/office-service/parity/request.ts <dir> <file> <method> <index> > request.json
import { readFileSync } from "node:fs";
import { join } from "node:path";

const [dir, file, method, nth] = process.argv.slice(2);
const blob = (sha: string) => readFileSync(join(dir, "blobs", sha));
const wire = (value: any): any =>
  Array.isArray(value)
    ? value.map(wire)
    : value && typeof value === "object"
      ? value.$blob
        ? { $bytes: blob(value.$blob).toString("base64") }
        : Object.fromEntries(Object.entries(value).map(([key, item]) => [key, wire(item)]))
      : value;
const call = readFileSync(join(dir, file), "utf8")
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .filter((entry) => entry.method === method)[Number(nth)];
process.stdout.write(
  `${JSON.stringify({ method, args: call.args.map(wire), clients: call.clients, nowMs: call.nowMs, perfMs: call.perfMs })}\n`
);
