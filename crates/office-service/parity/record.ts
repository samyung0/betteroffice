// Bun preload that records every call into shared/office-checkpoint.ts: its
// arguments, the client ids `randomInt` handed out during it, the frozen
// `Date.now()` and the result or error message, so replay.ts can run the same
// calls through the native crate.
//   PARITY_OUT=<dir> bun --preload ./crates/office-service/parity/record.ts <script>
//   PARITY_OUT=<dir> bun test --preload ./crates/office-service/parity/record.ts ./shared
import { plugin } from "bun";
import { AsyncLocalStorage } from "node:async_hooks";
import { createHash, randomInt as realRandomInt } from "node:crypto";
import { appendFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";

const out = process.env.PARITY_OUT;
if (!out) throw new Error("PARITY_OUT is required");
mkdirSync(join(out, "blobs"), { recursive: true });
const log = join(out, `calls-${process.pid}-${Date.now()}.jsonl`);
const self = import.meta.path.replaceAll("\\", "/");

const SYNC = ["configureOfficeReplicas", "dropOfficeReplica", "officeReplicaStats", "compareBaselines"];
const ASYNC = [
  "seedOffice",
  "officeBaseline",
  "compare",
  "rebaseOffice",
  "xlsxPendingEffects",
  "resolveAsset",
  "exportOffice",
  "inspectOffice",
  "applyOfficeCommands",
  "locateOfficeTargets",
  "runtimeManifest",
];

interface Call {
  clients: number[];
}
const calls = new AsyncLocalStorage<Call>();
let depth = 0;
let frozen = 0;
const realNow = Date.now;

export function randomInt(min: number, max: number): number {
  const value = realRandomInt(min, max);
  calls.getStore()?.clients.push(value);
  return value;
}

function blob(bytes: Uint8Array): { $blob: string } {
  const sha = createHash("sha256").update(bytes).digest("hex");
  const path = join(out!, "blobs", sha);
  if (!existsSync(path)) writeFileSync(path, bytes);
  return { $blob: sha };
}

function encode(value: unknown): unknown {
  if (value instanceof Uint8Array) return blob(value);
  if (Array.isArray(value)) return value.map(encode);
  if (value && typeof value === "object")
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, encode(item)]));
  return value;
}

function write(entry: Record<string, unknown>) {
  appendFileSync(log, `${JSON.stringify(entry)}\n`);
}

function enter(): { call: Call; nested: boolean; now: number; perf: number } {
  const nested = calls.getStore() !== undefined;
  const perf = performance.now();
  if (depth++ === 0) {
    frozen = realNow();
    Date.now = () => frozen;
  }
  return { call: { clients: [] }, nested, now: frozen, perf };
}

function leave() {
  if (--depth === 0) Date.now = realNow;
}

export function record<A extends unknown[], R>(name: string, fn: (...args: A) => R): (...args: A) => R {
  return (...args: A): R => {
    const { call, nested, now, perf } = enter();
    if (nested) {
      try {
        return fn(...args);
      } finally {
        leave();
      }
    }
    const done = (result: { value?: unknown; error?: string }) =>
      write({ method: name, args: encode(args), clients: call.clients, nowMs: now, perfMs: perf, ...result });
    let result: R;
    try {
      result = calls.run(call, () => fn(...args));
    } catch (error) {
      leave();
      done({ error: String((error as Error)?.message ?? error) });
      throw error;
    }
    if (result instanceof Promise)
      return result.then(
        (value) => {
          leave();
          done({ value: encode(value) });
          return value;
        },
        (error) => {
          leave();
          done({ error: String(error?.message ?? error) });
          throw error;
        }
      ) as R;
    leave();
    done({ value: encode(result) });
    return result;
  };
}

plugin({
  name: "office-checkpoint parity recorder",
  setup(build) {
    build.onLoad({ filter: /shared[\\/]office-checkpoint\.ts$/ }, ({ path }) => {
      let source = readFileSync(path, "utf8");
      const imports = 'import { createHash, randomInt } from "node:crypto";';
      if (!source.includes(imports)) throw new Error("office-checkpoint.ts imports changed");
      source = source.replace(
        imports,
        `import { createHash } from "node:crypto";\nimport { randomInt, record as __record } from "${self}";`
      );
      for (const name of [...SYNC, ...ASYNC]) {
        const declaration = new RegExp(`export (async )?function ${name}\\(`);
        if (!declaration.test(source)) throw new Error(`office-checkpoint.ts lost ${name}`);
        source = source.replace(declaration, (_match, async = "") => `${async}function __${name}(`);
        source += `\nexport const ${name} = __record(${JSON.stringify(name)}, __${name});`;
      }
      return { contents: source, loader: "ts" };
    });
  },
});
