// Runs the DOCX matrices and compares every row's class with the committed baseline (see README.md).
//   bun shared/matrix/run.ts [--update-baseline] [--jobs N] [generator…]
// Fails when a row got worse, or when the generated rows changed, unless --update-baseline.
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { availableParallelism, tmpdir } from "node:os";
import { join } from "node:path";
import { worse, type Result } from "./lib";

const GENERATORS = ["docx-breaks", "docx-comments-breaks", "docx-fields", "docx-review-all", "docx-toc"];
const argv = process.argv.slice(2);
const update = argv.includes("--update-baseline");
const jobsAt = argv.indexOf("--jobs");
const jobs = jobsAt >= 0 ? Number(argv[jobsAt + 1]) : Math.max(1, Math.min(8, availableParallelism() - 1));
const picked = argv.filter((a, i) => !a.startsWith("--") && argv[i - 1] !== "--jobs");
const names = picked.length ? picked : GENERATORS;
for (const name of names) if (!GENERATORS.includes(name)) throw new Error(`unknown generator ${name}`);

const dir = mkdtempSync(join(tmpdir(), "docx-matrix-"));
const started = performance.now();
const busy = new Map<string, number>();
const tasks = names.flatMap((name) => Array.from({ length: jobs }, (_, shard) => ({ name, shard })));
async function worker() {
  for (let task = tasks.shift(); task; task = tasks.shift()) {
    const out = join(dir, `${task.name}.${task.shard}.jsonl`);
    const start = performance.now();
    const child = Bun.spawn(["bun", join(import.meta.dir, `${task.name}.matrix.ts`), "--shard", `${task.shard}/${jobs}`, "--out", out], {
      stdout: "inherit",
      stderr: "inherit",
    });
    if ((await child.exited) !== 0) throw new Error(`${task.name} shard ${task.shard} failed`);
    busy.set(task.name, (busy.get(task.name) ?? 0) + (performance.now() - start) / 1000);
  }
}
await Promise.all(Array.from({ length: jobs }, worker));
const seconds = Math.round((performance.now() - started) / 1000);

let failed = false;
const report: string[] = [];
for (const name of names) {
  const results = new Map<string, Result>();
  for (let shard = 0; shard < jobs; shard += 1)
    for (const line of readFileSync(join(dir, `${name}.${shard}.jsonl`), "utf8").split("\n").filter(Boolean)) {
      const result = JSON.parse(line) as Result;
      results.set(result.id, result);
    }
  // The baseline lists the rows that are not exact, under a header naming the row set: count and id hash.
  const ids = [...results.keys()].sort();
  const header = `# ${name}: ${ids.length} rows, ids ${createHash("sha256").update(ids.join("\n")).digest("hex").slice(0, 16)}; rows not listed are exact`;
  const file = join(import.meta.dir, "baseline", `${name}.tsv`);
  const [was, ...lines] = existsSync(file) ? readFileSync(file, "utf8").split("\n").filter(Boolean) : [""];
  const baseline = new Map(lines.map((line) => line.split("\t") as [string, string]));
  const counts = new Map<string, number>();
  const regressed: Result[] = [];
  const improved: string[] = [];
  for (const result of results.values()) {
    counts.set(result.cls, (counts.get(result.cls) ?? 0) + 1);
    const before = baseline.get(result.id) ?? "exact";
    if (worse(before, result.cls)) regressed.push({ ...result, detail: `was ${before}\n${result.detail ?? ""}` });
    else if (before !== result.cls) improved.push(`${result.id}: ${before} -> ${result.cls}`);
    if (result.detail) report.push(`${result.cls} ${name} | ${result.id}\n${result.detail}\n`);
  }
  const summary = [...counts].sort((a, b) => b[1] - a[1]).map(([cls, n]) => `${cls} ${n}`).join(", ");
  console.log(`\n${name}: ${results.size} rows in ${Math.round(busy.get(name) ?? 0)}s of shard time: ${summary}`);
  for (const r of regressed) console.log(`WORSE ${r.id}: now ${r.cls}\n${r.detail}\n`);
  for (const line of improved) console.log(`better ${line}`);
  const changed = was !== header;
  if (changed) console.log(`the rows changed: baseline "${was}", now "${header}"`);
  if (update) {
    mkdirSync(join(import.meta.dir, "baseline"), { recursive: true });
    const listed = [...results.values()].filter((r) => r.cls !== "exact").map((r) => `${r.id}\t${r.cls}`).sort();
    writeFileSync(file, `${[header, ...listed].join("\n")}\n`);
  } else if (regressed.length || changed) failed = true;
}
writeFileSync(join(dir, "report.txt"), report.join("\n"));
console.log(`\n${seconds}s with ${jobs} jobs; rows that are not exact: ${join(dir, "report.txt")}`);
if (update) console.log("baseline updated");
else if (failed) {
  console.log("FAILED: a row got worse, or the rows changed; rerun with --update-baseline once the change is intended");
  process.exit(1);
}
