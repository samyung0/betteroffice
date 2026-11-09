import { readFileSync, writeFileSync } from 'node:fs';
import { updateDownloadSnapshots } from './download-snapshot.mjs';
import {
  RUST_CRATES,
  STANDALONE_WORKSPACES,
  WORKSPACE_MANIFEST,
  cargoMetadata,
  run,
  rustReleaseVersion,
  validateRustTrain
} from './rust-crates.mjs';

function cargoManifest(workspace) {
  return `${workspace}/Cargo.toml`;
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function workspaceVersion(source) {
  const section = source.match(/\[workspace\.package\]\n([\s\S]*?)(?=\n\[|$)/);
  const version = section?.[1].match(/^version = "([^"]+)"$/m)?.[1];
  if (!version) throw new Error('workspace.package.version is missing');
  return version;
}

function synchronizeCargoVersion(source, from, to) {
  if (workspaceVersion(source) !== from) {
    throw new Error(`Cargo release train does not match ${from}`);
  }

  let updated = source.replace(
    /(\[workspace\.package\]\n[\s\S]*?^version = ")[^"]+("$)/m,
    `$1${to}$2`
  );

  for (const crate of RUST_CRATES) {
    const key = escapeRegExp(crate.dependency);
    const pattern = new RegExp(`^(${key} = \\{[^\\n]*version = ")[^"]+("[^\\n]*\\})$`, 'm');
    if (!pattern.test(updated)) {
      throw new Error(`workspace dependency ${crate.dependency} has no version`);
    }
    updated = updated.replace(pattern, `$1${to}$2`);
  }

  return updated;
}

function validate(version, locked) {
  const metadata = cargoMetadata({ locked });
  validateRustTrain(metadata, version);
}

// Nothing else rewrites these lockfiles, which pin every bumped crate by version.
function synchronizeStandaloneLocks() {
  for (const workspace of STANDALONE_WORKSPACES) {
    cargoMetadata({ locked: false, manifestPath: cargoManifest(workspace) });
    // Re-assert under `--locked`, the way CI reads the lock it just wrote.
    cargoMetadata({ manifestPath: cargoManifest(workspace) });
  }
}

function synchronizeBunLock() {
  run('bun', ['install', '--lockfile-only']);
}

const checkOnly = process.argv.includes('--check');
const before = rustReleaseVersion();
const cargoBefore = readFileSync(WORKSPACE_MANIFEST, 'utf8');

if (workspaceVersion(cargoBefore) !== before) {
  throw new Error(`Rust changeset marker is ${before}, but Cargo is ${workspaceVersion(cargoBefore)}`);
}

if (checkOnly) {
  const simulated = synchronizeCargoVersion(cargoBefore, before, '999.999.999');
  if (workspaceVersion(simulated) !== '999.999.999') {
    throw new Error('Cargo release train version synchronization failed');
  }
  validate(before, true);
  for (const workspace of STANDALONE_WORKSPACES) {
    cargoMetadata({ manifestPath: cargoManifest(workspace) });
  }
  console.log(`Rust release train is synchronized at ${before}.`);
  for (const workspace of STANDALONE_WORKSPACES) {
    console.log(`${workspace}/Cargo.lock is current.`);
  }
  process.exit(0);
}

run('bun', ['run', 'changeset', 'version']);
try {
  const counts = await updateDownloadSnapshots();
  console.log(`Download snapshots: ${JSON.stringify(counts)}`);
} catch (error) {
  console.warn(`Download snapshots kept: ${error instanceof Error ? error.message : error}`);
}
const after = rustReleaseVersion();

if (after !== before) {
  writeFileSync(
    WORKSPACE_MANIFEST,
    synchronizeCargoVersion(readFileSync(WORKSPACE_MANIFEST, 'utf8'), before, after)
  );
  validate(after, false);
}

validate(after, true);
synchronizeStandaloneLocks();
synchronizeBunLock();
for (const workspace of STANDALONE_WORKSPACES) {
  console.log(`Synchronized ${workspace}/Cargo.lock.`);
}
console.log('Synchronized bun.lock.');
console.log(
  after === before
    ? `Rust release train remains at ${after}.`
    : `Synchronized Rust release train ${before} -> ${after}.`
);
