import { describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { RUST_CRATES } from './rust-crates.mjs';
import {
  CONTAINER_IMAGES,
  publishedCommands,
  publishedCrates,
  publishedPackages,
  workspacePackages
} from './published-packages.mjs';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const CONTENT = join(ROOT, 'apps/docs/content/docs');

function mdx(directory: string): string[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) return mdx(path);
    return entry.name.endsWith('.mdx') ? [readFileSync(path, 'utf8')] : [];
  });
}

const DOCS = mdx(CONTENT).join('\n');

describe('docs cover every published package', () => {
  test('names every published npm package', () => {
    expect(publishedPackages().filter((name) => !DOCS.includes(name))).toEqual([]);
  });

  test('names every published crate', () => {
    expect(publishedCrates().filter((name) => !DOCS.includes(name))).toEqual([]);
  });
});

describe('docs name nothing that does not exist', () => {
  test('every npm package named is a package in this repository', () => {
    const real = new Set(workspacePackages());
    const claimed = [...DOCS.matchAll(/@betteroffice\/[a-z0-9-]+/g)].map((m) => m[0]);
    expect([...new Set(claimed)].filter((name) => !real.has(name))).toEqual([]);
  });

  test('every crate-shaped name is a registered crate, executable or container image', () => {
    const real = new Set([
      ...RUST_CRATES.map((crate) => crate.name),
      ...publishedCommands(),
      ...CONTAINER_IMAGES
    ]);
    const claimed = [...DOCS.matchAll(/\bbetteroffice(?:-[a-z0-9]+)+\b/g)].map((m) => m[0]);
    expect([...new Set(claimed)].filter((name) => !real.has(name))).toEqual([]);
  });
});
