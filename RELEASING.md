# Releasing

Changesets drive npm and crates.io releases through the same release PR.

## Changesets

Run `bun changeset` and select every affected npm package. Select
`@betteroffice/rust-crates` when a change affects the published Rust API or
implementation. Every Rust crate in `scripts/rust-crates.mjs` versions and
publishes in lockstep, independently from npm versions.

Merging a changeset opens or updates `chore: release`. Merging that release PR
publishes every unpublished npm and Cargo version. The Cargo publisher checks
crates.io before each upload, so rerunning a partial release resumes safely.

## Initial crates.io release

Trusted Publishing can only be configured after a crate exists, so a crate
that is not on crates.io yet cannot be created with OIDC. Every release
attempts to obtain an OIDC token via `rust-lang/crates-io-auth-action` (pinned to the
official v1 commit) and publishes per crate: a crate that does not exist yet
goes straight to `CRATES_IO_BOOTSTRAP_TOKEN`, while a crate that already
exists tries OIDC first and falls back to the bootstrap token only if the
OIDC token is unavailable or the publish fails. The choice is independent per crate.
`scripts/publish-crates.mjs` passes only the selected token to that crate's
`cargo publish` (under both `CARGO_REGISTRY_TOKEN` and
`CARGO_REGISTRIES_CRATES_IO_TOKEN`, with `CRATES_IO_BOOTSTRAP_TOKEN` removed
from every cargo child environment), rechecks the target version after a
failed OIDC upload before falling back, and records each bootstrap use in the
step summary before the registry wait for that crate can fail.
If a failed attempt is followed by a visible registry version, the summary
reports an unconfirmed publishing credential without attributing it to either
token. Only a successful Cargo publish confirms the credential used.

1. Create a short-lived crates.io token authorized to publish new crates and
   any existing crates that may need the fallback.
2. Add it to the repository as `CRATES_IO_BOOTSTRAP_TOKEN` before merging the
   release PR that introduces the new crate names.
3. Merge the release PR and confirm every crate was published.
4. Add a GitHub Trusted Publisher to each newly created crate with owner
   `openooxml`, repository `betteroffice`, and workflow `release.yml`, before
   its next release.
5. Remove the GitHub secret and revoke the bootstrap token once OIDC works
   for all crates, not just once all names exist. The release records each
   crate it creates with the bootstrap token and each existing crate it
   publishes with the bootstrap fallback separately in the step summary,
   recording each use before that crate's registry wait runs.

`scripts/check-publish-targets.mjs --crates` runs before the crates publish and
fails the release, naming every crate that is not on crates.io yet — unless
`CRATES_IO_BOOTSTRAP_TOKEN` is set, in which case it reports the missing
crates as "will be created with the bootstrap token" and passes.

`release.yml` always attempts the OIDC exchange on the publish path. That
step may continue on failure only when the bootstrap token is detected, so an
OIDC exchange that yields no token falls back per crate as above; without the
bootstrap token an OIDC failure still fails the release. A registry lookup
error fails the release instead of guessing a crate is new.

## Initial npm release of a new package

`release.yml` authenticates to npm with OIDC Trusted Publishing and carries no
`NPM_TOKEN` fallback. Trusted Publishing can only be configured on a package
that already exists, so **a brand-new package name cannot be published by the
workflow** — its first version must be pushed by hand.

Do this *before* merging a release PR that would publish the new name.
`changeset publish` publishes each package independently: a run where the new
package fails auth but its dependents succeed leaves those dependents live on
npm depending on a name that 404s.

1. Confirm the name is new and unclaimed: `npm view <name>` returns E404.
2. Build the publishable artifacts: `bun run build:packages`.
3. Pin workspace dependency ranges: `bun scripts/rewrite-workspace-deps.ts`.
   This publish-path-only script rewrites the source package manifests in
   place, so run it from a clean checkout and restore them after publishing.
4. From the package directory, publish the current version by hand with an
   account that has `publish` rights on the `@betteroffice` scope:
   `npm publish --access public --provenance`.
5. From the repository root, restore the rewritten source manifests with
   `git restore -- packages/*/package.json`, then confirm the worktree is clean.
6. Add a Trusted Publisher to the new package on npmjs.com with owner
   `openooxml`, repository `betteroffice`, and workflow `release.yml`.
7. Verify: `npm view <name> version` matches, and the package page shows the
   provenance attestation.
8. Only then merge the release PR. Subsequent versions publish through OIDC
   like every other package.

An optional peer that 404s does not make installation fail: npm silently omits
it. Consumers then fall back to synthetic metrics while being told to install a
package that does not exist. `scripts/check-publish-targets.mjs --npm` fails the
release, naming every package that is not on npm yet, so a skipped bootstrap
stops the run instead of half-publishing it. Both guards run before the first
upload of either registry — a guard that failed after the crates went out would
leave the release half-published, and a crate cannot be unpublished. The
versions it reads are the ones about to be published: the publish path runs only
when no changesets are pending, and `changesets/action` runs no `version`
command then.

## Native packages

`@betteroffice/{docx,pptx,xlsx}-native` are three roots with five platform packages
each (macOS arm64/x64, GNU Linux arm64/x64, Windows x64): 18 npm names.
Each root joins its core's fixed Changesets group. `version-packages.mjs` syncs
the platform manifests and the Cargo version.
Repository roots are private so `changeset publish` skips them. The dispatch-only
`publish-native.yml` drops `private` and uploads with OIDC Trusted Publishing in
environment `npm-<format>`, refusing `NPM_TOKEN`.
After a successful npm release, `release.yml`'s `native-npm` job dispatches that
workflow for formats whose root version is missing and waits for every run.
`RELEASE_SKIP_NATIVE=true` disables this job.

Trusted Publishing cannot create a package. Publish all 18 names by hand first,
repeating these steps for `docx`, `pptx` and `xlsx`:

1. Run `publish-native.yml` with `dry_run=true` at the release commit. Download its
   five `<format>-native.*` artifacts into `bindings/<format>-native/artifacts`
   and the `<format>-native-loader` artifact's `index.js` into
   `bindings/<format>-native` in a clean checkout of that commit.
2. Run `bun install --frozen-lockfile`, then `bunx napi artifacts --cwd bindings/<format>-native`.
3. From `bindings/<format>-native`, run `npm pkg delete private`, then
   `npm publish --access public --provenance=false`. Its `prepublishOnly` publishes the five platform packages first.
4. Configure all six names' Trusted Publisher: owner `openooxml`, repository
   `betteroffice`, workflow `publish-native.yml`, environment `npm-<format>`.
   Restrict each `npm-<format>` environment to the main branch.

## Standalone Cargo workspace

The standalone workspace `fuzz/` carries a `Cargo.lock` that records the
workspace crates by version, so every bump invalidates it and CI's
`cargo clippy --locked` would fail. `scripts/rust-crates.mjs` lists it in
`STANDALONE_WORKSPACES`; `scripts/version-packages.mjs` regenerates its lock
after writing the new versions, and the release commit carries it alongside the
manifests. `scripts/standalone-workspaces.test.ts` fails when a committed
lockfile is not on that list or a path dependency on a workspace crate carries a
`version`.

## Publish order

The workflow publishes dependencies before consumers, in the order
`RUST_CRATES` lists them in `scripts/rust-crates.mjs`: the shared crates, then
each format's own layers, then the crate that ties them together.

Cargo versions and internal registry requirements live in the root
`Cargo.toml`. `scripts/version-packages.mjs` synchronizes them with the private
Changesets marker in `crates/package.json`.
