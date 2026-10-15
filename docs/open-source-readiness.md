# Initial public release gate

This document describes the gate, not a claim that publication has been approved.
Do not publish npm packages, push release tags, enable hosting or change repository
visibility as part of the local checks.

## Rights and public contents

- The maintainer confirmed personal ownership of Core/Editor/OOJSON code and
  that the three XLSX fixtures were self-generated; see `LICENSING.md`.
- Use reviewed XLSX copies in the public candidate. Private Git history still
  contains original author/account metadata. A private npm flag does not exclude
  a directory from GitHub visibility.
- Review all branches/tags and commit history, not just HEAD; run a redacted
  credential scanner. Also review fixture contents and document metadata.
- Review existing GitHub issues, PRs, releases, Actions logs/artifacts and other
  repository surfaces before changing visibility. Local Git scanning does not
  cover these surfaces.
- Retain historical attribution and dependency license notices. If history
  cannot be cleared, prepare an approved clean public snapshot without rewriting
  the existing private repository.

## Reproducible candidate

`node scripts/create-standalone-snapshot.mjs` exports a reviewed public allowlist
from the exact current `HEAD` into a fresh system temporary directory, without
Git metadata, `node_modules`, sibling repositories or internal planning notes.
Its path, source commit and excluded-file count are recorded in
`.release-check/snapshot.json`. The command refuses to run when tracked files are
dirty, so inspect and commit the selected changes before creating the candidate.

From the snapshot, run:

```bash
corepack pnpm install --frozen-lockfile
corepack pnpm build
corepack pnpm test:ci
corepack pnpm --filter @coco-vexel/coco-office test:types:excel
corepack pnpm example
corepack pnpm build:demo
corepack pnpm licenses:packages
corepack pnpm verify:packages
corepack pnpm bundle:sdk
```

The portable CI matrix targets Node 22/24 on Linux/Windows. A local Windows run
does not certify that matrix. Report optional corpus tests skipped separately.
Review browser open/edit/export behavior using the synthetic examples.

`verify:packages` packs Core/Editor/OOJSON, installs the tarballs into a fresh npm consumer,
checks package export targets and notices, runs ESM/CJS and Node worker checks,
verifies three-format round-trips, then builds the browser demo and compiles a
consumer TypeScript example. It needs network access to the public npm registry.
Temporary consumer paths and tarballs are recorded under `.release-check/`.

`bundle:sdk` creates a versioned ZIP under `.release-check/sdk/` containing the
Core package tarball, public integration guides, runnable examples, license and
third-party notices. It also records the ZIP's SHA-256 digest. This bundle is the
offline handoff artifact and should be attached to the corresponding GitHub
Release after publication approval.

## Package and publication review

- Pack Core, Editor and OOJSON with **pnpm pack** so workspace dependencies become version
  ranges. Inspect the tarballs and verify every export and type target exists.
- Install all three tarballs into a fresh consumer project with no workspace aliases.
  Check ESM/CJS, browser bundling, Node workers, and Word/Excel/PPT round-trips.
- Include LICENSE, package README and third-party notices in all three packages;
  inspect sourcemaps and package contents for private paths or unintended files.
- Verify npm scope access and whether the selected versions already exist before
  deciding release versions. No automatic version bump or overwrite is assumed.
- Enable and verify private vulnerability reporting; configure branch protection
  and maintainer access. Confirm where SDK issues should be filed.
- Present the exact commit, package hashes, checks and unresolved findings for
  maintainer approval. Publication is a separate authorized step.
