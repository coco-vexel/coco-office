# Contributing to Coco Office

Thank you for helping improve Coco Office. This repository contains public SDK
and editor packages; product hosts, agents, customer documents, and private test
corpora are intentionally out of scope.

## Before opening a change

- Search existing issues and pull requests first.
- Use an issue form for bugs, feature proposals, documentation problems, and
  questions.
- Open an issue before implementing a large feature, a new public API, or a
  compatibility-breaking change so the direction can be agreed first.
- Never upload customer documents, credentials, private source, or unlicensed
  fonts/images. Use minimal, generated fixtures.
- Report vulnerabilities privately as described in [SECURITY.md](SECURITY.md).

## Development setup

Use Node.js 22 or 24 and the pnpm version pinned in `package.json`.

```bash
corepack pnpm install --frozen-lockfile
corepack pnpm build
corepack pnpm test:ci
corepack pnpm example
corepack pnpm build:demo
corepack pnpm licenses:packages
corepack pnpm verify:packages
```

The portable CI suite covers Core package round-trips, selected Word and Excel
regressions, PPT SDK/docs contracts, all Editor unit tests, OOJSON conversion,
examples, demos, package-consumer checks, and SDK bundling. It is not the entire
Core test inventory or an Office application parity test.

For broader Core coverage, run:

```bash
corepack pnpm --filter @coco-vexel/coco-office test:all
```

Some tests build worker bundles. Host/COM/DocumentBuilder and private-corpus
checks are optional maintainer checks and must be reported separately.

## Architecture boundaries

- Keep Core independent of Editor, OOJSON, product hosts, and private packages.
- Editor projects Core data into browser views; it must not become a second
  document truth layer.
- Use public package imports instead of paths into sibling product repositories.
- Preserve unsupported document content where possible. Do not claim complete
  Office compatibility or pixel-identical rendering.
- Public API or behavior changes require focused tests and matching public docs.

## Branches and commits

Create a short-lived branch from `main`. Recommended names include
`feat/...`, `fix/...`, `docs/...`, and `chore/...`.

Use [Conventional Commits](https://www.conventionalcommits.org/), for example:

```text
feat(word): expose table metadata on paragraphs
fix(editor): preserve worksheet selection after export
docs(ppt): clarify unsupported animation behavior
```

Keep commits focused. Do not mix generated artifacts, unrelated formatting, or
private integration changes into a contribution.

## Pull requests

Pull requests should:

1. Explain the problem and the observable behavior change.
2. Link the relevant issue when one exists.
3. Include a focused regression test for document-behavior fixes.
4. List the checks actually run and any checks that were not run.
5. Update API documentation, examples, licenses, or notices when applicable.
6. Complete the repository pull request checklist.

`main` is protected. Changes normally require a pull request, one approving
review, resolved review conversations, and all required portable checks. New
commits dismiss stale approvals; the author of the latest push cannot provide
the final approval. Force-pushes and branch deletion are disabled on `main`.

Pull request titles should also follow Conventional Commits because squash
merges use the title as the resulting commit subject.

## Licensing

Contributions must be code you have the right to submit under the applicable
project license. Identify third-party sources and preserve their notices. See
[LICENSING.md](LICENSING.md) for license scope and release review requirements.
