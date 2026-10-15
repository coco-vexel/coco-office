# Contributing

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
regressions, PPT SDK/docs contracts, all Editor unit tests and OOJSON conversion
tests. It is not the entire Core test inventory or an Office application parity
test. `pnpm --filter @coco-vexel/coco-office test:all` runs the broader suite; some tests
build worker bundles. Host/COM/DocumentBuilder and private-corpus checks are
optional maintainer checks and must be reported separately.

Keep Core independent of Editor, OOJSON and host APIs. Editor projects Core data
into browser views; it must not become a second document truth layer. Use public
package imports instead of paths into sibling product repositories.

For a pull request, describe the problem, observable behavior and checks run.
Use Conventional Commits. Add a focused regression when fixing document behavior.
Use generated, minimal fixtures; do not attach customer documents, credentials or
unlicensed fonts/images. Preserve unsupported document content where possible and
document limitations without claiming full Office compatibility.

Contributions must be code you have the right to submit under the applicable
project license. Identify third-party sources and preserve their notices. See
[LICENSING.md](LICENSING.md) for the license scope and release review requirements.
