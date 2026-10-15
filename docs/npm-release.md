# npm release procedure

All three packages are release targets at version `1.0.0`: `@coco-vexel/coco-office`, `@coco-vexel/coco-office-editor`, and `@coco-vexel/coco-office-oojson`. These are candidate versions, not a claim that they are available on npm. The workspace root stays private.

## Prepare and review

1. Confirm the final public repository URL; update repository/homepage/bugs metadata before packing if its address changes.
2. Confirm npm account membership/publish rights for the `@coco-vexel` scope, chosen versions and authentication. A registry 404 does not prove scope ownership. Do not write tokens into this repository.
3. From the exact reviewed commit run install with the frozen lockfile, build, portable tests, examples, demo build, `licenses:packages`, and `verify:packages`.
4. Review the three tarballs in `.release-check/packages/`, their SHA-256, export targets and notices. Use **pnpm pack** so workspace references become registry version ranges. The consumer verifier does this automatically.
5. Check each exact tarball with `npm publish PATH_TO_TARBALL --dry-run --ignore-scripts --access public`. This does not publish or prove authentication rights.
6. Run `corepack pnpm bundle:sdk`. Review the ZIP and SHA-256 in `.release-check/sdk/`; attach the exact ZIP to the matching GitHub Release so users can download the package, offline documentation and runnable examples together.

## Publish after explicit approval

Publish the reviewed tarballs in this order: Core, then Editor and OOJSON. Both depend on Core `^1.0.0`. Do not recursively publish the entire workspace.

```bash
npm publish .release-check/packages/coco-vexel-coco-office-1.0.0.tgz --access public
npm publish .release-check/packages/coco-vexel-coco-office-editor-1.0.0.tgz --access public
npm publish .release-check/packages/coco-vexel-coco-office-oojson-1.0.0.tgz --access public
```

Adjust filenames only when approved versions have changed, and regenerate/reverify their tarballs. If a version already exists, choose a new version and update dependents; do not attempt to replace it. Do not rebuild different bytes between review and publish.

After publication, install the exact versions from the public registry into a fresh directory, repeat consumer/demo verification and confirm npm package pages and distribution tags. The current local consumer test uses tarballs; it is not evidence of a successful registry publication.

GitHub source publication, npm publication and static Demo hosting require separate concrete approval. CI currently validates and never publishes automatically. Trusted publishing can be configured once the target repository and npm scope access are finalized.

References: [npm public scoped packages](https://docs.npmjs.com/creating-and-publishing-scoped-public-packages/), [npm scopes](https://docs.npmjs.com/using-npm/scope.html/).
