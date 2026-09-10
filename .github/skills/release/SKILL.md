---
name: release
description: 'Publish @missing-elements/office-viewer to npm and create its matching GitHub release tag. Use when asked to release, publish, version, tag, or ship this package.'
argument-hint: 'Version to publish, such as 0.1.1'
---

# Release Office Viewer

Use this workflow only after the user explicitly asks to publish a release.

## Preconditions

1. Confirm the worktree is clean and `main` is at the intended release commit:

   ```bash
   git status --short
   git log -1 --oneline
   ```

2. Confirm the version in `package.json` is the requested version.
3. Verify npm authentication:

   ```bash
   npm whoami
   ```

4. Check whether that version is already published:

   ```bash
   npm view @missing-elements/office-viewer@<version> version
   ```

   An `E404` means the version is available. Any returned version means do not
   publish it again.

## Validation

Run the release gates before publishing:

```bash
pnpm test && pnpm build
npm pack --dry-run
```

The package tarball must contain only the intended `dist` output. Do not publish
if the worktree becomes dirty or a validation command fails.

## Publish

Publish the package publicly:

```bash
npm publish --access public
```

If npm reports `Scope not found` for `@missing-elements`, do not retry until the
user has created the npm scope or granted the authenticated account access. If
npm asks for browser authentication, let the user complete it, then retry the
publish command once.

## Tag And Push

Only after a successful publish, create and push an annotated tag for the same
version on the release commit:

```bash
git tag -a v<version> -m "Release v<version>"
git push origin main v<version>
```

Confirm the publish result and pushed tag in the final response.
