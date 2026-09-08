# Release Workflow Definition

## 1. Versioning
- Use Semantic Versioning (SemVer): `MAJOR.MINOR.PATCH`.
- Initial version is `0.0.0-spike`.
- Upon first public release, this must advance to `1.0.0`.

## 2. Build Steps (The Release Pipeline)
The following steps must be executed sequentially and validated at each stage:

### A. Pre-Release Build (Staging/Internal)
1.  **Cleanup:** Ensure all temporary build artifacts are cleared.
2.  **Build:** Run `pnpm run build` to generate the production bundle in `dist/`.
3.  **Asset Check:** Verify that all necessary assets (e.g., WASM, workers, fixtures) are present in `dist/assets/`.
4.  **Smoke Test:** Run essential integration tests against the *built* output (if possible, e.g., simulating loading `dist/demo/index.html`).
5.  **Measurements:** Generate and record the full bundle-size report. Update `docs/bundle-size-spike.md` with these metrics.

### B. Release Candidate (RC) Build
1.  **Versioning:** Bump the version number in `package.json` to `X.Y.Z-rc.N`.
2.  **Build:** Run `pnpm run build`.
3.  **Documentation:** Update `CHANGELOG.md` with the feature/fix set for this RC.
4.  **Verification:** Run a full set of tests, specifically targeting the architecture of the *next* planned feature set.

### C. Final Production Release (vX.Y.Z)
1.  **Versioning:** Bump the version number in `package.json` to `X.Y.Z`.
2.  **Changelog:** Update `CHANGELOG.md` with the final Release Notes.
3.  **Build:** Run `pnpm run build`.
4.  **Publish:** Execute `pnpm publish` (or equivalent deployment command).
5.  **Validation:** Immediately after publishing, run a basic smoke test against the live published artifacts to confirm CDN/module resolution works.

## 3. CDN/Plain Module Strategy
- When preparing for CDNs, verify that the asset paths (e.g., `dist/assets/...`) are relative to the module loading point (`<script type="module" src="...">`).
- For plain module consumption, ensure that the `package.json` `module` field correctly points to the compiled output bundle.

## 4. Hotfix Flow
- For urgent bug fixes:
    1.  Increment `PATCH` version in `package.json`.
    2.  Commit changes with a clear message referencing the fix.
    3.  Run build and publish immediately. No full regression test suite is required unless the fix touches core APIs.