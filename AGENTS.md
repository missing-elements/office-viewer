# AGENTS.md

## Project

`@missing-elements/office-viewer` is a small ESM TypeScript library that
provides the headless `<office-viewer>` custom element. It orchestrates
`@silurus/ooxml` viewers for DOCX, XLSX, and PPTX; it is not a rendering UI.

Read `ARCHITECTURE.md` before changing lifecycle behavior or the public API.

## Repository Layout

- `src/office-viewer-element.ts`: component lifecycle, source normalization,
  format selection, and upstream viewer creation.
- `src/types.ts`: public TypeScript types.
- `src/index.ts`: package exports. Update and verify it when adding a public
  type or API.
- `tests/office-viewer-element.test.ts`: jsdom unit tests.
- `tests/browser/file-formats.test.ts`: Firefox browser tests using real Office
  fixtures in both `main` and `worker` modes.
- `public/fixtures/`: Office files used by browser tests.
- `demo/`: manual Vite demo, not a production UI.

## Design Constraints

- Keep this a thin orchestration layer. Do not add viewer UI, document editing,
  or wrappers that rename upstream APIs.
- Keep the Shadow DOM open and retain the `#viewer` container contract.
- Load format modules dynamically so consumers only fetch the selected format.
- Default to `mode: 'worker'`; preserve `main` as a fallback.
- Forward supported upstream options, such as `wasmUrl`, instead of duplicating
  upstream loader behavior. Hosted editors may need an absolute `wasmUrl`.
- Preserve load ownership: a superseded load must destroy only its own viewer
  and must not update the newer load's state.
- `load()` accepts URL strings, `ArrayBuffer`, `Blob`/`File`, and
  `ReadableStream<Uint8Array>`; normalize non-native upstream inputs before
  calling `viewer.load()`.
- If an unsupported source type is passed, reject the load promise and emit
  `loaderror` with a descriptive error before calling `viewer.load()`.

## Public API Expectations

- `load(source, { format, mode?, wasmUrl? })` starts a load and rejects on
  failure.
- On every load: clear the previous error, emit `loadstart`, and emit either
  `ready` or `loaderror` with `{ error }`.
- `reload()` repeats the retained source and options. `destroy()` cancels an
  active load, destroys the current upstream viewer, resets retained state, and
  emits `destroy`.
- Calling `destroy()` when no load has started or after a previous destroy()
  must be a safe no-op that still resets state without throwing.
- If `reload()` is called before any `load()` has succeeded, reject with an
  explicit error indicating that there is no prior load to repeat.
- The public instance surface is `load`, `reload`, `destroy`, `getViewer`, and
  the read-only `ready`, `error`, `format`, and `mode` properties. Keep it
  aligned with `ARCHITECTURE.md`, `README.md`, and `src/index.ts` exports.

## Development

Use pnpm. Do not edit generated `dist/` output.

```bash
pnpm typecheck
pnpm test:unit
pnpm test:browser
pnpm build
```

`pnpm test` runs unit and browser tests. Browser tests require Firefox and can
take longer because they load real fixture documents and Web Workers.

`pnpm test:build` is configured for build-output tests that are not currently
present; do not rely on it as a validation command until those tests exist.

For manual testing, run `pnpm dev`. Vite is configured for
`http://127.0.0.1:5173` and uses a strict port.

## Change Workflow

1. Make the smallest change at the owning layer; avoid unrelated refactors.
2. Add or update a focused test for behavior changes. Test all affected formats
   and both render modes when changing viewer creation, assets, or loading.
3. Run `pnpm typecheck` after TypeScript changes. Run `pnpm test:unit` for
  logic changes; run `pnpm test:browser` for changes to viewer creation,
  assets, or loading; run both if a change spans both areas. Then run
  `pnpm build` for package or bundling changes.
4. Keep `README.md` and `ARCHITECTURE.md` accurate when public behavior,
   supported sources, modes, or deployment requirements change.

## Style

- Follow the existing TypeScript style: no semicolons, single quotes, and
  concise comments only for non-obvious lifecycle or compatibility decisions.
- Prefer explicit types at public boundaries; avoid speculative abstractions.
- Do not modify unrelated user changes. Do not commit unless explicitly asked.