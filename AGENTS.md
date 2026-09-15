# AGENTS.md

## Project

`@missing-elements/office-viewer` is a small ESM TypeScript library that
provides the headless `<office-viewer>` custom element. It orchestrates
`@silurus/ooxml` viewers for DOCX, XLSX, and PPTX; it is not a rendering UI.

Read `ARCHITECTURE.md` before changing lifecycle behavior or the public API.

## Repository Layout

- `src/office-viewer-element.ts`: component lifecycle, source normalization,
  format selection, and upstream viewer creation.
- `src/types.ts`: public TypeScript types, the `OfficeViewer` instance union,
  and `OFFICE_FORMATS`, the single source for the format list.
- `src/index.ts`: package exports. Update and verify it when adding a public
  type or API.
- `tests/office-viewer-element.test.ts`: jsdom unit tests for the lifecycle
  contract. They mock the three `@silurus/ooxml` entrypoints with
  `tests/helpers/fake-viewer.ts`, a controllable stand-in whose `load()` can be
  held, finished, or failed. It mirrors upstream's worst case: it detaches the
  `ArrayBuffer` it is given and does not settle an in-flight `load()` on
  `destroy()`.
- `tests/browser/file-formats.test.ts`: Firefox browser tests using real Office
  fixtures in both `main` and `worker` modes.
- `public/fixtures/`: Office files used by browser tests.
- `demo/`: manual Vite demo, not a production UI. It must only use the public
  element API.

## Design Constraints

- Keep this a thin orchestration layer. Do not add viewer UI, document editing,
  or wrappers that rename upstream APIs.
- Keep the Shadow DOM open and retain the `#viewer` container contract.
- Load format modules dynamically so consumers only fetch the selected format.
- Default to `mode: 'worker'`; preserve `main` as a fallback.
- Forward supported upstream options, such as `wasmUrl`, instead of duplicating
  upstream loader behavior. Hosted editors may need an absolute `wasmUrl`; the
  `wasm-url` attribute exists for declarative use.
- Preserve load ownership: a newer `load()`, `destroy()`, or removal from the
  document cancels the in-flight load immediately (stream read cancelled, viewer
  destroyed, nothing mounted afterwards). Race every await against the abort
  signal: upstream does not settle `load()` when its viewer is destroyed. The
  cancelled promise rejects with a `DOMException` named `AbortError`, emits
  nothing, and never updates element state.
- Validate the format before importing a module; never construct a viewer for
  an invalid format or source.
- `load()` accepts URL strings, `ArrayBuffer`, `Blob`/`File`, and
  `ReadableStream<Uint8Array>`; normalize non-native upstream inputs before
  calling `viewer.load()`. Upstream may detach the `ArrayBuffer` it receives, so
  never retain it: keep URL strings and `Blob`/`File` as-is and keep a `Blob`
  copy for `ArrayBuffer` and stream sources.
- Coalesce attribute changes into one load per microtask. An explicit `load()`
  outranks attribute changes made earlier in the same task. Clearing `src`
  unloads only an attribute-driven document.
- On removal from the document, release the viewer and in-flight load after one
  microtask (so synchronous moves keep the viewer) but keep the retained
  request; reconnecting restores it unless attributes changed while detached.
  Moves without attribute changes and failed loads never start work.
- Do not read upstream private fields. Upstream scroll viewers expose no public
  `mode`; the element reports the mode it created the viewer with.

## Public API Expectations

- `load(source, { format, mode?, wasmUrl? })` starts a load and rejects on
  failure. Options are shallow-copied.
- On every load: clear the previous error, emit `loadstart`, and then emit
  either `ready` or `loaderror` with `{ error }`, unless the load is cancelled.
- An unsupported format or source type rejects the load promise and emits
  `loaderror` with a descriptive error before any viewer is created.
- `ready`, `getViewer()`, `format`, and `mode` describe the committed viewer
  and change together on `ready`. `mode` is the mode the viewer was created with
  (`worker` when omitted). `error` describes the most recent owning load.
- `reload()` repeats the most recent request, including a failed one. If no
  request has been made since the last `destroy()`, reject with an explicit
  error.
- `destroy()` cancels an active load, destroys the current upstream viewer,
  forgets the retained request, resets state, and emits `destroy`. Calling it
  when nothing is loaded or after a previous `destroy()` is a safe no-op that
  still resets state and emits `destroy`.
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
2. Add or update a focused test for behavior changes. Lifecycle, ownership,
   attribute, and source-normalization behavior belongs in the jsdom unit tests
   with the fake viewer. Test all affected formats and both render modes in the
   browser tests when changing viewer creation, assets, or loading against the
   real upstream.
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
