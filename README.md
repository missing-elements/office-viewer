# office-viewer

Architecture spike for a standalone, read-only browser-based Office Open XML viewer built on `@silurus/ooxml`.

## Spike scope

This repository currently contains a technical spike, not the finished `<office-viewer>` custom element.

Included in this milestone:

- verified upstream API notes for `@silurus/ooxml`
- a raw integration harness under `/demo`
- representative DOCX, XLSX, and PPTX fixtures under `/public/fixtures`
- focused Vitest + WebdriverIO tests
- production-build asset checks and bundle-size notes

Not included yet:

- the polished public custom-element API
- common printing
- PDF support
- a public `workerUrl` option

## Install

```bash
corepack enable
corepack prepare pnpm@latest --activate
pnpm install
```

## Run the spike harness

```bash
pnpm dev
```

Then open:

- `http://127.0.0.1:5173/demo/`

The harness lets you:

- load sample DOCX/XLSX/PPTX fixtures by URL
- switch between `worker` and `main` mode
- load a local file
- destroy and reload the active viewer
- experiment with `wasmUrl` when you have a custom hosted WASM path

## Test

```bash
pnpm test
pnpm build
pnpm test:build
```

## Inspect the production build

```bash
pnpm preview
```

Then open:

- `http://127.0.0.1:4173/demo/`

## Notes and findings

- Verified API notes: [`docs/ooxml-integration-notes.md`](./docs/ooxml-integration-notes.md)
- Bundle and asset measurements: [`docs/bundle-size-spike.md`](./docs/bundle-size-spike.md)

## Known limitations

- The spike intentionally avoids the final custom-element API.
- The harness normalizes `File`, `Blob`, and `Uint8Array` inputs to `ArrayBuffer` because current upstream declarations type `load(source)` as `string | ArrayBuffer`.
- The spike verifies Vite 8 development/build behavior and a local static preview. It does not add custom worker hosting because upstream does not expose a public `workerUrl`.
- If a bundler does not preserve upstream `new URL(..., import.meta.url)` WASM resolution, copy the required `*_parser_bg.wasm` file into served assets and pass `wasmUrl`.
