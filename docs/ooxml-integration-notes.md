# `@silurus/ooxml` integration notes

Verified against `@silurus/ooxml@0.85.3` using the published npm package declarations and README, plus local Vite/browser experiments in this repository.

## Sources checked

- npm package: `node_modules/@silurus/ooxml/package.json`
- declarations:
  - `node_modules/@silurus/ooxml/dist/types/docx.d.ts`
  - `node_modules/@silurus/ooxml/dist/types/xlsx.d.ts`
  - `node_modules/@silurus/ooxml/dist/types/pptx.d.ts`
- README: `node_modules/@silurus/ooxml/README.md`
- internal worker host evidence:
  - `node_modules/@silurus/ooxml/dist/render-worker-host-*.js`

## Constructors and load signatures

| Format | Headless engine | Viewers | Typed load input |
| --- | --- | --- | --- |
| DOCX | `DocxDocument.load(source, opts?)` | `new DocxViewer(canvas, opts?)`, `new DocxScrollViewer(container, opts?)`, `fromDocument(...)` | `string \| ArrayBuffer` |
| XLSX | `XlsxWorkbook.load(source, opts?)` | `new XlsxViewer(container, opts?)`, `new XlsxSheetViewer(canvas, opts?)`, `fromWorkbook(...)` | `string \| ArrayBuffer` |
| PPTX | `PptxPresentation.load(source, opts?)` | `new PptxViewer(canvas, opts?)`, `new PptxScrollViewer(container, opts?)`, `fromPresentation(...)` | `string \| ArrayBuffer` |

Relevant declarations:

- DOCX: `dist/types/docx.d.ts:2185`, `:2289-2290`, `:2447-2448`
- XLSX: `dist/types/xlsx.d.ts:2053`, `:2494-2506`
- PPTX: `dist/types/pptx.d.ts:1929`, `:2031-2033`, `:2189-2191`

## Supported source types

### Verified upstream typing

Current declarations type all `load(...)` and `*.load(...)` entry points as:

```ts
string | ArrayBuffer
```

### README behavior vs declarations

The README shows `await viewer.load(file)` examples, but the published declarations do not type `File`, `Blob`, or `Uint8Array` directly.

Spike conclusion:

- URL strings are safe direct inputs.
- `URL` objects should be converted to strings.
- `File`, `Blob`, and `Uint8Array` should be normalized to a copied `ArrayBuffer`.
- caller-owned buffers should be copied before worker handoff to preserve ownership semantics.

This is the largest practical discrepancy between the planning docs and the current published typings.

## Viewer behavior

### `DocxScrollViewer`

- container-based continuous scroll viewer
- virtualizes pages and recycles canvases
- requires bounded container height
- exposes `pageCount`, `layoutComplete`, `waitUntilLayoutComplete()`, `topVisiblePage`
- callback: `onVisiblePageChange(topIndex, total, layoutComplete)`
- supports `setScale`, `fitWidth`, `fitPage`, `findText`, `findNext`, `findPrev`, `clearFind`, `destroy`

Declarations: `dist/types/docx.d.ts:2353-2549`

### `XlsxViewer`

- container-based workbook/grid viewer
- owns its own canvas, scroll surface, and tab UI
- XLSX remains sheet/grid-specific; not page-based
- exposes `sheetIndex`, `sheetCount`, `sheetNames`, `goToSheet`, `nextSheet`, `prevSheet`
- exposes `setViewportOffset`, `scrollToCell`, `setSelection`, `copySelection`
- callbacks include `onReady(sheetNames)` and `onSheetChange(index, total)`
- supports `setScale`, `fitWidth`, `fitPage`, `findText`, `findNext`, `findPrev`, `clearFind`, `destroy`

Declarations: `dist/types/xlsx.d.ts:2175-2540`

### `PptxScrollViewer`

- container-based continuous slide scroller
- virtualizes slides and recycles canvases
- requires bounded container height
- exposes `slideCount`, `availableSlideCount`, `layoutComplete`, `waitUntilLayoutComplete()`, `topVisibleSlide`
- callback: `onVisibleSlideChange(topIndex, total, layoutComplete)`
- supports `scrollToSlide`, `goToComment`, `setScale`, `fitWidth`, `fitPage`, `findText`, `findNext`, `findPrev`, `clearFind`, `destroy`

Declarations: `dist/types/pptx.d.ts:2103-2299`

## Worker mode

- all three formats accept `mode?: 'main' | 'worker'`
- worker mode is real upstream behavior, not a wrapper invention
- worker mode uses worker-side parse/layout/paint and bitmap presentation on the main thread
- main-thread canvas render methods are not the worker-friendly APIs for headless use; use `renderPageToBitmap`, `renderViewportToBitmap`, or `renderSlideToBitmap`
- PPTX and XLSX explicitly require `Worker` and `OffscreenCanvas` support
- DOCX can effectively fall back to main mode for browser-only shaping; the engine `mode` getter is the reliable place to inspect the effective mode

Important integration conclusion:

- the viewer classes do not expose a stable public `mode` getter
- the engine classes do expose `mode`
- recording effective mode should therefore happen at the engine layer, then pass the loaded engine into `fromDocument` / `fromWorkbook` / `fromPresentation`

Relevant declarations:

- DOCX load options: `dist/types/docx.d.ts:2137-2147`
- XLSX load options: `dist/types/xlsx.d.ts:2022-2024`
- PPTX load options: `dist/types/pptx.d.ts:1868-1871`

README worker notes: `README.md:238-282`

## `wasmUrl` and asset resolution

`wasmUrl` is a real supported upstream option:

```ts
wasmUrl?: string | URL
```

Present in all three declaration files:

- DOCX: `dist/types/docx.d.ts:1177`
- XLSX: `dist/types/xlsx.d.ts:1181`
- PPTX: `dist/types/pptx.d.ts:1384`

Verified behavior:

- the package ships parser WASM as separate assets
- upstream README documents `wasmUrl` as the escape hatch for esbuild/Angular-style builds
- Vite 8 worked without extra configuration in this spike
- the spike production build emitted:
  - `docx_parser_bg-*.wasm`
  - `xlsx_parser_bg-*.wasm`
  - `pptx_parser_bg-*.wasm`

## Worker assets

Verified upstream package exports do **not** expose a public worker URL contract.

Observed behavior:

- internal `render-worker-host-*.js` files create workers with:
  - `new Worker(new URL('assets/render-worker-*.js', import.meta.url), { type: 'module' })`
- worker hosting is therefore upstream-internal and bundler-managed

Conclusion:

- do **not** document or expose a public `workerUrl`
- document only observed asset behavior

## Destroy and ownership semantics

Two supported modes exist upstream:

1. **Self-loaded viewer**
   - `new Viewer(...); await viewer.load(source)`
   - viewer owns the engine
   - `destroy()` tears down the owned engine

2. **Borrowed engine**
   - `const engine = await Engine.load(...)`
   - `const viewer = Viewer.fromDocument/fromWorkbook/fromPresentation(...)`
   - viewer borrows the engine
   - `destroy()` does **not** destroy the borrowed engine
   - caller must destroy the engine separately

This is the right pattern for this spike because it lets us:

- inspect `engine.mode`
- normalize source inputs before load
- destroy/reload deterministically

README ownership notes: `README.md:444-468`

## Navigation, zoom, find, hyperlinks, resize, layout completion

Common verified themes:

- zoom: `getScale`, `setScale`, `zoomIn`, `zoomOut`, `fitWidth`, `fitPage`
- find: `findText`, `findNext`, `findPrev`, `clearFind`
- hyperlinks:
  - `enableHyperlinks?: boolean`
  - `onHyperlinkClick?: (target: HyperlinkTarget) => void`
- resize:
  - DOCX/PPTX scroll viewers support `refitOnResize`
  - cross-format helper `autoResize(...)` is exported

Format-specific:

- DOCX single-page viewer: `goToPage`, `nextPage`, `prevPage`
- DOCX scroll viewer: `topVisiblePage`
- XLSX workbook viewer: `goToSheet`, `scrollToCell`, selection APIs
- PPTX single-slide viewer: `goToSlide`, `nextSlide`, `prevSlide`
- PPTX scroll viewer: `scrollToSlide`, `topVisibleSlide`

Layout completion:

- DOCX: `layoutComplete`, `waitUntilLayoutComplete()`
- PPTX: `layoutComplete`, `availableSlideCount`, `waitUntilLayoutComplete()`
- XLSX does not expose the same progressive layout contract

## Stable errors and resource limits

### Stable typed errors

- `OoxmlError`
  - code union includes:
    - `'encrypted'`
    - `'invalid-password'`
    - `'unsupported-encryption'`
    - `'legacy-binary-format'`
    - `'not-ooxml'`
- `OoxmlResourceLimitError`
  - `code === 'ooxml-resource-limit'`
- `OoxmlDecodedImageLimitError`
  - `code === 'ooxml-decoded-image-limit'`
- `TiffDecodeError`
  - `code === 'ooxml-tiff-decode'`

### Important caveat

The README explicitly says not every configuration/fetch/parser/worker/media failure has a stable code. During this spike, malformed input reliably produced a useful diagnostic, but not every malformed path surfaced a typed `not-ooxml` code.

Resource-limit notes:

- `resourceLimits` is the current preferred option
- `maxZipEntryBytes` remains as a compatibility alias
- `getResourceMetrics()` exists on engines and viewers

README error notes: `README.md:960-1038`

## Optional renderer injection

Load options include optional injected renderers:

- `math`
- `threeD`
- `regionMap`
- `chartEx`
- `tiff`

Observed conclusions:

- no optional renderer is needed for this spike
- keep them out of the default graph
- built-in optional modules are documented to work in both modes when explicitly imported
- custom renderer objects are not worker-transferable and use upstream fallback behavior in `mode: 'worker'`

## Bundler and static-hosting behavior

Verified in this spike:

- Vite 8 dev/build: works
- Vite preview production build: works
- production output preserved relative asset references in `dist/demo/index.html`
- sample fixtures, worker assets, and parser WASM were all emitted as separate files

Upstream README claims zero-config support for:

- webpack 5
- Next.js/Turbopack
- Vite 8
- Vite 7 production
- plain `<script type="module">`

But only Vite 8 and the local static preview were rechecked here.

## Plan vs actual upstream API

Verified mismatches worth preserving:

1. `File`, `Blob`, and `Uint8Array` are not typed direct inputs in current declarations.
2. `workerUrl` is not a public upstream API.
3. Effective render mode is reliably observable on the engine, not on the viewer wrapper.
4. `onError` is for later asynchronous viewer failures, not the primary `load()` failure channel.
5. XLSX has a much richer format-specific surface than a page-based abstraction would suggest.

## Spike implementation choice

This repository’s raw harness uses:

- engine-first loading
- `fromDocument` / `fromWorkbook` / `fromPresentation`
- worker mode by default
- explicit main-mode selection
- source normalization for URL/File/binary inputs
- no public `workerUrl`
- `wasmUrl` only as an optional passthrough because upstream explicitly supports it
