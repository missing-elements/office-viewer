# Office Viewer Implementation Plan

## Project

- Repository: `missing-elements/office-viewer`
- Package: `@missing-elements/office-viewer`
- Custom element: `<office-viewer>`

## Vision

Build a framework-neutral, read-only Web Component that renders Office Open XML documents directly in the browser without uploading files to a service and without embedding a full Office editor runtime.

## Scope

### Initial formats

- `.docx`
- `.xlsx`
- `.pptx`

### Explicitly out of scope

- PDF. PDF remains the responsibility of `@missing-elements/pdfjs-viewer-element`.
- Editing or mutation APIs.
- Saving or lossless round-tripping.
- Collaboration or server-side document sessions.
- Macro execution.
- Running embedded OLE applications.
- Legacy binary formats such as `.doc`, `.xls`, and `.ppt`.
- PowerPoint animations and transitions unless later supported by the upstream engine.
- Printing in the initial common API.

## Core dependency

Use `@silurus/ooxml` as the rendering engine. Do not use ONLYOFFICE, `x2t`, `pdfjs-dist`, or `pdfjs-viewer-element` in this package.

```text
DOCX / XLSX / PPTX
        |
        v
@silurus/ooxml Rust/WASM parser
        |
        v
validated format-specific model
        |
        v
Canvas renderer
```

`@missing-elements/office-viewer` is an orchestration and presentation layer around the upstream format-specific viewers. It must not duplicate their document models or parser logic.

## Development stack

Follow the conventions of `missing-elements/pdfjs-viewer`:

- TypeScript
- pnpm
- Vite
- Vitest
- Vitest browser tests
- WebdriverIO browser provider
- generated TypeScript declarations
- ESM package output
- `src/`, `types/`, `tests/`, `demo/`, and `scripts/`
- strict lifecycle and browser compatibility tests

## Architecture

```text
src/
  index.ts
  office-viewer-element.ts
  types.ts
  source-resolver.ts
  format-detector.ts
  load-controller.ts
  viewer-adapter.ts
  viewers/
    docx-adapter.ts
    xlsx-adapter.ts
    pptx-adapter.ts
  ui/
    loading-view.ts
    error-view.ts
    toolbar.ts
  themes/
    shell.css
```

The custom element owns Shadow DOM, status UI, attributes, source resolution, adapter selection, common events, resize handling, and lifecycle. Adapters own format-specific upstream viewer instances.

## Upstream API study — mandatory first step

Before finalizing the public API, carefully read the current `@silurus/ooxml` README, generated declarations, format documentation, examples, and bundle-size documentation.

Verify:

- constructors and `load()` signatures for DOCX, XLSX, and PPTX;
- accepted source types;
- `DocxScrollViewer`, `XlsxViewer`, and `PptxScrollViewer` behavior;
- worker-mode options and fallback behavior;
- destroy and ownership semantics;
- zoom, fit, navigation, find, and hyperlink APIs;
- progressive-layout callbacks and completion semantics;
- WASM URL configuration;
- worker asset handling;
- stable error types and codes;
- optional renderer injection;
- bundler and static-hosting behavior.

Record verified facts and deviations in `docs/ooxml-integration-notes.md`. Do not treat the provisional interfaces in this plan as verified until this step is complete.

## Format adapters

### DOCX

Use `DocxScrollViewer` where appropriate. DOCX is a paginated document and benefits from continuous scrolling, virtualization, progressive layout, text selection, find, hyperlinks, zoom, and page navigation.

### XLSX

Use `XlsxViewer`. XLSX is not paginated. Preserve worksheet tabs, grid scrolling, cell/range selection, zoom, find, hyperlinks, frozen panes, and workbook-specific behavior.

Do not expose page-oriented methods as if they applied to XLSX.

### PPTX

Use `PptxScrollViewer` where appropriate. Preserve continuous slide scrolling, virtualization, slide navigation, zoom, find, text selection, hyperlinks, and any verified speaker-note capability.

## Rendering mode

The public mode is:

```text
mode="worker|main"
```

- Worker mode is the default.
- `mode="worker"` means prefer worker rendering.
- `mode="main"` explicitly requests main-thread rendering.
- If the upstream engine must fall back to main mode for compatibility, the component may do so and should make the effective mode observable internally.
- There is no `auto` mode.
- There is no `worker-required` option in the initial release.

Worker mode requires verification of `Worker`, `OffscreenCanvas`, `ImageBitmap`, worker asset paths, and DOCX shaping fallback behavior. `office-viewer` must not create a second rendering worker; it should use the worker lifecycle managed by `@silurus/ooxml`.

## Source and loading rules

```ts
type OfficeSource =
  | string
  | URL
  | File
  | Blob
  | ArrayBuffer
  | Uint8Array
```

`load(source)` is canonical. The `src` attribute is a convenience URL API:

- changing `src` loads the new URL;
- `load(source)` does not need to reflect the source into `src`;
- `reload()` reloads the last retained source and options;
- `file-name` supplies metadata for sources without a filename;
- `file-type` supplies an explicit format override.

Caller-owned buffers must not be detached or mutated unexpectedly. If transferables are used internally, copy the buffer before transferring it to a worker. Do not expose a transfer-ownership option initially.

The component owns and revokes only object URLs that it creates. It must retain enough source information to implement `downloadOriginal()` and reload direct binary sources.

## Format detection

Use conservative detection for the MVP:

1. `file-type` attribute.
2. `format` option passed to `load()`.
3. `File.name`, `file-name`, or URL pathname extension.
4. Otherwise fail clearly for extensionless binary input.

Do not add ZIP package-part inspection to the MVP solely for ambiguous extensionless files. Consider it later if real use cases justify the dependency and complexity.

## Public API direction

Initial attributes:

```text
src
file-name
file-type
mode="worker|main"
theme
zoom
page
slide
sheet
enable-text-selection
enable-hyperlinks
enable-download
show-toolbar
```

`locale` is intentionally excluded. Printing is intentionally excluded from the initial common API.

Initial methods:

```ts
load(source: OfficeSource, options?: OfficeViewerLoadOptions): Promise<void>
reload(): Promise<void>
setScale(scale: number): void
fitWidth(): void
fitPage(): void
findText(query: string): void
clearFind(): void
downloadOriginal(): Promise<void>
destroy(): void
```

`downloadOriginal()` downloads the original source bytes only. It does not export PDF, images, or modified Office documents.

The exact `wasmUrl` type must match the current upstream declarations. Expose `wasmUrl` only if verified. Do not expose `workerUrl` in the initial release.

Initial events:

```text
loadstart
progress
ready
loaderror
pagechange
slidechange
sheetchange
zoomchange
destroy
```

Events must be typed, small, and serializable. Format-specific operations remain adapter-specific unless their semantics are truly common.

## Lifecycle

Use an explicit state machine:

```text
idle → loading → ready
                 ↘ error

ready/error → loading on new load
any temporary state → detached on disconnect
any state → destroyed on permanent destroy()
```

Each load gets a generation token and an `AbortController`. A stale asynchronous result must never replace a newer load.

`disconnectedCallback()` temporarily releases the active upstream viewer and retains the source/configuration needed for reconnection. Reconnection recreates and reloads when appropriate. `destroy()` permanently aborts work, releases the upstream viewer and render resources, revokes component-owned URLs, clears active resources, and requires a new explicit `load()` call.

## Optional modules

Do not enable optional renderers by default:

- MathJax
- ChartEx
- 3-D charts
- Region Maps
- TIFF decoding

Add public configuration only after verifying upstream types and worker-mode behavior. Optional assets must not be fetched for documents that do not need them.

## Styling and accessibility

Use an open Shadow DOM, but provide only shell styling where it has clear value:

- viewer background/desk;
- page or slide gaps;
- shadows;
- status UI;
- optional toolbar;
- XLSX shell elements where supported.

Do not imply that authored document content can be themed. `locale` is not part of the MVP.

The host must have a bounded height:

```css
office-viewer {
  display: block;
  height: 100dvh;
}
```

Required accessibility work:

- `aria-busy` while loading;
- accessible loading and error status;
- labelled viewer region;
- keyboard navigation where supported;
- focus-visible styles;
- clear fallback when Canvas or worker rendering is unavailable.

## Bundle and asset strategy

Measure, do not assume. Compare:

- combined all-format import;
- DOCX-only, XLSX-only, and PPTX-only imports;
- dynamic adapter imports;
- JavaScript and WASM transfer sizes;
- worker assets;
- first visible render time;
- memory usage;
- warm-cache behavior.

Potential subpath exports remain deferred:

```text
@missing-elements/office-viewer
@missing-elements/office-viewer/docx
@missing-elements/office-viewer/xlsx
@missing-elements/office-viewer/pptx
```

Use `@silurus/ooxml` as a normal dependency initially. Decide on dynamic imports and subpath exports only after verifying real bundle graphs and asset reliability.

## Testing

Use Vitest and WebdriverIO as the primary test stack.

### Unit tests

- source normalization;
- filename and extension inference;
- format detection;
- attribute parsing;
- mode selection;
- abort and stale-load handling;
- error normalization;
- source ownership;
- event payloads;
- cleanup.

### Browser tests

- load each format from URL, File, Blob, ArrayBuffer, and Uint8Array;
- worker mode and main mode;
- verified worker fallback;
- reload and replacement;
- disconnect/reconnect;
- permanent destroy;
- resize and refit;
- zoom and navigation;
- find and hyperlinks;
- theme/shell behavior;
- original download;
- malformed, encrypted, legacy, and unsupported input.

### Asset tests

- production WASM and worker paths;
- non-root static deployment;
- explicit `wasmUrl` when supported;
- no accidental inline WASM;
- no development-only worker path;
- optional modules absent by default.

### Visual fixtures

```text
simple.docx
tables.docx
images.docx
equations.docx
simple.xlsx
formatted.xlsx
charts.xlsx
simple.pptx
shapes.pptx
images.pptx
```

## Milestones

### Milestone 0 — Architecture spike

- Complete the upstream API study.
- Create `docs/ooxml-integration-notes.md`.
- Build a raw DOCX/XLSX/PPTX harness.
- Verify worker mode and main mode in production builds.
- Verify WASM and worker asset paths.
- Test `File` and URL sources.
- Measure per-format bundle/runtime costs.
- Record all upstream pitfalls.

Do not finalize the custom-element API before this milestone is complete.

### Milestone 1 — Minimal DOCX component

- Create package skeleton.
- Implement Shadow DOM and bounded container.
- Implement source loading and lifecycle state.
- Add DOCX adapter using worker mode by default.
- Add loading/error states.
- Add disconnect/reconnect and destroy behavior.
- Add initial browser tests.

### Milestone 2 — XLSX and PPTX

- Add XLSX and PPTX adapters.
- Add conservative format detection.
- Add binary source types.
- Add common events and format-specific navigation.

### Milestone 3 — Lifecycle and asset hardening

- Add abortable loads and stale-load protection.
- Verify source ownership and buffer-copy behavior.
- Verify worker fallback.
- Verify WASM override if supported.
- Test static hosting, Vite, plain modules, and supported framework bundlers.

### Milestone 4 — Verified interactions

- Add zoom and fit.
- Add find and hyperlinks.
- Add page/slide/sheet events.
- Add `downloadOriginal()`.
- Do not add common printing until separately verified.

### Milestone 5 — UX and accessibility

- Add minimal shell styling.
- Add optional toolbar only if it does not distort format-specific behavior.
- Add keyboard and status accessibility.
- Add drag-and-drop and mobile demo behavior.

### Milestone 6 — Production release

- Finalize exports and declarations.
- Complete browser and visual tests.
- Publish bundle-size measurements.
- Complete documentation and troubleshooting guidance.
- Add changelog and release workflow.
- Verify CDN/plain-module usage.

## Definition of success

A user can write:

```html
<script type="module" src="office-viewer-element.js"></script>

<office-viewer
  src="./report.docx"
  mode="worker"
  style="height: 100dvh">
</office-viewer>
```

and receive a private, browser-rendered, read-only Office viewer with no server, upload service, ONLYOFFICE runtime, editor UI, PDF.js dependency, or manual extraction.

For local `File`, `Blob`, and binary sources, the component processes bytes locally and does not upload them. URL sources are fetched from the URL supplied by the consumer.

## First implementation rule

Carefully verify the current `@silurus/ooxml` documentation, declarations, examples, worker behavior, asset handling, and bundle measurements before locking the stable public API.
