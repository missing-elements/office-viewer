# Office Viewer Implementation Plan

## Project

Repository: `missing-elements/office-viewer`

Package: `@missing-elements/office-viewer`

Custom element: `<office-viewer>`

## Vision

Build a framework-neutral, read-only Web Component that renders Office Open XML documents directly in the browser without uploading files to a server and without embedding a full Office editor runtime.

## Scope

### Initial supported formats

- `.docx`
- `.xlsx`
- `.pptx`

### Explicitly out of scope

- PDF support; PDF remains the responsibility of `@missing-elements/pdfjs-viewer-element`.
- Editing or mutation APIs.
- Saving or lossless round-tripping of Office documents.
- Collaboration.
- Macro execution.
- Running embedded OLE applications.
- Legacy binary formats such as `.doc`, `.xls`, and `.ppt`.
- PowerPoint animations and transitions unless supported later by the underlying library.

## Core dependency

Use `@silurus/ooxml` as the rendering engine. Do not use ONLYOFFICE, `x2t`, `pdfjs-dist`, or `pdfjs-viewer-element` in this package.

The intended runtime pipeline is:

```text
DOCX / XLSX / PPTX
        |
        v
@silurus/ooxml Rust/WASM parser
        |
        v
validated document model
        |
        v
Canvas renderer
```

## Related project conventions

Follow the development stack and conventions used by `missing-elements/pdfjs-viewer`:

- TypeScript
- pnpm
- Vite
- Vitest
- Vitest browser tests
- WebdriverIO browser provider
- generated TypeScript declarations
- `src/`, `types/`, `tests/`, `demo/`, and `scripts/` directories
- ESM package output
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
    paper-and-ink.css
```

### Custom element

`office-viewer-element.ts` owns:

- Shadow DOM creation.
- Observed attributes and property reflection.
- Viewer container creation.
- Loading, error, and empty states.
- Adapter selection.
- Common event dispatch.
- Resize handling.
- Destruction and reconnection behavior.

### Source resolver

Support these sources:

```ts
type OfficeSource =
  | string
  | URL
  | File
  | Blob
  | ArrayBuffer
  | Uint8Array
```

The resolver must:

- Fetch URL sources.
- Read `File` and `Blob` sources.
- Preserve the filename where available.
- Infer a format from explicit options or the filename.
- Support `AbortSignal`.
- Avoid unexpectedly mutating caller-owned buffers.
- Preserve the original source for download where possible.

### Format detection

Determine the format in this order:

1. Explicit `file-type` attribute.
2. Explicit `load()` option.
3. Filename extension.
4. Container signature validation where practical.

Unknown or unsupported formats must produce a clear error. Do not silently assume DOCX.

### Load controller

Loading must be race-safe:

```text
load(A) starts
load(B) starts
A finishes after B
A must not replace B
```

Use an `AbortController` and a load generation/token. A new load must cancel or invalidate the previous load and destroy the old viewer only when it is safe to do so.

## Format adapters

Keep format-specific behavior behind adapters. Do not force DOCX, XLSX, and PPTX into an identical internal model.

### DOCX

Use the `@silurus/ooxml` continuous scroll viewer where appropriate.

Expected capabilities:

- Virtualized page scrolling.
- Progressive layout where useful.
- Worker rendering.
- Zoom and fit controls.
- Text selection.
- Find.
- Hyperlinks.
- Page navigation.

### XLSX

Use the native workbook/sheet viewer.

Expected capabilities:

- Worksheet tabs.
- Grid scrolling.
- Cell and range selection.
- Zoom.
- Find.
- Hyperlinks.
- Frozen panes and workbook-specific behavior.

Do not model XLSX as a paginated document.

### PPTX

Use the continuous slide viewer where appropriate.

Expected capabilities:

- Virtualized slide scrolling.
- Slide navigation.
- Zoom and fit controls.
- Text selection.
- Find.
- Hyperlinks.
- Optional speaker-note access if it fits the common API.

## Adapter interface

The exact interface must be finalized after reading the current `@silurus/ooxml` declarations and documentation.

Initial direction:

```ts
interface ViewerAdapter {
  readonly format: OfficeFormat
  load(source: ResolvedSource, options?: ViewerLoadOptions): Promise<void>
  destroy(): void
  setScale?(scale: number): void
  fitWidth?(): void
  fitPage?(): void
  findText?(query: string): void
  clearFind?(): void
  print?(): Promise<void>
}
```

The adapter must not expose internal parser models as stable public API.

## Public element API

### Example

```html
<script type="module" src="./office-viewer-element.js"></script>

<office-viewer
  src="/documents/report.docx"
  mode="worker"
  style="height: 800px">
</office-viewer>
```

### Initial attributes

Only add attributes that can be implemented consistently:

- `src`
- `file-name`
- `file-type`
- `mode="main|worker"`
- `theme="automatic|light|dark"`
- `locale`
- `zoom`
- `page`
- `slide`
- `sheet`
- `enable-text-selection`
- `enable-hyperlinks`
- `enable-download`
- `enable-print`
- `show-toolbar`

### Initial methods

```ts
load(source: OfficeSource, options?: OfficeViewerLoadOptions): Promise<void>
reload(): Promise<void>
setScale(scale: number): void
fitWidth(): void
fitPage(): void
findText(query: string): void
clearFind(): void
print(): Promise<void>
download(): Promise<void>
destroy(): void
```

Format-specific operations should be added only when their semantics are clear. For example, `goToSheet()` is meaningful for XLSX but not DOCX.

### Events

Initial events:

- `loadstart`
- `progress`
- `ready`
- `loaderror`
- `pagechange`
- `slidechange`
- `sheetchange`
- `zoomchange`
- `destroy`

Event details should be typed, small, and serializable.

## Worker mode

Prefer worker rendering by default when supported:

```text
mode="worker"
```

Verify the current `@silurus/ooxml` behavior before implementing defaults. Worker mode requires attention to:

- `Worker` support.
- `OffscreenCanvas` support.
- `ImageBitmap` ownership and cleanup.
- Worker and WASM asset URLs.
- CSP and static hosting.
- DOCX shaping fallback to main mode.
- Limitations on custom renderer objects.

Recommended behavior:

- If worker mode is requested and supported, use it.
- If worker mode is preferred but unavailable, fall back to main mode.
- If a future `worker-required` option is added, fail clearly when unsupported.
- Expose the effective mode so consumers can observe fallbacks.

## Optional modules

Keep optional renderers out of the default application graph:

- MathJax equations.
- ChartEx.
- 3-D charts.
- Region Maps.
- TIFF decoding.

Do not enable or bundle these automatically in the MVP. Add explicit configuration only after verifying the current `@silurus/ooxml` injection types and worker behavior.

## Styling

Use an open Shadow DOM, consistent with `pdfjs-viewer-element`.

The host must have a bounded height:

```css
office-viewer {
  display: block;
  height: 100dvh;
}
```

Expose CSS custom properties for integration:

```css
--office-viewer-background
--office-viewer-page-gap
--office-viewer-page-shadow
--office-viewer-accent-color
```

Provide a restrained Paper & Ink-inspired default theme without depending on PDF.js styles.

## Toolbar

The first technical spike should focus on rendering, not a complete toolbar.

Later, add an optional toolbar with common actions:

- Zoom out/in.
- Zoom level.
- Fit width/page.
- Previous/next page or slide.
- Find.
- Print.
- Download.

XLSX controls must remain sheet-aware and must not pretend that the workbook is a paginated document.

## Accessibility

Required for the MVP:

- `aria-busy` while loading.
- Accessible loading status.
- Accessible error status.
- Keyboard navigation where supported.
- Focus-visible styles.
- A labelled viewer region.
- A clear fallback if Canvas or worker rendering is unavailable.

Canvas output must not be described as fully accessible by itself. Integrate the underlying library's documented text-selection overlays and test them separately.

## Bundle and asset strategy

Measure rather than assume. Compare:

- Combined all-format import.
- DOCX-only import.
- XLSX-only import.
- PPTX-only import.
- Dynamic adapter imports.

Record:

- JavaScript size.
- WASM size per format.
- First-load requests.
- Time to first render.
- Memory usage.
- Repeat-load behavior.
- Optional module cost.

Potential future entry points:

```text
@missing-elements/office-viewer
@missing-elements/office-viewer/docx
@missing-elements/office-viewer/xlsx
@missing-elements/office-viewer/pptx
```

Do not finalize these exports until the technical spike confirms that they improve real application graphs and do not create unreliable WASM/worker asset handling.

## Testing plan

### Unit tests

- Source normalization.
- Filename inference.
- Format detection.
- Container validation.
- Attribute parsing.
- Abort behavior.
- Stale-load protection.
- Event payloads.
- Unsupported formats.
- Error mapping.
- Cleanup and destruction.

### Browser tests

- Load DOCX, XLSX, and PPTX from URL.
- Load `File`, `Blob`, `ArrayBuffer`, and `Uint8Array` sources.
- Worker mode.
- Main mode.
- Worker fallback.
- Destroy and reconnect.
- Reload a different document.
- Resize.
- Zoom.
- Find.
- Hyperlinks.
- Theme changes.
- Loading and error states.

### Visual fixtures

Use small representative fixtures:

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

Visual tests should verify the component shell and lifecycle as well as the underlying rendering output.

## Demo

Create a framework-neutral demo with:

- File picker.
- Drag-and-drop.
- URL loading.
- Format display.
- Loading progress.
- Error display.
- Zoom controls.
- Find.
- Theme switching.
- Original-file download.

Framework examples for React, Vue, and Svelte can be added after the core custom element is stable.

## Documentation

Create:

```text
README.md
docs/getting-started.md
docs/api.md
docs/loading-sources.md
docs/worker-mode.md
docs/bundle-size.md
docs/accessibility.md
docs/format-support.md
docs/troubleshooting.md
```

Documentation must clearly state:

- This is a read-only viewer.
- PDF is not supported by this package.
- Processing occurs locally in the browser.
- Supported formats are DOCX, XLSX, and PPTX.
- WASM assets are required.
- Worker mode has browser requirements and possible fallbacks.
- The host must provide a bounded height.
- Compatibility and limitations are inherited from `@silurus/ooxml`.

## Milestones

### Milestone 0 — Architecture spike

- Study the complete `@silurus/ooxml` documentation and public declarations.
- Inspect DOCX, XLSX, and PPTX examples.
- Confirm source and lifecycle semantics.
- Confirm worker mode and fallback behavior.
- Confirm WASM and worker asset handling.
- Measure format-specific bundle sizes.
- Load one real document of each format in a plain browser demo.
- Record pitfalls and decisions in an architecture note.

No polished custom element is required in this milestone.

### Milestone 1 — Minimal DOCX component

- Create the package skeleton.
- Implement Shadow DOM and bounded viewer container.
- Add source loading.
- Add DOCX adapter.
- Add loading/error states.
- Add worker mode.
- Add destruction and cleanup.
- Add the first browser tests.

### Milestone 2 — XLSX and PPTX

- Add XLSX adapter.
- Add PPTX adapter.
- Add format detection.
- Add common events.
- Add format-specific navigation where justified.
- Add `File`, `Blob`, and binary sources.

### Milestone 3 — Lifecycle hardening

- Add abortable loads.
- Add stale-load protection.
- Add reconnect handling.
- Add resize handling.
- Add effective-mode reporting.
- Add structured error mapping.
- Add asset-path and CSP tests.

### Milestone 4 — Common interaction API

- Add zoom.
- Add fit width/page.
- Add find.
- Add hyperlinks.
- Add page/slide/sheet events.
- Add original-file download.
- Add print where supported and verified.

### Milestone 5 — UX and accessibility

- Add optional toolbar.
- Add theme support.
- Add keyboard navigation.
- Improve loading and error announcements.
- Add drag-and-drop demo.
- Add mobile layout.

### Milestone 6 — Production release

- Finalize package exports.
- Generate and verify declarations.
- Complete browser matrix.
- Publish bundle-size measurements.
- Complete visual regression tests.
- Complete documentation.
- Add changelog and release workflow.
- Verify CDN and plain-module usage.

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

and receive a private, browser-only, read-only Office viewer with:

- no server;
- no upload;
- no ONLYOFFICE runtime;
- no editor UI;
- no PDF.js dependency;
- no manual extraction;
- clear loading and failure states;
- stable Web Component lifecycle;
- good DOCX, XLSX, and PPTX rendering.

## First implementation rule

Do not lock the public API before completing the architecture spike. Carefully verify the current `@silurus/ooxml` documentation, declarations, examples, worker behavior, asset handling, and bundle measurements first.
