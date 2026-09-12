# office-viewer

[![npm version](https://img.shields.io/npm/v/@missing-elements/office-viewer?logo=npm&logoColor=fff)](https://www.npmjs.com/package/@missing-elements/office-viewer)

Headless Web Component orchestration layer for browser-based Office Open XML viewing built on [@silurus/ooxml](https://github.com/yukiyokotani/office-open-xml-viewer#third-party-notices).

`<office-viewer>` is an **orchestration shell**, not a viewer widget. It loads DOCX, XLSX, and PPTX documents, manages the upstream viewer lifecycle, and exposes the upstream viewer instance directly so consumers can build their own UI.

- Open Shadow DOM protects the viewer surface from global styles.
- No UI chrome, toolbars, or status widgets.
- Direct access to the upstream `getViewer()`.

## Installation

```bash
npm install @missing-elements/office-viewer
```

## Usage

```html
<script type="module">
  import { defineOfficeViewerElement } from '@missing-elements/office-viewer';
  defineOfficeViewerElement();

  const element = document.querySelector('office-viewer');
  await element.load('/report.docx', { format: 'docx' });

  const viewer = element.getViewer();
  viewer.setScale(1.5);
</script>

<office-viewer style="height: 100svh"></office-viewer>
```

### CodePen and other hosted editors

Worker mode loads the format parser as a WebAssembly asset. Hosted editors that
bundle JavaScript but do not serve package assets need an absolute `wasmUrl`.
Use the asset that matches the document format and pin it to the installed
`@silurus/ooxml` version:

```js
await element.load(file, {
  format: 'docx',
  mode: 'worker',
  wasmUrl: 'https://cdn.jsdelivr.net/npm/@silurus/ooxml@0.86.1/dist/docx_parser_bg.wasm'
});
```

Replace `docx` in both places with `xlsx` or `pptx` for those formats. The
`main` mode does not require this workaround.

## Public API

### Attributes

| Attribute | Description |
| --- | --- |
| `src` | URL to load. |
| `file-type` | Explicit format: `docx`, `xlsx`, `pptx`. |
| `mode` | Rendering mode: `worker` (default) or `main`. |

### Methods

| Method | Description |
| --- | --- |
| `load(source, options)` | Load a document. `source` is a URL string, `ArrayBuffer`, `Blob`/`File`, or `ReadableStream<Uint8Array>`. `options.format` is required; `options.wasmUrl` may point to a served parser WASM asset. |
| `reload()` | Reload the last source and options. |
| `destroy()` | Tear down the upstream viewer. |
| `getViewer()` | Returns the upstream viewer instance (`DocxScrollViewer`, `XlsxViewer`, or `PptxScrollViewer`). |

### Properties

| Property | Description |
| --- | --- |
| `ready` | `true` when a document is loaded and ready. |
| `error` | Last error, if any. |
| `format` | Loaded format. |
| `mode` | Effective rendering mode. |

### Events

Events are dispatched by the element but not required for SSR-compatible usage.

| Event | Detail |
| --- | --- |
| `loadstart` | — |
| `ready` | — |
| `loaderror` | `{ error }` |
| `destroy` | — |

## Reacting to loads

Use the element's events for client-side load state. The `ready` property is
`true` only after a successful load, and `error` contains the latest failure.

```html
<button id="open" type="button">Open report</button>
<span id="state" role="status"></span>

<office-viewer style="height: 80svh;"></office-viewer>

<script type="module">
  import { defineOfficeViewerElement } from '@missing-elements/office-viewer';

  defineOfficeViewerElement();

  const viewer = document.querySelector('office-viewer');
  const openButton = document.querySelector('#open');
  const state = document.querySelector('#state');

  viewer.addEventListener('loadstart', () => {
    openButton.disabled = true;
    state.textContent = 'Loading report...';
  });

  viewer.addEventListener('ready', () => {
    openButton.disabled = false;
    state.textContent = 'Report ready.';
  });

  viewer.addEventListener('loaderror', (event) => {
    openButton.disabled = false;
    state.textContent = `Could not open report: ${event.detail.error.message}`;
  });

  openButton.addEventListener('click', async () => {
    try {
      await viewer.load('/report.docx', { format: 'docx' });
    } catch {
      // loaderror has already updated the UI.
    }
  });
</script>
```

## Source types

`load()` accepts a URL `string`, `ArrayBuffer`, `Blob` (including `File`), or
`ReadableStream<Uint8Array>`. Convert `Uint8Array` to `ArrayBuffer` before
calling `load()`.

```js
// Load from a File object (e.g., from an <input type="file">)
const file = document.querySelector('input[type="file"]').files[0];
const arrayBuffer = await file.arrayBuffer();
await element.load(arrayBuffer, { format: 'docx' });

// Or load directly from a URL string
await element.load('/report.docx', { format: 'docx' });
```

## Detecting format

Use [`file-type`](https://www.npmjs.com/package/file-type) to detect a local
file before loading it:

```js
import { fileTypeFromBuffer } from 'file-type';

const arrayBuffer = await file.arrayBuffer();
const detected = await fileTypeFromBuffer(arrayBuffer);

if (!detected || !['docx', 'xlsx', 'pptx'].includes(detected.ext)) {
  throw new Error('Select a DOCX, XLSX, or PPTX file.');
}

await element.load(arrayBuffer, { format: detected.ext });
```

## Browser support

- Modern evergreen browsers with custom element support.
- Web Workers are used by default; fall back to `mode="main"` for environments that block workers.

## License

MIT

## Third-Party Notices

https://github.com/yukiyokotani/office-open-xml-viewer#third-party-notices
