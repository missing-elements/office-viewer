# office-viewer

Headless Web Component orchestration layer for browser-based Office Open XML viewing built on `@silurus/ooxml`.

`<office-viewer>` is an **orchestration shell**, not a viewer widget. It loads DOCX, XLSX, and PPTX documents, manages the upstream viewer lifecycle, and exposes the upstream engine instances directly so consumers can build their own UI.

- No Shadow DOM.
- No rendered children or built-in chrome.
- No wrapper methods like `setScale()` or `findText()`.
- Direct access to upstream `getViewer()`, `getDocument()`, and `getEngine()`.

## Status

Milestone 4 (API Stabilization) is complete. The package emits a production ESM bundle, generated TypeScript declarations, and has verified CDN/plain-module consumption.

## Installation

```bash
npm install @missing-elements/office-viewer
```

```html
<script type="module">
  import { defineOfficeViewerElement } from '@missing-elements/office-viewer';
  defineOfficeViewerElement();
</script>

<office-viewer src="/report.docx" file-type="docx" mode="worker"></office-viewer>
```

## Usage

### Declarative

```html
<office-viewer
  src="/report.docx"
  file-name="report.docx"
  file-type="docx"
  mode="worker"
  wasm-url="/assets/docx_parser_bg.wasm"
></office-viewer>
```

### Imperative

```ts
import { defineOfficeViewerElement, type OfficeViewerElement } from '@missing-elements/office-viewer';

defineOfficeViewerElement();

const element = document.createElement('office-viewer') as OfficeViewerElement;
document.body.append(element);

element.addEventListener('ready', () => {
  const viewer = element.getViewer();
  const document = element.getDocument();
  // Use upstream APIs directly.
});

element.addEventListener('loaderror', (event) => {
  console.error(event.detail.error);
});

await element.load('/report.docx', { format: 'docx', mode: 'worker' });
```

### Programmatic sources

The `load()` method accepts URLs, `File`, `Blob`, `ArrayBuffer`, and `Uint8Array`.

```ts
const file = document.querySelector('input[type="file"]').files[0];
await element.load(file);
```

Binary sources require an explicit `format` option or a `file-name` attribute unless the source is a `File` with a recognized extension.

## Public API

### Attributes

| Attribute | Description |
| --- | --- |
| `src` | URL to load. |
| `file-name` | Metadata filename for format detection. |
| `file-type` | Explicit format override: `docx`, `xlsx`, `pptx`. |
| `mode` | Rendering mode: `worker` (default) or `main`. |
| `wasm-url` | Custom WASM asset URL passed to the upstream engine. |

### Methods

| Method | Description |
| --- | --- |
| `load(source, options?)` | Load a document from any supported source. |
| `reload()` | Reload the last source and options. |
| `destroy()` | Cancel pending loads and tear down the upstream viewer. |
| `getViewer()` | Returns the upstream viewer instance. |
| `getDocument()` | Returns the upstream document model. |
| `getEngine()` | Returns the upstream engine instance. |

### Properties

| Property | Description |
| --- | --- |
| `ready` | `true` when a document is loaded and ready. |
| `error` | Last error, if any. |
| `format` | Detected or explicit format. |
| `mode` | Effective rendering mode. |

### Events

| Event | Detail |
| --- | --- |
| `loadstart` | — |
| `ready` | `{ format, requestedMode, effectiveMode }` |
| `loaderror` | `{ error }` |
| `destroy` | — |

## Bundle size

Measured from `pnpm build`. The library entry itself is a small orchestration layer (< 100 KB uncompressed). Upstream parser WASM and worker assets are loaded on demand per format.

| Entry | Approximate JS |
| --- | ---: |
| Library entry (`office-viewer.es.js`) | < 100 KB |
| DOCX viewer + runtime | ~1.4 MB |
| XLSX viewer + runtime | ~0.9 MB |
| PPTX viewer + runtime | ~1.0 MB |

See [`docs/bundle-size-spike.md`](docs/bundle-size-spike.md) for full measurements.

## Browser support

- Modern evergreen browsers with custom element support.
- Web Workers are used by default; fall back to `mode="main"` for environments that block workers or blob-worker URLs.

## License

MIT
