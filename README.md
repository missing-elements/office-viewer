# office-viewer

Headless Web Component orchestration layer for browser-based Office Open XML viewing built on `@silurus/ooxml`.

`<office-viewer>` is an **orchestration shell**, not a viewer widget. It loads DOCX, XLSX, and PPTX documents, manages the upstream viewer lifecycle, and exposes the upstream viewer instance directly so consumers can build their own UI.

- Open Shadow DOM protects the viewer surface from global styles.
- No UI chrome, toolbars, or status widgets.
- No wrapper methods like `setScale()` or `findText()`.
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

<office-viewer style="display: block; height: 100dvh;"></office-viewer>
```

## Public API

### Attributes

| Attribute | Description |
| --- | --- |
| `src` | URL to load. |
| `file-type` | Explicit format: `docx`, `xlsx`, `pptx`. |
| `mode` | Rendering mode: `worker` (default) or `main`. |
| `wasm-url` | Custom WASM asset URL. |

### Methods

| Method | Description |
| --- | --- |
| `load(source, options)` | Load a document. `source` is `string \| ArrayBuffer`. `options.format` is required. |
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

| Event | Detail |
| --- | --- |
| `loadstart` | — |
| `ready` | — |
| `loaderror` | `{ error }` |
| `destroy` | — |

## Source types

`load()` accepts `string` (URL) or `ArrayBuffer`. Convert `File`, `Blob`, or `Uint8Array` to `ArrayBuffer` before calling `load()`.

```ts
const file = document.querySelector('input[type="file"]').files[0];
const arrayBuffer = await file.arrayBuffer();
await element.load(arrayBuffer, { format: 'docx' });
```

## Browser support

- Modern evergreen browsers with custom element support.
- Web Workers are used by default; fall back to `mode="main"` for environments that block workers.

## License

MIT
