# office-viewer

[![npm version](https://img.shields.io/npm/v/@missing-elements/office-viewer?logo=npm&logoColor=fff)](https://www.npmjs.com/package/@missing-elements/office-viewer)

> **Status: not production ready yet.** This is an early 0.x release. The API,
> lifecycle behavior, and upstream dependency can still change between minor
> versions. Try it and report issues, but do not rely on it in production.

**Live demo:** [office-viewer-pi.vercel.app](https://office-viewer-pi.vercel.app/)

Headless Web Component orchestration layer for browser-based Office Open XML viewing built on [@silurus/ooxml](https://github.com/yukiyokotani/office-open-xml-viewer#third-party-notices).

`<office-viewer>` is an **orchestration shell**, not a viewer widget. It loads DOCX, XLSX, and PPTX documents, manages the upstream viewer lifecycle, and exposes the upstream viewer instance directly so consumers can build their own UI.

- Open Shadow DOM protects the viewer surface from global styles.
- No UI chrome, toolbars, or status widgets.
- Direct access to the upstream `getViewer()`.

## Installation

```bash
npm install @missing-elements/office-viewer
```

## Demo

The demo at [office-viewer-pi.vercel.app](https://office-viewer-pi.vercel.app/)
is the `demo/` folder built with Vite. Locally, `pnpm dev` serves it from
source and `pnpm build:demo` writes the same static site to `demo/dist`.

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

The same load can be declared with attributes:

```html
<office-viewer src="/report.docx" file-type="docx" style="height: 100svh"></office-viewer>
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
`wasm-url` attribute takes the same value for declarative loads. The `main`
mode does not require this workaround.

## Public API

### Attributes

| Attribute | Description |
| --- | --- |
| `src` | URL to load. Requires `file-type`. Clearing it unloads a document that was loaded from attributes. |
| `file-type` | Format of `src`: `docx`, `xlsx`, or `pptx`. Required for attribute-driven loading; a missing or unknown value emits `loaderror`. |
| `mode` | Rendering mode: `worker` (default) or `main`. Unknown values fall back to the default. |
| `wasm-url` | Optional absolute URL of the parser WASM asset. See the hosted-editor note above. |

Attributes set together in one synchronous run start a single load. A `load()`
call made after attribute changes in the same run outranks them. Moving the
element in the DOM without changing attributes does not reload it.

### Methods

| Method | Description |
| --- | --- |
| `load(source, options)` | Load a document. `source` is a URL string, `ArrayBuffer`, `Blob`/`File`, or `ReadableStream<Uint8Array>`. `options.format` is required, `options.mode` defaults to `worker`, and `options.wasmUrl` may point to a served parser WASM asset. Resolves after `ready` and rejects after `loaderror`. A newer `load()`, `destroy()`, or removal from the document cancels an in-flight load at once: its promise rejects with a `DOMException` named `AbortError` and no `loaderror` is emitted. |
| `reload()` | Repeat the most recent `load()` request, including one that failed. URL strings are fetched again; `Blob`/`File` sources are read again; `ArrayBuffer` and stream sources are replayed from a copy kept by the element. Rejects if nothing has been requested since the last `destroy()`. |
| `destroy()` | Cancel an in-flight load, tear down the upstream viewer, forget the retained request, reset all state, and emit `destroy`. Safe to call at any time. |
| `getViewer()` | Returns the upstream viewer instance (`DocxScrollViewer`, `XlsxViewer`, or `PptxScrollViewer`), or `null` while nothing is loaded. |

### Properties

| Property | Description |
| --- | --- |
| `ready` | `true` while a document is loaded. Stays `true` for the current document while a newer load is in flight. |
| `error` | Error from the most recent load, or `null`. Cleared when a load starts and on `destroy()`. |
| `format` | Format of the loaded document, or `null`. |
| `mode` | Rendering mode the loaded document's viewer was created with (`worker` when none was given), or `null`. |

`format` and `mode` describe the document returned by `getViewer()`. They
change when `ready` fires, not when a load starts. Upstream may fall back to
main-thread rendering internally for some documents; `mode` reports what was
requested, not that fallback.

### Events

| Event | Detail |
| --- | --- |
| `loadstart` | — |
| `ready` | — |
| `loaderror` | `{ error }` (typed as `OfficeViewerLoadErrorDetail`) |
| `destroy` | — |

Every `load()` emits `loadstart` and then exactly one of `ready` or
`loaderror`, unless it is cancelled, in which case neither follows. `destroy`
is emitted by `destroy()` and when a removed element releases its viewer.

## Lifecycle

- Loads are last-writer-wins. Calling `load()` while another load is running
  cancels the older one immediately: its stream read is cancelled, its upstream
  viewer is destroyed, and its promise rejects with `AbortError` without waiting
  for the parser. Element state and events only ever reflect the newest request.
- Removing an element that holds a viewer or an in-flight load releases them
  and emits `destroy`. The request stays retained, so adding the element back
  loads the same document again, whether it came from `load()` or from
  attributes. Attributes changed while detached take precedence. Call
  `destroy()` to discard the request instead. The release is deferred by one
  microtask, so moving the element synchronously (remove, then append in the
  same task) keeps the viewer alive. An idle element is not affected by
  removal.

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
    } catch (error) {
      // loaderror has already updated the UI; AbortError means a newer load took over.
      if (error.name !== 'AbortError') console.error(error);
    }
  });
</script>
```

## Source types

`load()` accepts a URL `string`, `ArrayBuffer`, `Blob` (including `File`), or
`ReadableStream<Uint8Array>`. Convert `Uint8Array` to `ArrayBuffer` before
calling `load()`. The parser may take ownership of an `ArrayBuffer` it is
given, leaving the buffer detached (empty) after the load, so the element keeps
a copy for `reload()`. Streams are read once into a `Blob` that `reload()`
replays.

```js
// Load from a File object (e.g., from an <input type="file">)
const file = document.querySelector('input[type="file"]').files[0];
await element.load(file, { format: 'docx' });

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
