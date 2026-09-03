# Bundle size and runtime spike

Measured from `pnpm build` and a Firefox headless production-preview run on September 3, 2026.

## Build environment

- `@silurus/ooxml@0.85.3`
- `vite@8.2.2`
- Firefox `154.0`
- production preview served from `pnpm preview`

## JavaScript graph size

Transitive built JavaScript requested by the measurement entry pages:

| Entry | Total JS bytes | Notes |
| --- | ---: | --- |
| demo shell (`demo/index.html`) | 10,684 | initial shell only; formats are lazy-loaded |
| DOCX-only (`demo/measure-docx.html`) | 1,471,340 | includes shared renderer/runtime chunks |
| XLSX-only (`demo/measure-xlsx.html`) | 929,085 | includes shared renderer/runtime chunks |
| PPTX-only (`demo/measure-pptx.html`) | 956,660 | includes shared renderer/runtime chunks |
| Combined static import (`demo/measure-combined.html`) | 2,014,425 | all three formats loaded together |

Largest built JS assets:

| Asset | Size |
| --- | ---: |
| `renderer-module-contract-*.js` | 649,410 |
| `docx-*.js` | 783,421 |
| `xlsx-*.js` | 274,259 |
| `pptx-*.js` | 265,468 |

## WASM sizes

| Asset | Size |
| --- | ---: |
| `docx_parser_bg-*.wasm` | 1,830,632 |
| `xlsx_parser_bg-*.wasm` | 1,617,054 |
| `pptx_parser_bg-*.wasm` | 1,691,099 |

## Worker asset sizes

| Asset | Size |
| --- | ---: |
| DOCX worker `render-worker-D1O73gYW-*.js` | 1,655,727 |
| XLSX worker `render-worker-DHP0rgcW-*.js` | 1,054,740 |
| PPTX worker `render-worker-V2wWYA4O-*.js` | 1,111,481 |

## First-load requests observed in production preview

Main-thread resource timing from the built demo:

### DOCX worker load

1. `rolldown-runtime-*.js`
2. `highlight-rect-*.js`
3. `docx-*.js`
4. `renderer-module-contract-*.js`
5. `fixtures/sample.docx`
6. `render-worker-host-*.js`
7. `render-worker-D1O73gYW-*.js`

### XLSX worker load

1. `rolldown-runtime-*.js`
2. `renderer-module-contract-*.js`
3. `xlsx-*.js`
4. `resource-measurement-*.js`
5. `fixtures/sample.xlsx`
6. `render-worker-host-*.js`
7. `render-worker-DHP0rgcW-*.js`

### PPTX worker load

1. `rolldown-runtime-*.js`
2. `pptx-*.js`
3. `renderer-module-contract-*.js`
4. `highlight-rect-*.js`
5. `resource-measurement-*.js`
6. `fixtures/sample.pptx`
7. `render-worker-host-*.js`
8. `render-worker-V2wWYA4O-*.js`

### DOCX main-mode comparison

1. `rolldown-runtime-*.js`
2. `docx-*.js`
3. `renderer-module-contract-*.js`
4. `highlight-rect-*.js`
5. `fixtures/sample.docx`

Notes:

- page `performance.getEntriesByType('resource')` did not expose worker-internal WASM fetches in Firefox headless, even though the worker-mode demo rendered successfully
- the build output and artifact tests independently confirmed emitted `.wasm` assets

## First visible timings

Representative preview timings from the raw harness summary:

| Scenario | First visible ms | Load completed ms |
| --- | ---: | ---: |
| DOCX worker | 348 | 348 |
| XLSX worker | 200 | 200 |
| PPTX worker | 213 | 214 |
| DOCX main | 189 | 190 |

These numbers are only spike evidence from a single CI runner + headless Firefox session, not release benchmarks.

## Dynamic import conclusion

Dynamic imports materially improve the initial graph for the harness:

- initial demo shell JS stays ~10.6 KB
- format payloads are deferred until the selected format is loaded
- a static combined import would ship ~2.0 MB of JS before any WASM or worker assets

## Are format-specific exports justified?

Yes.

Reasons:

- DOCX is substantially larger than XLSX or PPTX
- worker and WASM assets are also format-specific
- format-specific loading keeps the initial shell very small
- a combined static entry is acceptable for dedicated apps, but not for a general `office-viewer` package default
