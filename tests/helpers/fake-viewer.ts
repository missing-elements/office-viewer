// Stand-in for the @silurus/ooxml viewers so jsdom tests can drive the element's
// lifecycle without WASM, workers, or fixtures. It mirrors the parts of the upstream
// contract the element depends on:
// - the constructor mounts into the container synchronously;
// - load() takes ownership of an ArrayBuffer and detaches it;
// - destroy() is idempotent and does NOT settle a load() already in flight;
// - load() on a destroyed viewer rejects immediately.

export type FakeLoadBehavior = 'resolve' | 'hold'

interface Deferred {
  resolve(): void
  reject(reason: unknown): void
}

export class FakeViewer {
  static instances: FakeViewer[] = []
  static loadBehavior: FakeLoadBehavior = 'resolve'

  static reset(): void {
    FakeViewer.instances = []
    FakeViewer.loadBehavior = 'resolve'
  }

  static get last(): FakeViewer {
    const viewer = FakeViewer.instances.at(-1)
    if (!viewer) {
      throw new Error('No FakeViewer has been constructed.')
    }
    return viewer
  }

  readonly mount = document.createElement('div')
  /** What each load() received; ArrayBuffers are copied before the original is detached. */
  readonly sources: Array<string | ArrayBuffer> = []
  destroyed = false
  private deferred: Deferred | null = null

  constructor(
    readonly container: HTMLElement,
    readonly options: { mode?: string; wasmUrl?: string | URL } = {}
  ) {
    FakeViewer.instances.push(this)
    this.mount.className = 'fake-viewer'
    container.append(this.mount)
  }

  load(source: string | ArrayBuffer): Promise<void> {
    if (this.destroyed) {
      return Promise.reject(new Error('FakeViewer is destroyed'))
    }
    if (source instanceof ArrayBuffer) {
      this.sources.push(source.slice(0))
      structuredClone(source, { transfer: [source] })
    } else {
      this.sources.push(source)
    }
    if (FakeViewer.loadBehavior === 'resolve') {
      return Promise.resolve()
    }
    return new Promise<void>((resolve, reject) => {
      this.deferred = { resolve, reject }
    })
  }

  finish(): void {
    const deferred = this.deferred
    this.deferred = null
    deferred?.resolve()
  }

  fail(reason: unknown): void {
    const deferred = this.deferred
    this.deferred = null
    deferred?.reject(reason)
  }

  destroy(): void {
    if (this.destroyed) {
      return
    }
    this.destroyed = true
    this.mount.remove()
  }
}

export class FakeDocxViewer extends FakeViewer {}
export class FakeXlsxViewer extends FakeViewer {}
export class FakePptxViewer extends FakeViewer {}
