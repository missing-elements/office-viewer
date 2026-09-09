export interface ControlledRequest {
  readonly id: number
  readonly signal: AbortSignal
  readonly detach: () => void
}

export class LoadController {
  private generation = 0
  private activeController: AbortController | null = null

  /**
   * Starts a new request generation. Cancels any previous request and returns
   * a controller-backed signal. The caller must call `detach()` when done.
   */
  beginRequest(externalSignal?: AbortSignal): ControlledRequest {
    const id = ++this.generation
    this.cancel(createAbortError('Superseded by a newer request.'))

    const controller = new AbortController()
    this.activeController = controller

    let removeExternalListener = () => {}
    if (externalSignal) {
      if (externalSignal.aborted) {
        controller.abort(externalSignal.reason)
      } else {
        const onExternalAbort = () => {
          controller.abort(externalSignal.reason)
        }
        externalSignal.addEventListener('abort', onExternalAbort, { once: true })
        removeExternalListener = () => {
          externalSignal.removeEventListener('abort', onExternalAbort)
        }
      }
    }

    return {
      id,
      signal: controller.signal,
      detach: () => {
        removeExternalListener()
        if (this.activeController === controller) {
          this.activeController = null
        }
      }
    }
  }

  /**
   * Cancels the active request, if any.
   */
  cancel(reason: unknown): void {
    if (!this.activeController) {
      return
    }

    this.activeController.abort(reason)
    this.activeController = null
  }

  /**
   * Returns true if `id` matches the current generation.
   */
  isCurrent(id: number): boolean {
    return this.generation === id
  }

  /**
   * Permanently invalidates any current and future requests.
   */
  dispose(): void {
    this.cancel(createAbortError('LoadController disposed.'))
    this.generation += 1
  }
}

function createAbortError(reason: string): DOMException {
  return new DOMException(reason, 'AbortError')
}
