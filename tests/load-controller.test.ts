import { describe, expect, it } from 'vitest'

import { LoadController } from '../src'

describe('LoadController', () => {
  it('returns a non-aborted request with a unique id', () => {
    const controller = new LoadController()
    const request = controller.beginRequest()

    expect(request.id).toBeGreaterThan(0)
    expect(request.signal.aborted).toBe(false)

    request.detach()
  })

  it('cancels the previous request when a new one begins', () => {
    const controller = new LoadController()
    const first = controller.beginRequest()
    const second = controller.beginRequest()

    expect(first.signal.aborted).toBe(true)
    expect(second.signal.aborted).toBe(false)
    expect(first.id).toBeLessThan(second.id)

    first.detach()
    second.detach()
  })

  it('propagates an already-aborted external signal', () => {
    const controller = new LoadController()
    const external = new AbortController()
    external.abort('reason')

    const request = controller.beginRequest(external.signal)

    expect(request.signal.aborted).toBe(true)
    request.detach()
  })

  it('propagates external abort after the request begins', () => {
    const controller = new LoadController()
    const external = new AbortController()
    const request = controller.beginRequest(external.signal)

    expect(request.signal.aborted).toBe(false)
    external.abort('external reason')
    expect(request.signal.aborted).toBe(true)

    request.detach()
  })

  it('isCurrent returns true for the latest request and false for superseded ones', () => {
    const controller = new LoadController()
    const first = controller.beginRequest()
    expect(controller.isCurrent(first.id)).toBe(true)

    const second = controller.beginRequest()
    expect(controller.isCurrent(first.id)).toBe(false)
    expect(controller.isCurrent(second.id)).toBe(true)

    first.detach()
    second.detach()
  })

  it('dispose cancels the active request and invalidates the current generation', () => {
    const controller = new LoadController()
    const request = controller.beginRequest()

    controller.dispose()

    expect(request.signal.aborted).toBe(true)
    expect(controller.isCurrent(request.id)).toBe(false)

    request.detach()
  })

  it('beginRequest after dispose starts a fresh generation', () => {
    const controller = new LoadController()
    const first = controller.beginRequest()
    controller.dispose()
    const second = controller.beginRequest()

    expect(second.id).toBeGreaterThan(first.id)
    expect(second.signal.aborted).toBe(false)

    first.detach()
    second.detach()
  })
})
