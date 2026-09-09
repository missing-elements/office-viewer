import { describe, expect, it } from 'vitest'

import { defineOfficeViewerElement, OfficeViewerElement } from '../src'

describe('office-viewer element', () => {
  it('exports the headless custom-element surface', () => {
    expect(OfficeViewerElement).toBeTypeOf('function')
    expect(defineOfficeViewerElement).toBeTypeOf('function')
  })

  it('creates a Shadow DOM when created', () => {
    defineOfficeViewerElement()
    const element = document.createElement('office-viewer') as OfficeViewerElement
    document.body.append(element)
    expect(element.shadowRoot).not.toBeNull()
    expect(element.shadowRoot?.getElementById('viewer')).not.toBeNull()
    element.remove()
  })

  it('reflects idle status before load', () => {
    defineOfficeViewerElement()
    const element = document.createElement('office-viewer') as OfficeViewerElement
    expect(element.ready).toBe(false)
    expect(element.error).toBeNull()
    expect(element.format).toBeNull()
    expect(element.mode).toBeNull()
    expect(element.getViewer()).toBeNull()
  })
})
