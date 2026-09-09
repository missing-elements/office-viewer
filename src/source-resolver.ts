import type { OfficeFormat, OfficeSource, OfficeSourceKind } from './types'

const EXTENSION_TO_FORMAT: Record<string, OfficeFormat> = {
  docx: 'docx',
  xlsx: 'xlsx',
  pptx: 'pptx'
}

export interface ResolvedSource {
  readonly source: string | ArrayBuffer
  readonly sourceKind: OfficeSourceKind
  readonly format: OfficeFormat
  readonly fileName?: string
}

export async function resolveOfficeSource(
  source: OfficeSource,
  fileName: string | undefined,
  explicitFormat: OfficeFormat | undefined
): Promise<ResolvedSource> {
  if (source instanceof File) {
    const resolvedName = fileName ?? source.name
    return {
      source: copyArrayBuffer(await source.arrayBuffer()),
      sourceKind: 'file',
      format: detectOfficeFormat(resolvedName, explicitFormat),
      fileName: resolvedName
    }
  }

  if (source instanceof Blob) {
    const resolvedName = fileName
    return {
      source: copyArrayBuffer(await source.arrayBuffer()),
      sourceKind: 'blob',
      format: detectOfficeFormat(resolvedName, explicitFormat),
      fileName: resolvedName
    }
  }

  if (source instanceof Uint8Array) {
    const resolvedName = fileName
    return {
      source: source.slice().buffer,
      sourceKind: 'uint8array',
      format: detectOfficeFormat(resolvedName, explicitFormat),
      fileName: resolvedName
    }
  }

  if (source instanceof ArrayBuffer) {
    const resolvedName = fileName
    return {
      source: copyArrayBuffer(source),
      sourceKind: 'array-buffer',
      format: detectOfficeFormat(resolvedName, explicitFormat),
      fileName: resolvedName
    }
  }

  if (source instanceof URL) {
    const href = source.toString()
    const resolvedName = fileName ?? inferFileNameFromUrl(href)
    return {
      source: href,
      sourceKind: 'url',
      format: detectOfficeFormat(resolvedName, explicitFormat),
      fileName: resolvedName
    }
  }

  // String — treat as URL.
  const resolvedName = fileName ?? inferFileNameFromUrl(source)
  return {
    source,
    sourceKind: 'url',
    format: detectOfficeFormat(resolvedName, explicitFormat),
    fileName: resolvedName
  }
}

export function detectOfficeFormat(fileName: string | undefined, explicitFormat?: OfficeFormat): OfficeFormat {
  if (explicitFormat) {
    return explicitFormat
  }

  const extension = fileName?.split('.').pop()?.toLowerCase()
  if (extension && extension in EXTENSION_TO_FORMAT) {
    return EXTENSION_TO_FORMAT[extension]
  }

  throw new Error(
    'Unable to determine Office format. Provide options.format, file-name, or a URL/file name with .docx, .xlsx, or .pptx.'
  )
}

function inferFileNameFromUrl(urlString: string): string | undefined {
  try {
    const url = new URL(urlString, typeof location !== 'undefined' ? location.href : undefined)
    const parts = url.pathname.split('/').filter(Boolean)
    return parts.at(-1) || undefined
  } catch {
    return undefined
  }
}

function copyArrayBuffer(buffer: ArrayBuffer): ArrayBuffer {
  return buffer.slice(0)
}
