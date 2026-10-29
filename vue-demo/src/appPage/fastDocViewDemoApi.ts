// Stands in for the app's `@/webview-host/fastDocViewApi` (vite.config.ts aliases that module here),
// so the app's own FastDocView page runs unchanged against this repo's C# dev server. Everything else
// the module exports (types, the source-changed marker, the weights decoder) is the app's own code.
import { reactive } from 'vue'
import { FAST_DOC_VIEW_SOURCE_CHANGED, decodeFastDocViewWeights } from '@app-webview-host/fastDocViewApi'
import type {
  FastDocViewApi, FastDocViewBlock, FastDocViewHeading, FastDocViewScanState, FastDocViewSummary,
} from '@app-webview-host/fastDocViewApi'

export * from '@app-webview-host/fastDocViewApi'

/** The dev server answers 409 when the file changed on disk: the page reopens on this marker, as with the service. */
async function readJson<T>(response: Response): Promise<T> {
  if (response.status === 409) throw new Error(FAST_DOC_VIEW_SOURCE_CHANGED)
  if (!response.ok) {
    let message = `${response.status} ${response.statusText}`
    try { message = ((await response.json()) as { error?: string }).error ?? message } catch { /* not json */ }
    throw new Error(message)
  }
  return (await response.json()) as T
}

/** WordDocumentBlockKind values -> the names the app's Core sends (FastDocViewOpenDocuments.BlockKindName). */
const BLOCK_KIND_NAMES: Record<number, string> = { 0: 'paragraph', 1: 'table', 2: 'contentControl', 3: 'other', 4: 'tableRow' }

interface DevServerSummary {
  id: number; blockCount: number; knownBlockCount: number; isIndexComplete: boolean; isRightToLeft: boolean
  stylesheet: string; headingCount: number; footnoteCount: number; endnoteCount: number
  openMilliseconds: number; mode: string; hasCachedIndex: boolean; readerType: string; reason: string
}

/** What the demo host shows beside the app page: the last open's server-side facts (not part of the app's API). */
/** Demo-only open options for the next open (the Playwright checks force a cold first open): consumed by open(). */
export const demoOpenOptions: { rebuildIndex?: boolean; forceSidecar?: boolean } = {}

export const demoDiagnostics = reactive({ lastOpen: null as null | { fileName: string; openMilliseconds: number; mode: string; hasCachedIndex: boolean; readerType: string; reason: string } })

export const fastDocViewApi: FastDocViewApi = {
  async open(path) {
    const summary = await readJson<DevServerSummary>(await fetch('/api/documents', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ path, rebuildIndex: demoOpenOptions.rebuildIndex, forceSidecar: demoOpenOptions.forceSidecar }),
    }))
    demoOpenOptions.rebuildIndex = demoOpenOptions.forceSidecar = undefined   // once
    demoDiagnostics.lastOpen = { fileName: path.split(/[\/]/).pop() ?? path, openMilliseconds: summary.openMilliseconds, mode: summary.mode,
                                 hasCachedIndex: summary.hasCachedIndex, readerType: summary.readerType, reason: summary.reason }
    const result: FastDocViewSummary = {
      documentId: String(summary.id), blockCount: summary.blockCount, knownBlockCount: summary.knownBlockCount,
      isIndexComplete: summary.isIndexComplete, isRightToLeft: summary.isRightToLeft, stylesheet: summary.stylesheet,
      headingCount: summary.headingCount, footnoteCount: summary.footnoteCount, endnoteCount: summary.endnoteCount,
    }
    return result
  },
  async blocks(documentId, first, count) {
    const blocks = await readJson<{ index: number; kind: number; html: string; notes: FastDocViewBlock['notes'] }[]>(
      await fetch(`/api/documents/${documentId}/blocks?first=${first}&count=${count}`))
    return blocks.map(block => ({ index: block.index, kind: BLOCK_KIND_NAMES[block.kind] ?? 'other', html: block.html, notes: block.notes }))
  },
  async sections(documentId) {
    return readJson<{ firstBlockIndex: number; textWidthTwips: number }[]>(await fetch(`/api/documents/${documentId}/sections`))
  },
  async bookmark(documentId, name) {
    const reply = await readJson<{ blockIndex: number | null }>(await fetch(`/api/documents/${documentId}/bookmark?name=${encodeURIComponent(name)}`))
    return reply.blockIndex
  },
  async headings(documentId) {
    return readJson<FastDocViewHeading[]>(await fetch(`/api/documents/${documentId}/headings`))
  },
  async weights(documentId) {
    const response = await fetch(`/api/documents/${documentId}/weights`)
    if (response.status === 409) throw new Error(FAST_DOC_VIEW_SOURCE_CHANGED)
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`)
    return decodeFastDocViewWeights(new Uint8Array(await response.arrayBuffer()))
  },
  async scan(documentId, blocks) {
    return readJson<FastDocViewScanState>(await fetch(`/api/documents/${documentId}/scan?blocks=${blocks}`, { method: 'POST' }))
  },
  async close(documentId) {
    await fetch(`/api/documents/${documentId}`, { method: 'DELETE' })
  },
}
