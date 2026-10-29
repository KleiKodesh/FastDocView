// Thin client for the C# dev server (proxied under /api). Shapes mirror DevServer/Program.cs.

export interface WordFolderEntry {
  path: string
  fileName: string
  fileBytes: number
  mode: 'InMemory' | 'Sidecar'
  hasCachedIndex: boolean
  bodyBytes: number
  estimatedIndexBytes: number
  reason: string
}

export interface WordHeading { blockIndex: number; level: number; text: string }

export interface WordDocumentSummary {
  id: number
  fileName: string
  path: string
  blockCount: number
  isRightToLeft: boolean
  stylesheet: string
  openMilliseconds: number
  mode: string
  hasCachedIndex: boolean
  reason: string
  readerType: string
  headingCount: number
  footnoteCount: number
  endnoteCount: number
  /** Blocks whose position is final; equals blockCount once the index is complete. */
  knownBlockCount: number
  /** False while a .docx first open builds its index on demand (blockCount is then an estimate). */
  isIndexComplete: boolean
}

/** Index state carried by every blocks response (headers): the on-demand scan moves with each request. */
export interface WordIndexState { knownBlockCount: number; isIndexComplete: boolean }
/** A scan reply: the index state plus the notes counted so far (they grow with the index). */
export interface WordScanState extends WordIndexState { footnoteCount: number; endnoteCount: number }

export type WordNoteKind = 'footnote' | 'endnote'

/** A footnote/endnote body, delivered with the block that references it. */
export interface WordNote { kind: WordNoteKind; number: number; label: string; html: string }

/** A note as the notes pane lists it: which block referenced it and a key unique within the document. */
export interface WordVisibleNote extends WordNote { key: string; blockIndex: number }

/** The block shape the app's lines view consumes: index + html (+ the notes it references). */
export interface WordBlock { index: number; kind: number; html: string; notes: WordNote[] }

export interface WordServerMemory { managedHeapMegabytes: number; workingSetMegabytes: number; openDocuments: number }

async function readJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    let message = `${response.status} ${response.statusText}`
    try { message = ((await response.json()) as { error?: string }).error ?? message } catch { /* not json */ }
    throw new Error(message)
  }
  return (await response.json()) as T
}

export function listFolder(path: string): Promise<WordFolderEntry[]> {
  return fetch(`/api/folder?path=${encodeURIComponent(path)}`).then(readJson<WordFolderEntry[]>)
}

export function openDocument(path: string, options: { forceSidecar?: boolean; rebuildIndex?: boolean } = {}): Promise<WordDocumentSummary> {
  return fetch('/api/documents', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ path, ...options }),
  }).then(readJson<WordDocumentSummary>)
}

export function closeDocument(id: number): Promise<Response> {
  return fetch(`/api/documents/${id}`, { method: 'DELETE' })
}

/** Raw little-endian Int32 per block — no JSON for 100k+ numbers. */
export async function fetchWeights(id: number): Promise<Int32Array> {
  const response = await fetch(`/api/documents/${id}/weights`)
  if (!response.ok) throw new Error(`${response.status} ${response.statusText}`)
  return new Int32Array(await response.arrayBuffer())
}

export function fetchHeadings(id: number): Promise<WordHeading[]> {
  return fetch(`/api/documents/${id}/headings`).then(readJson<WordHeading[]>)
}

/** Thrown when the file changed on disk (Word saved it): the caller reopens the document. */
export class WordSourceChangedError extends Error {}

export async function fetchBlocks(id: number, first: number, count: number): Promise<{ blocks: WordBlock[]; renderMilliseconds: number; roundTripMilliseconds: number; indexState: WordIndexState | null }> {
  const started = performance.now()
  const response = await fetch(`/api/documents/${id}/blocks?first=${first}&count=${count}`)
  if (response.status === 409) throw new WordSourceChangedError('The document changed on disk')
  const blocks = await readJson<WordBlock[]>(response)
  const known = response.headers.get('X-Known-Block-Count')
  return {
    blocks,
    indexState: known === null ? null : {
      knownBlockCount: Number(known),
      isIndexComplete: response.headers.get('X-Index-Complete') === 'true',
    },
    renderMilliseconds: Number(response.headers.get('X-Render-Milliseconds') ?? 0),
    roundTripMilliseconds: performance.now() - started,
  }
}

/** Builds `blocks` more blocks of a first open's index on the server (renders and sends nothing) and reports the new state. */
export async function scanAhead(id: number, blocks: number): Promise<WordScanState> {
  const response = await fetch(`/api/documents/${id}/scan?blocks=${blocks}`, { method: 'POST' })
  if (response.status === 409) throw new WordSourceChangedError('The document changed on disk')
  return readJson<WordScanState>(response)
}

export function fetchServerMemory(collect = false): Promise<WordServerMemory> {
  return fetch(`/api/memory${collect ? '?collect=true' : ''}`).then(readJson<WordServerMemory>)
}
