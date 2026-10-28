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
  headings: WordHeading[]
  openMilliseconds: number
  headingsMilliseconds: number
  mode: string
  hasCachedIndex: boolean
  reason: string
  readerType: string
}

/** The block shape the app's lines view consumes: index + html. */
export interface WordBlock { index: number; kind: number; html: string }

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

export function fetchWeights(id: number): Promise<number[]> {
  return fetch(`/api/documents/${id}/weights`).then(readJson<number[]>)
}

export async function fetchBlocks(id: number, first: number, count: number): Promise<{ blocks: WordBlock[]; renderMilliseconds: number; roundTripMilliseconds: number }> {
  const started = performance.now()
  const response = await fetch(`/api/documents/${id}/blocks?first=${first}&count=${count}`)
  const blocks = await readJson<WordBlock[]>(response)
  return {
    blocks,
    renderMilliseconds: Number(response.headers.get('X-Render-Milliseconds') ?? 0),
    roundTripMilliseconds: performance.now() - started,
  }
}

export function fetchProgress(): Promise<{ bytesDone: number; bytesTotal: number; blocksFound: number }> {
  return fetch('/api/progress').then(readJson<{ bytesDone: number; bytesTotal: number; blocksFound: number }>)
}

export function fetchServerMemory(collect = false): Promise<WordServerMemory> {
  return fetch(`/api/memory${collect ? '?collect=true' : ''}`).then(readJson<WordServerMemory>)
}
