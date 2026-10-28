<script setup lang="ts">
import { computed, onBeforeUnmount, ref, shallowRef } from 'vue'
import { useIntervalFn, useThrottleFn } from '@vueuse/core'
import WindowedScroller from '@/components/WindowedScroller.vue'
import type { WindowedScrollerApi } from '@/components/windowedScrollerTypes'
import WordNotesPane from './WordNotesPane.vue'
import {
  closeDocument, fetchBlocks, fetchHeadings, fetchServerMemory, fetchWeights, openDocument, WordSourceChangedError,
  type WordDocumentSummary, type WordHeading, type WordIndexState, type WordNote, type WordServerMemory, type WordVisibleNote,
} from './wordDemoApi'

// The reader: the app's WindowedScroller over blocks fetched on demand from the C# host, the way
// the lines view fetches lines from SQLite. Only the blocks near the viewport are ever held; the
// cache is trimmed to a window around the last request. Text mounts as soon as the document is
// open; headings and scrollbar weights arrive in parallel and are applied when they land.

const BLOCK_CACHE_RADIUS = 1500

// A .docx first open builds its index on demand: the server scans only as far as the blocks asked
// for, so the block count starts as an estimate and every blocks response carries the corrected
// one (no polling). When the scan completes, headings and weights are fetched again in full.
const blockCount = ref(0)
const isIndexComplete = ref(true)
const knownBlockCount = ref(0)

const summary = shallowRef<WordDocumentSummary | null>(null)
const weights = shallowRef<Int32Array | null>(null)
const headings = shallowRef<WordHeading[]>([])
const blockHtmlByIndex = new Map<number, string>()
const blockNotesByIndex = new Map<number, WordNote[]>()
const renderVersion = ref(0)          // bumps when cached html changes so rows re-render
const status = ref('No document')
const lastFetch = ref('')
const serverMemory = ref<WordServerMemory | null>(null)
const isOpening = ref(false)
const visibleNotes = shallowRef<WordVisibleNote[]>([])
const highlightedNoteKey = ref<string | null>(null)
const tooltip = ref<{ html: string; x: number; y: number } | null>(null)
const scrollerRef = ref<InstanceType<typeof WindowedScroller> | null>(null)
const scroller = (): WindowedScrollerApi | null => scrollerRef.value
let currentPath: string | null = null
let requestGeneration = 0
// Where the scroller lands when it (re)mounts — the lines view's initialIndex/initialOffsetPx.
// Read from the live scroller just before a reload, so the same block sits at the same pixel.
const restoreIndex = ref(0)
const restoreOffsetPx = ref(0)

const landmarkTiers = computed(() => {
  const tiers: number[][] = []
  for (let level = 1; level <= 3; level++) {
    const starts = headings.value.filter(h => h.level === level).map(h => h.blockIndex)
    if (starts.length) tiers.push(starts)
  }
  return tiers
})

const hasNotes = computed(() => (summary.value?.footnoteCount ?? 0) + (summary.value?.endnoteCount ?? 0) > 0)

async function open(path: string, options: { forceSidecar?: boolean; rebuildIndex?: boolean; keepPosition?: boolean } = {}) {
  const generation = ++requestGeneration
  const visible = options.keepPosition ? scroller()?.readFirstVisible() : null
  restoreIndex.value = visible?.index ?? 0
  restoreOffsetPx.value = visible?.offsetWithinRow ?? 0
  isOpening.value = true
  status.value = `Opening ${path}…`
  const previous = summary.value
  summary.value = null
  weights.value = null
  headings.value = []
  visibleNotes.value = []
  blockHtmlByIndex.clear()
  blockNotesByIndex.clear()
  if (previous) void closeDocument(previous.id)
  const started = performance.now()
  try {
    const opened = await openDocument(path, options)
    if (generation !== requestGeneration) { void closeDocument(opened.id); return }
    currentPath = path
    blockCount.value = opened.blockCount
    knownBlockCount.value = opened.knownBlockCount
    isIndexComplete.value = opened.isIndexComplete
    summary.value = opened                       // text mounts now; weights + headings follow
    const textShownAt = performance.now()
    const [fetchedWeights, fetchedHeadings] = await Promise.all([fetchWeights(opened.id), fetchHeadings(opened.id)])
    if (generation !== requestGeneration) return
    weights.value = fetchedWeights
    headings.value = fetchedHeadings
    status.value = `${opened.fileName}: ${opened.blockCount.toLocaleString()} blocks, ${opened.readerType}, ${opened.mode}${opened.hasCachedIndex ? ' (cached)' : ''} — `
      + `server open ${opened.openMilliseconds.toFixed(0)} ms, text after ${(textShownAt - started).toFixed(0)} ms, headings+weights after ${(performance.now() - started).toFixed(0)} ms — `
      + `${opened.headingCount} headings, ${opened.footnoteCount} footnotes, ${opened.endnoteCount} endnotes — ${opened.reason}`
    void refreshMemory()
  } catch (exception) {
    status.value = `Failed: ${(exception as Error).message}`
  } finally {
    if (generation === requestGeneration) isOpening.value = false
  }
}

function reload() {
  if (currentPath) void open(currentPath, { keepPosition: true })
}

async function ensureBlocks(start: number, end: number): Promise<void> {
  const document = summary.value
  if (!document) return
  let first = start
  while (first < end && blockHtmlByIndex.has(first)) first++
  let last = end
  while (last > first && blockHtmlByIndex.has(last - 1)) last--
  if (first >= last) return
  let fetched: Awaited<ReturnType<typeof fetchBlocks>>
  try {
    fetched = await fetchBlocks(document.id, first, last - first)
  } catch (exception) {
    // Word saved the file: reopen (re-index) and stay where the reader is.
    if (exception instanceof WordSourceChangedError && summary.value === document && currentPath) {
      status.value = 'The document changed on disk — reopening…'
      void open(currentPath, { keepPosition: true })
      return
    }
    throw exception
  }
  const { blocks, renderMilliseconds, roundTripMilliseconds, indexState } = fetched
  if (summary.value !== document) return
  if (indexState) applyIndexState(document, indexState)
  for (const block of blocks) {
    blockHtmlByIndex.set(block.index, block.html)
    if (block.notes.length) blockNotesByIndex.set(block.index, block.notes)
  }
  trimCache(first)
  renderVersion.value++
  lastFetch.value = `blocks ${first.toLocaleString()}–${(last - 1).toLocaleString()}: render ${renderMilliseconds} ms, round trip ${roundTripMilliseconds.toFixed(0)} ms, cached ${blockHtmlByIndex.size}`
  requestAnimationFrame(() => refreshVisibleNotes())
}

/** Takes the count / completion a blocks response reports; refetches headings + weights once the scan is complete. */
function applyIndexState(document: WordDocumentSummary, state: WordIndexState) {
  // the scan only moves forward: a response older than one already applied (or any after completion) is stale
  if (isIndexComplete.value || state.knownBlockCount < knownBlockCount.value) return
  knownBlockCount.value = state.knownBlockCount
  if (blockCount.value !== state.estimatedBlockCount) blockCount.value = state.estimatedBlockCount
  if (!state.isIndexComplete) return
  isIndexComplete.value = true
  void Promise.all([fetchWeights(document.id), fetchHeadings(document.id)]).then(([fetchedWeights, fetchedHeadings]) => {
    if (summary.value !== document) return
    weights.value = fetchedWeights
    headings.value = fetchedHeadings
  })
}

/** Weights only while they describe the current count (an estimate moves; the complete set is fetched on completion). */
const rowWeights = computed(() => (weights.value && weights.value.length === blockCount.value ? weights.value : undefined))

function trimCache(around: number) {
  if (blockHtmlByIndex.size <= BLOCK_CACHE_RADIUS * 2) return
  for (const index of blockHtmlByIndex.keys()) {
    if (Math.abs(index - around) <= BLOCK_CACHE_RADIUS) continue
    blockHtmlByIndex.delete(index)
    blockNotesByIndex.delete(index)
  }
}

function blockHtml(index: number): string {
  void renderVersion.value
  return blockHtmlByIndex.get(index) ?? ''
}

function noteKey(blockIndex: number, note: { kind: string; number: number }, position: number) {
  return `${note.kind}:${note.number || `${blockIndex}.${position}`}`
}

/** The notes referenced by the rows that are actually on screen — what the notes pane lists. */
function refreshVisibleNotes() {
  const api = scroller()
  const viewport = api?.scrollerElement?.getBoundingClientRect()
  if (!api || !viewport) return
  const notes: WordVisibleNote[] = []
  for (let index = api.renderedRange.start; index < api.renderedRange.end; index++) {
    const blockNotes = blockNotesByIndex.get(index)
    if (!blockNotes) continue
    const row = api.rowElement(index)?.getBoundingClientRect()
    if (!row || row.bottom < viewport.top || row.top > viewport.bottom) continue
    blockNotes.forEach((note, position) => notes.push({ key: noteKey(index, note, position), blockIndex: index, ...note }))
  }
  visibleNotes.value = notes
}
const onScrollerScroll = useThrottleFn(refreshVisibleNotes, 100, true)

function referenceKeyOf(target: EventTarget | null): { key: string; html: string } | null {
  const reference = (target as HTMLElement | null)?.closest?.('.note-reference') as HTMLElement | null
  if (!reference) return null
  const row = reference.closest('[data-block]') as HTMLElement | null
  const blockIndex = Number(row?.dataset.block)
  const kind = reference.dataset.noteKind ?? 'footnote'
  const number = Number(reference.dataset.noteNumber ?? 0)
  const notes = blockNotesByIndex.get(blockIndex) ?? []
  const position = Math.max(0, notes.findIndex(note => note.kind === kind && note.number === number))
  const note = notes[position]
  return note ? { key: noteKey(blockIndex, note, position), html: note.html } : null
}

function onContentClick(event: MouseEvent) {
  const found = referenceKeyOf(event.target)
  if (found) { refreshVisibleNotes(); highlightedNoteKey.value = found.key }
}

function onContentPointerOver(event: PointerEvent) {
  const found = referenceKeyOf(event.target)
  tooltip.value = found ? { html: found.html, x: event.clientX, y: event.clientY } : null
}

function scrollToBlock(index: number) {
  return scroller()?.scrollToIndex(index)
}

async function refreshMemory(collect = false) {
  try { serverMemory.value = await fetchServerMemory(collect) } catch { /* server gone */ }
}
useIntervalFn(() => void refreshMemory(), 2000)

function debugState() {
  const api = scroller()
  return {
    document: summary.value && {
      id: summary.value.id, blockCount: blockCount.value, mode: summary.value.mode, openMilliseconds: summary.value.openMilliseconds,
      openedBlockCount: summary.value.blockCount, knownBlockCount: knownBlockCount.value, isIndexComplete: isIndexComplete.value,
      footnoteCount: summary.value.footnoteCount, endnoteCount: summary.value.endnoteCount,
    },
    hasWeights: weights.value !== null,
    headingCount: headings.value.length,
    renderedRange: api?.renderedRange ?? null,
    firstVisible: api?.readFirstVisible() ?? null,
    cachedBlocks: blockHtmlByIndex.size,
    visibleNotes: visibleNotes.value.map(note => ({ key: note.key, label: note.label, textLength: note.html.replace(/<[^>]+>/g, '').length })),
    lastFetch: lastFetch.value,
    serverMemory: serverMemory.value,
  }
}

onBeforeUnmount(() => { if (summary.value) void closeDocument(summary.value.id) })

defineExpose({ open, reload, scrollToBlock, debugState })
</script>

<template>
  <section class="reader-panel">
    <header class="toolbar">
      <button :disabled="!summary || isOpening" data-test="reload-button" @click="reload">Reload</button>
      <button :disabled="!summary || isOpening" @click="currentPath && open(currentPath, { rebuildIndex: true, forceSidecar: true, keepPosition: true })">Rebuild index</button>
      <button @click="refreshMemory(true)">Collect GC</button>
      <label>Go to block <input type="number" min="0" :max="blockCount - 1" data-test="goto-input" @change="scrollToBlock(Number(($event.target as HTMLInputElement).value))" /></label>
      <span v-if="summary && !isIndexComplete" data-test="index-state">indexing on demand: {{ knownBlockCount.toLocaleString() }} of ~{{ blockCount.toLocaleString() }} blocks known</span>
      <span v-if="serverMemory" class="memory" data-test="server-memory">
        server heap {{ serverMemory.managedHeapMegabytes }} MB · process {{ serverMemory.workingSetMegabytes }} MB
      </span>
    </header>
    <p class="status" data-test="status">{{ status }}</p>
    <p class="status fetch" data-test="last-fetch">{{ lastFetch }}</p>

    <div class="body">
      <nav v-if="summary" class="headings">
        <button
          v-for="heading in headings"
          :key="heading.blockIndex"
          :style="{ paddingInlineStart: `${8 + (heading.level - 1) * 12}px` }"
          :title="`block ${heading.blockIndex}`"
          data-test="heading"
          @click="scrollToBlock(heading.blockIndex)"
        >{{ heading.text }}</button>
      </nav>
      <component :is="'style'" v-if="summary">{{ summary.stylesheet }}</component>
      <div v-if="summary" class="text-and-notes">
        <WindowedScroller
          ref="scrollerRef"
          class="scroller word-document"
          :count="blockCount"
          :initial-index="restoreIndex"
          :initial-offset-px="restoreOffsetPx"
          :ensure-rows="ensureBlocks"
          :row-source-version="renderVersion"
          :row-weights="rowWeights"
          :landmark-tiers="landmarkTiers"
          :dir="summary.isRightToLeft ? 'rtl' : 'ltr'"
          tabindex="0"
          data-test="scroller"
          @scroll="onScrollerScroll"
          @click="onContentClick"
          @pointerover="onContentPointerOver"
          @pointerleave="tooltip = null"
        >
          <template #row="{ index }">
            <div class="block" :data-block="index" v-html="blockHtml(index)"></div>
          </template>
        </WindowedScroller>
        <WordNotesPane
          v-if="hasNotes"
          class="notes"
          :notes="visibleNotes"
          :highlighted-key="highlightedNoteKey"
          :is-right-to-left="summary.isRightToLeft"
          :footnote-count="summary.footnoteCount"
          :endnote-count="summary.endnoteCount"
        />
      </div>
      <div v-else class="placeholder">{{ isOpening ? 'Opening…' : 'Pick a file on the left.' }}</div>
    </div>
    <div
      v-if="tooltip"
      class="note-tooltip word-document"
      :dir="summary?.isRightToLeft ? 'rtl' : 'ltr'"
      :style="{ left: `${tooltip.x + 12}px`, top: `${tooltip.y + 16}px` }"
      data-test="note-tooltip"
      v-html="tooltip.html"
    ></div>
  </section>
</template>

<style scoped>
.reader-panel { display: flex; flex-direction: column; height: 100%; position: relative; }
.toolbar { display: flex; gap: 10px; align-items: center; padding: 6px 10px; background: #fff; border-bottom: 1px solid #ddd; font-size: 13px; }
.toolbar input { width: 90px; }
.memory { margin-inline-start: auto; color: #666; }
.status { margin: 0; padding: 4px 10px; font-size: 12px; color: #444; background: #fff; border-bottom: 1px solid #eee; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.status.fetch { color: #777; }
.body { flex: 1; min-height: 0; display: flex; }
.headings { width: 220px; flex: none; overflow: auto; border-inline-end: 1px solid #ddd; background: #fafafa; display: flex; flex-direction: column; }
.headings button { text-align: start; border: 0; background: none; padding: 3px 8px; cursor: pointer; font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; direction: rtl; }
.headings button:hover { background: #e8eefc; }
.text-and-notes { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.scroller { flex: 1; min-height: 0; background: #fff; }
.notes { height: 30%; flex: none; }
.block { padding: 0 48px; max-width: 760px; margin: 0 auto; }
.placeholder { flex: 1; display: grid; place-items: center; color: #888; }
.note-tooltip { position: fixed; z-index: 10; max-width: 420px; max-height: 240px; overflow: hidden; padding: 6px 10px; background: #fffef6; border: 1px solid #d8d2a8; box-shadow: 0 2px 8px rgba(0,0,0,.15); font-size: 0.9em; pointer-events: none; }
</style>
