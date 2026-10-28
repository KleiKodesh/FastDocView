<script setup lang="ts">
import { computed, onBeforeUnmount, ref, shallowRef } from 'vue'
import { useIntervalFn } from '@vueuse/core'
import WindowedScroller from '@/components/WindowedScroller.vue'
import type { WindowedScrollerApi } from '@/components/windowedScrollerTypes'
import {
  closeDocument, fetchBlocks, fetchServerMemory, fetchWeights, openDocument,
  type WordDocumentSummary, type WordServerMemory,
} from './wordDemoApi'

// The reader: the app's WindowedScroller over blocks fetched on demand from the C# host, the way
// the lines view fetches lines from SQLite. Only the blocks near the viewport are ever held;
// the cache is trimmed to a window around the last request.

const BLOCK_CACHE_RADIUS = 1500

const summary = shallowRef<WordDocumentSummary | null>(null)
const weights = shallowRef<number[] | null>(null)
const blockHtmlByIndex = new Map<number, string>()
const renderVersion = ref(0)          // bumps when cached html changes so rows re-render
const status = ref('No document')
const lastFetch = ref('')
const serverMemory = ref<WordServerMemory | null>(null)
const isOpening = ref(false)
const scrollerRef = ref<InstanceType<typeof WindowedScroller> | null>(null)
const scroller = (): WindowedScrollerApi | null => scrollerRef.value
let currentPath: string | null = null
let requestGeneration = 0
// Where the scroller lands when it (re)mounts — the lines view's initialIndex/initialOffsetPx.
// Read from the live scroller just before a reload, so the same block sits at the same pixel.
const restoreIndex = ref(0)
const restoreOffsetPx = ref(0)

const landmarkTiers = computed(() => {
  if (!summary.value) return []
  const tiers: number[][] = []
  for (let level = 1; level <= 3; level++) {
    const starts = summary.value.headings.filter(h => h.level === level).map(h => h.blockIndex)
    if (starts.length) tiers.push(starts)
  }
  return tiers
})

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
  blockHtmlByIndex.clear()
  if (previous) void closeDocument(previous.id)
  try {
    const opened = await openDocument(path, options)
    if (generation !== requestGeneration) { void closeDocument(opened.id); return }
    currentPath = path
    const weightsStarted = performance.now()
    const fetchedWeights = await fetchWeights(opened.id)
    if (generation !== requestGeneration) { void closeDocument(opened.id); return }
    weights.value = fetchedWeights
    summary.value = opened
    status.value = `${opened.fileName}: ${opened.blockCount.toLocaleString()} blocks, ${opened.readerType}, ${opened.mode}${opened.hasCachedIndex ? ' (cached)' : ''} — opened in ${opened.openMilliseconds.toFixed(0)} ms, headings ${opened.headingsMilliseconds.toFixed(0)} ms, weights ${(performance.now() - weightsStarted).toFixed(0)} ms — ${opened.reason}`
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
  const { blocks, renderMilliseconds, roundTripMilliseconds } = await fetchBlocks(document.id, first, last - first)
  if (summary.value !== document) return
  for (const block of blocks) blockHtmlByIndex.set(block.index, block.html)
  trimCache(first)
  renderVersion.value++
  lastFetch.value = `blocks ${first.toLocaleString()}–${(last - 1).toLocaleString()}: render ${renderMilliseconds} ms, round trip ${roundTripMilliseconds.toFixed(0)} ms, cached ${blockHtmlByIndex.size}`
}

function trimCache(around: number) {
  if (blockHtmlByIndex.size <= BLOCK_CACHE_RADIUS * 2) return
  for (const index of blockHtmlByIndex.keys())
    if (Math.abs(index - around) > BLOCK_CACHE_RADIUS) blockHtmlByIndex.delete(index)
}

function blockHtml(index: number): string {
  void renderVersion.value
  return blockHtmlByIndex.get(index) ?? ''
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
    document: summary.value && { id: summary.value.id, blockCount: summary.value.blockCount, mode: summary.value.mode, openMilliseconds: summary.value.openMilliseconds },
    renderedRange: api?.renderedRange ?? null,
    firstVisible: api?.readFirstVisible() ?? null,
    cachedBlocks: blockHtmlByIndex.size,
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
      <label>Go to block <input type="number" min="0" :max="(summary?.blockCount ?? 1) - 1" data-test="goto-input" @change="scrollToBlock(Number(($event.target as HTMLInputElement).value))" /></label>
      <span v-if="serverMemory" class="memory" data-test="server-memory">
        server heap {{ serverMemory.managedHeapMegabytes }} MB · process {{ serverMemory.workingSetMegabytes }} MB
      </span>
    </header>
    <p class="status" data-test="status">{{ status }}</p>
    <p class="status fetch" data-test="last-fetch">{{ lastFetch }}</p>

    <div class="body">
      <nav v-if="summary" class="headings">
        <button
          v-for="heading in summary.headings"
          :key="heading.blockIndex"
          :style="{ paddingInlineStart: `${8 + (heading.level - 1) * 12}px` }"
          :title="`block ${heading.blockIndex}`"
          data-test="heading"
          @click="scrollToBlock(heading.blockIndex)"
        >{{ heading.text }}</button>
      </nav>
      <component :is="'style'" v-if="summary">{{ summary.stylesheet }}</component>
      <WindowedScroller
        v-if="summary && weights"
        ref="scrollerRef"
        class="scroller word-document"
        :count="summary.blockCount"
        :initial-index="restoreIndex"
        :initial-offset-px="restoreOffsetPx"
        :ensure-rows="ensureBlocks"
        :row-source-version="renderVersion"
        :row-weights="weights"
        :landmark-tiers="landmarkTiers"
        :dir="summary.isRightToLeft ? 'rtl' : 'ltr'"
        tabindex="0"
        data-test="scroller"
      >
        <template #row="{ index }">
          <div class="block" :data-block="index" v-html="blockHtml(index)"></div>
        </template>
      </WindowedScroller>
      <div v-else class="placeholder">{{ isOpening ? 'Opening…' : 'Pick a file on the left.' }}</div>
    </div>
  </section>
</template>

<style scoped>
.reader-panel { display: flex; flex-direction: column; height: 100%; }
.toolbar { display: flex; gap: 10px; align-items: center; padding: 6px 10px; background: #fff; border-bottom: 1px solid #ddd; font-size: 13px; }
.toolbar input { width: 90px; }
.memory { margin-inline-start: auto; color: #666; }
.status { margin: 0; padding: 4px 10px; font-size: 12px; color: #444; background: #fff; border-bottom: 1px solid #eee; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.status.fetch { color: #777; }
.body { flex: 1; min-height: 0; display: flex; }
.headings { width: 220px; flex: none; overflow: auto; border-inline-end: 1px solid #ddd; background: #fafafa; display: flex; flex-direction: column; }
.headings button { text-align: start; border: 0; background: none; padding: 3px 8px; cursor: pointer; font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; direction: rtl; }
.headings button:hover { background: #e8eefc; }
.scroller { flex: 1; min-width: 0; background: #fff; }
.block { padding: 0 48px; max-width: 760px; margin: 0 auto; }
.placeholder { flex: 1; display: grid; place-items: center; color: #888; }
</style>
