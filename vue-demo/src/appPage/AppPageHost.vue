<script setup lang="ts">
// Hosts the app's FastDocViewPage the way the app's AppShell does: one pane with one tab whose
// localFilePath is the open document (PANE_NAVIGATION_KEY), the page keyed by tab + file so another
// file remounts it - as AppPageView keys it.
import { computed, provide, reactive, ref } from 'vue'
import FastDocViewPage from '@/features/fast-doc-view/FastDocViewPage.vue'
import { PANE_NAVIGATION_KEY, type PaneNavigation } from '@/composables/usePaneNavigation'
import type { Tab } from '@/stores/tabStore'
import { SAMPLE_FILES } from '../wordDemoDefaults'
import { demoDiagnostics } from './fastDocViewDemoApi'

// Server memory on demand (no timers in the page): heap / process / open documents, after an optional GC.
const memory = ref('')
async function readMemory(collect: boolean) {
  try {
    const m = await (await fetch(`/api/memory?collect=${collect}`)).json() as { managedHeapMegabytes: number; workingSetMegabytes: number; openDocuments: number }
    memory.value = `heap ${m.managedHeapMegabytes} MB, process ${m.workingSetMegabytes} MB, ${m.openDocuments} open`
  } catch { memory.value = 'server unreachable' }
}

const LAST_FILE_KEY = 'fastdocview-app-page:last-file'
let lastFile: string | null = null
try { lastFile = localStorage.getItem(LAST_FILE_KEY) } catch { /* private window */ }
const queryFile = new URLSearchParams(location.search).get('file')

const tab = reactive({ id: 'fastdocview-demo-tab', route: '/fast-doc-view', title: '', localFilePath: queryFile || lastFile || SAMPLE_FILES[0]!.path } as Tab)
const pathInput = ref(tab.localFilePath ?? '')
const pageKey = computed(() => `${tab.id}:${tab.localFilePath}`)

function open(path: string) {
  const trimmed = path.trim()
  if (!trimmed) return
  pathInput.value = trimmed
  tab.localFilePath = trimmed
  tab.title = trimmed.split(/[\\/]/).pop() ?? trimmed
  try { localStorage.setItem(LAST_FILE_KEY, trimmed) } catch { /* private window */ }
}

// The page only reads its tab and asks for in-place updates; opening other routes is out of scope here.
const navigation: PaneNavigation = {
  updateActiveTab: patch => { Object.assign(tab, patch) },
  openTab: partial => ({ ...partial, id: 'fastdocview-demo-other' }) as Tab,
  openOrUpdateActiveTab: patch => { Object.assign(tab, patch) },
  openBookTarget: () => {},
  navigateToDestination: () => {},
  switchTab: () => {},
  get activeTabId() { return tab.id },
  get activeTab() { return tab },
  get tabs() { return [tab] },
}
provide(PANE_NAVIGATION_KEY, navigation)

// Playwright drives the page through the same open path the buttons use.
;(window as unknown as { __fastDocViewDemo: { open: (path: string) => void } }).__fastDocViewDemo = { open }
</script>

<template>
  <div class="app-page-host">
    <form class="app-page-host-bar" dir="ltr" @submit.prevent="open(pathInput)">
      <input v-model="pathInput" type="search" placeholder="Document path" />
      <button type="submit">Open</button>
      <button v-for="sample in SAMPLE_FILES" :key="sample.path" type="button" @click="open(sample.path)">{{ sample.label }}</button>
    </form>
    <div class="app-page-host-diagnostics" dir="ltr" data-test="demo-diagnostics">
      <span v-if="demoDiagnostics.lastOpen">
        {{ demoDiagnostics.lastOpen.fileName }}: {{ demoDiagnostics.lastOpen.readerType }}, {{ demoDiagnostics.lastOpen.mode }}{{ demoDiagnostics.lastOpen.hasCachedIndex ? ' (cached index)' : '' }},
        server open {{ demoDiagnostics.lastOpen.openMilliseconds.toFixed(0) }} ms - {{ demoDiagnostics.lastOpen.reason }}
      </span>
      <button type="button" @click="readMemory(false)">Memory</button>
      <button type="button" @click="readMemory(true)">GC + memory</button>
      <span>{{ memory }}</span>
    </div>
    <div class="app-page-host-page">
      <FastDocViewPage :key="pageKey" />
    </div>
  </div>
</template>

<style scoped>
.app-page-host { display: flex; flex-direction: column; height: 100vh; background: var(--bg-primary); }
.app-page-host-bar { display: flex; gap: 8px; padding: 6px 10px; border-block-end: 1px solid var(--border-color); }
.app-page-host-bar input { flex: 1; min-width: 0; }
.app-page-host-diagnostics { display: flex; gap: 8px; align-items: center; padding: 2px 10px; font: 12px system-ui, sans-serif; color: var(--text-secondary); border-block-end: 1px solid var(--border-color); }
.app-page-host-page { flex: 1; min-height: 0; display: flex; flex-direction: column; }
</style>
