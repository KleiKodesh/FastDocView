<script setup lang="ts">
// Hosts the app's FastDocViewPage the way the app's AppShell does: the tab store's active tab carries the
// document (route /fast-doc-view, localFilePath) - a REAL tab, so the page's saved position and zoom persist and
// come back after a reload, exactly as in the app (the store saves view state only for open tabs). The page is
// keyed by tab + file (+ an open counter, so re-opening the same file remounts it, as a cold open).
import { computed, provide, ref } from 'vue'
import FastDocViewPage from '@/features/fast-doc-view/FastDocViewPage.vue'
import { PANE_NAVIGATION_KEY, type PaneNavigation } from '@/composables/usePaneNavigation'
import { useTabStore, type Tab } from '@/stores/tabStore'
import { SAMPLE_FILES } from '../wordDemoDefaults'
import { demoDiagnostics, demoOpenOptions } from './fastDocViewDemoApi'

const tabStore = useTabStore()
const tab = computed<Tab>(() => tabStore.activeTab)
const queryFile = new URLSearchParams(location.search).get('file')
// first load: the restored tab's document, else ?file=, else the first sample
if (queryFile || tab.value.route !== '/fast-doc-view' || !tab.value.localFilePath)
  tabStore.updateActiveTab({ route: '/fast-doc-view', localFilePath: queryFile || SAMPLE_FILES[0]!.path, title: '' })

const pathInput = ref(tab.value.localFilePath ?? '')
const openCount = ref(0)
const pageKey = computed(() => `${tab.value.id}:${tab.value.localFilePath}:${openCount.value}`)

// Server memory on demand (no timers in the page): heap / process / open documents, after an optional GC.
const memory = ref('')
async function readMemory(collect: boolean) {
  try {
    const m = await (await fetch(`/api/memory?collect=${collect}`)).json() as { managedHeapMegabytes: number; workingSetMegabytes: number; openDocuments: number }
    memory.value = `heap ${m.managedHeapMegabytes} MB, process ${m.workingSetMegabytes} MB, ${m.openDocuments} open`
  } catch { memory.value = 'server unreachable' }
}

function open(path: string, options: { rebuildIndex?: boolean; forceSidecar?: boolean } = {}) {
  const trimmed = path.trim()
  if (!trimmed) return
  Object.assign(demoOpenOptions, options)
  pathInput.value = trimmed
  openCount.value++
  tabStore.updateActiveTab({ route: '/fast-doc-view', localFilePath: trimmed, title: trimmed.split(/[\/]/).pop() ?? trimmed })
}

// The app's pane navigation over the tab store (one pane); opening other routes is out of scope here.
const navigation: PaneNavigation = {
  updateActiveTab: patch => tabStore.updateActiveTab(patch),
  openTab: partial => tabStore.openTab(partial),
  openOrUpdateActiveTab: patch => tabStore.updateActiveTab(patch),
  openBookTarget: () => {},
  navigateToDestination: () => {},
  switchTab: id => tabStore.switchTab(id),
  get activeTabId() { return tab.value.id },
  get activeTab() { return tab.value },
  get tabs() { return [tab.value] },
}
provide(PANE_NAVIGATION_KEY, navigation)

// Playwright drives the page through the same open path the buttons use.
;(window as unknown as { __fastDocViewDemo: { open: typeof open } }).__fastDocViewDemo = { open }
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
