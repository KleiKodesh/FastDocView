<script setup lang="ts">
import { onMounted, ref } from 'vue'
import WordFilesPanel from './WordFilesPanel.vue'
import WordReaderPanel from './WordReaderPanel.vue'
import { DEFAULT_FILE_PATH, LAST_FILE_STORAGE_KEY, SAMPLE_FILES } from './wordDemoDefaults'

const chosenPath = ref<string | null>(null)
const readerPanel = ref<InstanceType<typeof WordReaderPanel> | null>(null)

function onFileChosen(path: string) {
  chosenPath.value = path
  try { localStorage.setItem(LAST_FILE_STORAGE_KEY, path) } catch { /* private window */ }
  readerPanel.value?.open(path)
}

// Open the last file (or the default big one) straight away, so a page load is the test.
// ?file=<path> overrides both; ?file= (empty) opens nothing.
onMounted(() => {
  const fromQuery = new URLSearchParams(location.search).get('file')
  let lastFile: string | null = null
  try { lastFile = localStorage.getItem(LAST_FILE_STORAGE_KEY) } catch { /* private window */ }
  const path = fromQuery ?? lastFile ?? DEFAULT_FILE_PATH
  if (path) { chosenPath.value = path; readerPanel.value?.open(path) }
})

// Hooks for live debugging from Playwright / the console, like the app's __khTest.
;(window as unknown as { __wordDemo: unknown }).__wordDemo = {
  open: (path: string, options?: { forceSidecar?: boolean; rebuildIndex?: boolean }) => readerPanel.value?.open(path, options),
  reload: () => readerPanel.value?.reload(),
  scrollToBlock: (index: number) => readerPanel.value?.scrollToBlock(index),
  state: () => readerPanel.value?.debugState(),
}
</script>

<template>
  <div class="demo-layout">
    <div class="side">
      <div class="samples">
        <button
          v-for="sample in SAMPLE_FILES"
          :key="sample.path"
          :class="{ chosen: sample.path === chosenPath }"
          :title="sample.path"
          data-test="sample-button"
          @click="onFileChosen(sample.path)"
        >{{ sample.label }}</button>
      </div>
      <WordFilesPanel class="files" :chosen-path="chosenPath" @file-chosen="onFileChosen" />
    </div>
    <WordReaderPanel ref="readerPanel" class="reader" />
  </div>
</template>

<style scoped>
.demo-layout { display: flex; height: 100%; }
.side { width: 360px; flex: none; display: flex; flex-direction: column; border-inline-end: 1px solid #d5d5d5; background: #fafafa; }
.samples { display: flex; flex-direction: column; gap: 4px; padding: 8px; border-bottom: 1px solid #ddd; }
.samples button { text-align: start; padding: 4px 8px; cursor: pointer; }
.samples button.chosen { background: #dde6fb; }
.files { flex: 1; min-height: 0; }
.reader { flex: 1; min-width: 0; }
</style>
