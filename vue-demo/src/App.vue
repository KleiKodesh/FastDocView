<script setup lang="ts">
import { ref } from 'vue'
import WordFilesPanel from './WordFilesPanel.vue'
import WordReaderPanel from './WordReaderPanel.vue'

const chosenPath = ref<string | null>(null)
const readerPanel = ref<InstanceType<typeof WordReaderPanel> | null>(null)

function onFileChosen(path: string) {
  chosenPath.value = path
  readerPanel.value?.open(path)
}

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
    <WordFilesPanel class="files" :chosen-path="chosenPath" @file-chosen="onFileChosen" />
    <WordReaderPanel ref="readerPanel" class="reader" />
  </div>
</template>

<style scoped>
.demo-layout { display: flex; height: 100%; }
.files { width: 360px; flex: none; border-inline-end: 1px solid #d5d5d5; background: #fafafa; }
.reader { flex: 1; min-width: 0; }
</style>
