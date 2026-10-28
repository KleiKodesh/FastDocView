<script setup lang="ts">
import { onMounted, ref } from 'vue'
import { listFolder, type WordFolderEntry } from './wordDemoApi'

defineProps<{ chosenPath: string | null }>()
const emit = defineEmits<{ fileChosen: [path: string] }>()

const FOLDER_STORAGE_KEY = 'word-demo.folder'
const folderPath = ref('')
const entries = ref<WordFolderEntry[]>([])
const isLoading = ref(false)
const error = ref('')

async function scan() {
  error.value = ''
  isLoading.value = true
  try {
    entries.value = await listFolder(folderPath.value)
    try { localStorage.setItem(FOLDER_STORAGE_KEY, folderPath.value) } catch { /* private window */ }
  } catch (exception) {
    entries.value = []
    error.value = (exception as Error).message
  } finally {
    isLoading.value = false
  }
}

function formatBytes(bytes: number): string {
  return bytes >= 1 << 20 ? `${(bytes / (1 << 20)).toFixed(1)} MB` : `${Math.max(1, bytes >> 10)} KB`
}

onMounted(() => {
  try { folderPath.value = localStorage.getItem(FOLDER_STORAGE_KEY) ?? '' } catch { /* private window */ }
  if (folderPath.value) scan()
})
</script>

<template>
  <aside class="files-panel">
    <form class="folder-form" @submit.prevent="scan">
      <input v-model="folderPath" type="text" placeholder="Folder with .doc / .docx files" data-test="folder-input" />
      <button type="submit" :disabled="isLoading" data-test="scan-button">{{ isLoading ? 'Scanning…' : 'Scan' }}</button>
    </form>
    <p v-if="error" class="error">{{ error }}</p>
    <p v-else-if="entries.length" class="summary">
      {{ entries.length }} files — {{ entries.filter(e => e.mode === 'Sidecar').length }} worth an index
    </p>
    <ul class="entries">
      <li
        v-for="entry in entries"
        :key="entry.path"
        :class="{ chosen: entry.path === chosenPath }"
        :title="`${entry.path}\n${entry.reason}`"
        data-test="file-entry"
        @click="emit('fileChosen', entry.path)"
      >
        <span class="name" dir="rtl">{{ entry.fileName }}</span>
        <span class="meta">
          {{ formatBytes(entry.fileBytes) }} · body {{ formatBytes(entry.bodyBytes) }} ·
          <b :class="entry.mode.toLowerCase()">{{ entry.mode }}</b><span v-if="entry.hasCachedIndex"> · cached</span>
        </span>
      </li>
    </ul>
  </aside>
</template>

<style scoped>
.files-panel { display: flex; flex-direction: column; height: 100%; font-size: 13px; }
.folder-form { display: flex; gap: 6px; padding: 8px; }
.folder-form input { flex: 1; min-width: 0; padding: 4px 6px; }
.summary, .error { margin: 0; padding: 0 10px 6px; color: #666; }
.error { color: #b00020; }
.entries { list-style: none; margin: 0; padding: 0; overflow: auto; flex: 1; }
.entries li { padding: 6px 10px; cursor: pointer; border-bottom: 1px solid #eee; }
.entries li:hover { background: #eef2fb; }
.entries li.chosen { background: #dde6fb; }
.name { display: block; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.meta { display: block; color: #777; font-size: 11px; }
.meta b.sidecar { color: #1a6; }
.meta b.inmemory { color: #777; font-weight: normal; }
</style>
