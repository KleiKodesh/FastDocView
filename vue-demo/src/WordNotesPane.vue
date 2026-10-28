<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import type { WordNoteKind, WordVisibleNote } from './wordDemoApi'

// Word's draft (continuous) view keeps footnotes out of the text flow and shows them in a notes
// pane under it, with a footnotes / endnotes switch. This pane shows the notes of the blocks
// currently on screen, in reference order; a clicked reference highlights its note.

const props = defineProps<{
  notes: WordVisibleNote[]
  highlightedKey: string | null
  isRightToLeft: boolean
  footnoteCount: number
  endnoteCount: number
}>()

const shownKind = ref<WordNoteKind>('footnote')
const listElement = ref<HTMLElement | null>(null)

const shownNotes = computed(() => props.notes.filter(note => note.kind === shownKind.value))

// A clicked reference of the other kind switches the pane over, then the note is brought into view.
watch(() => props.highlightedKey, async key => {
  if (!key) return
  const note = props.notes.find(candidate => candidate.key === key)
  if (note && note.kind !== shownKind.value) shownKind.value = note.kind
  await nextTick()
  listElement.value?.querySelector(`[data-note-key="${CSS.escape(key)}"]`)?.scrollIntoView({ block: 'nearest' })
})
</script>

<template>
  <section class="notes-pane" data-test="notes-pane">
    <header class="notes-header">
      <select v-model="shownKind" data-test="notes-kind">
        <option value="footnote">All footnotes ({{ footnoteCount }})</option>
        <option value="endnote">All endnotes ({{ endnoteCount }})</option>
      </select>
      <span class="hint">{{ shownNotes.length ? `${shownNotes.length} on screen` : 'none on screen' }}</span>
    </header>
    <ol ref="listElement" class="notes-list word-document" :dir="isRightToLeft ? 'rtl' : 'ltr'">
      <li
        v-for="note in shownNotes"
        :key="note.key"
        :data-note-key="note.key"
        :class="{ highlighted: note.key === highlightedKey }"
        data-test="note"
        v-html="note.html"
      ></li>
    </ol>
  </section>
</template>

<style scoped>
.notes-pane { display: flex; flex-direction: column; min-height: 0; border-top: 1px solid #ccc; background: #fff; }
.notes-header { display: flex; align-items: center; gap: 10px; padding: 4px 10px; background: #f4f4f4; border-bottom: 1px solid #ddd; font-size: 12px; }
.hint { color: #777; }
.notes-list { list-style: none; margin: 0; padding: 4px 48px; overflow: auto; flex: 1; font-size: 0.9em; }
.notes-list li { padding: 2px 4px; border-radius: 3px; }
.notes-list li.highlighted { background: #fff3c4; }
</style>
