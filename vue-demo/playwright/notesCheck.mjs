// Live check of the APP's notes pane (DOM + console + logged notes fetches; no screenshots).
//   node playwright/notesCheck.mjs <file-with-notes> [--headed]
// One windowed scroller over every note: the pane is full, fetches are bounded windows, sync on follows
// the text, sync off stays put while the text scrolls, a clicked reference lands its note (eased), a
// note's number jumps the text back, the note number is a superscript link, notes share the column's width.
import { chromium } from 'playwright'
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
if (!filePath) { console.error('usage: node playwright/notesCheck.mjs <file> [--headed]'); process.exit(2) }
const locales = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../../../KitveiHakodesh/vue-frontend/src/i18n/locales')
const syncTitles = ['en', 'he'].flatMap(locale => {
  const json = JSON.parse(readFileSync(path.join(locales, locale, 'fastDocView.json'), 'utf8'))
  return [json.syncNotesOnTitle, json.syncNotesOffTitle]
})

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch({ headless: !process.argv.includes('--headed') })
const page = await browser.newPage({ viewport: { width: 1300, height: 1000 } })
const consoleErrors = []
page.on('console', message => { if (message.type() === 'error') consoleErrors.push(message.text().slice(0, 200)) })
page.on('pageerror', error => consoleErrors.push(String(error).slice(0, 200)))
const noteFetches = []
page.on('request', request => {
  const match = /\/notes\?first=(\d+)&count=(\d+)/.exec(request.url())
  if (match) noteFetches.push([Number(match[1]), Number(match[2])])
})
const waitFor = (predicate, argument, timeout = 60_000) => page.waitForFunction(predicate, argument, { timeout })
const hook = (fn, arg) => page.evaluate(([body, a]) => new Function('t', 'a', body)(window.__fastDocViewTest, a), [fn, arg])
const NOTES = '.fast-doc-view-notes-pane .windowed-scroller-viewport'
const TEXT = '.fast-doc-view-scroller'
const notesState = () => page.evaluate(selector => {
  const viewport = document.querySelector(selector)
  if (!viewport) return null
  const top = viewport.getBoundingClientRect().top
  const rows = [...viewport.querySelectorAll('[data-note-index]')]
  const first = rows.find(row => row.getBoundingClientRect().bottom > top + 1)
  return {
    mounted: rows.length,
    firstIndex: first ? Number(first.dataset.noteIndex) : null,
    firstOffset: first ? Math.round(first.getBoundingClientRect().top - top) : null,
    full: viewport.scrollHeight > viewport.clientHeight,
    scrollTop: viewport.scrollTop,
    clientHeight: viewport.clientHeight,
  }
}, NOTES)

step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await waitFor(() => !!window.__fastDocViewDemo, null, 120_000)
await page.evaluate(p => window.__fastDocViewDemo.open(p), filePath)
await waitFor(() => { const s = window.__fastDocViewTest?.state(); return s?.documentId && s.blockCount > 0 && s.isIndexComplete }, null, 180_000)
const opened = await hook('return t.state()')
const totalNotes = opened.footnoteCount + opened.endnoteCount
console.log(`  blocks ${opened.blockCount}, notes ${totalNotes}`)

step('click the first reference in the text')
let reference = null
for (let block = 0; block < opened.blockCount && !reference; block += 30) {
  await hook('return t.scrollToBlock(a)', block)
  await page.waitForTimeout(150)
  reference = await page.$(`${TEXT} .note-reference[data-note-kind]`)
}
check(!!reference, 'found a note reference in the text')
await reference.click()
await waitFor(selector => document.querySelectorAll(`${selector} [data-note-index]`).length > 0, NOTES, 30_000)
await page.waitForTimeout(500)
let notes = await notesState()
console.log(`  pane: ${JSON.stringify(notes)}; fetches ${JSON.stringify(noteFetches)}`)
check(notes.full, 'the notes pane is full (content taller than the pane)')
check(noteFetches.every(([, count]) => count <= 100), 'every notes fetch is at most 100 notes')
check(noteFetches.reduce((sum, [, count]) => sum + count, 0) < totalNotes || totalNotes <= 60, 'not all notes fetched')

step('scroll the text down a lot with sync ON: the notes follow')
const syncButton = await page.$$eval('button[title]', (buttons, titles) => buttons.findIndex(button => titles.includes(button.title)), syncTitles)
check(syncButton >= 0, 'found the sync toggle')
const textTarget = Math.floor(opened.blockCount * 0.6)
await hook('return t.scrollToBlock(a)', textTarget)
await page.waitForTimeout(800)
const followed = await notesState()
const firstText = (await hook('return t.state()')).firstVisible?.index
const followedBlock = await page.evaluate(([selector, index]) => Number(document.querySelector(`${selector} [data-note-index="${index}"]`)?.dataset.blockIndex ?? -1), [NOTES, followed.firstIndex])
console.log(`  text first block ${firstText}; notes first ${followed.firstIndex} (block ${followedBlock}, offset ${followed.firstOffset})`)
check(followedBlock >= firstText && Math.abs(followed.firstOffset) <= 2, 'sync on: the first note at or after the first block in view is at the top')

step('sync OFF: the text scrolls, the notes stay')
await page.$$('button[title]').then(buttons => buttons[syncButton].click())
await page.waitForTimeout(200)
const before = await notesState()
await hook('return t.scrollToBlock(a)', Math.floor(opened.blockCount * 0.2))
await page.waitForTimeout(800)
const after = await notesState()
check(after.firstIndex === before.firstIndex && after.scrollTop === before.scrollTop, `sync off: notes stayed (${before.firstIndex} -> ${after.firstIndex})`)

step('sync OFF: the notes scroll on their own, across windows')
const fetchesBefore = noteFetches.length
for (let i = 0; i < 25; i++) { await page.mouse.move(650, 900); await page.mouse.wheel(0, 1500); await page.waitForTimeout(60) }
await page.waitForTimeout(600)
const scrolled = await notesState()
console.log(`  notes first ${before.firstIndex} -> ${scrolled.firstIndex}, mounted ${scrolled.mounted}; new fetches ${JSON.stringify(noteFetches.slice(fetchesBefore))}`)
check(scrolled.firstIndex > before.firstIndex, 'the notes pane scrolled by itself')

step('click a reference with sync off: its note lands (eased)')
// a reference far from where the notes pane is now: the block of a note mounted early in the pane earlier (note 3)
const farBlock = await page.evaluate(selector => Number(document.querySelector(`${selector} [data-note-index]`)?.dataset.blockIndex ?? 0), NOTES)
await page.mouse.move(650, 900)
for (let i = 0; i < 25; i++) { await page.mouse.wheel(0, -3000); await page.waitForTimeout(40) }
await page.waitForTimeout(400)
await hook('return t.scrollToBlock(a)', farBlock)
await page.waitForTimeout(500)
const clicked = await page.$(`${TEXT} .note-reference[data-note-kind]`)
const key = await clicked.evaluate(el => `${el.dataset.noteKind}|${el.dataset.noteNumber}`)
const tops = []
await clicked.click()
for (let i = 0; i < 25; i++) { const s = await notesState(); tops.push(s.scrollTop); await page.waitForTimeout(20) }
await page.waitForTimeout(300)
const landed = await notesState()
const landedRow = await page.evaluate(([selector, index]) => document.querySelector(`${selector} [data-note-index="${index}"]`)?.textContent?.slice(0, 40), [NOTES, landed.firstIndex])
const distinctTops = new Set(tops).size
console.log(`  reference ${key}; landed note ${landed.firstIndex} offset ${landed.firstOffset}; ${distinctTops} distinct scroll positions during the landing; "${landedRow}"`)
check(Math.abs(landed.firstOffset) <= 2, 'the clicked note is at the pane top')
check(distinctTops >= 3, 'the landing eased (several intermediate positions)')

step('the note number: superscript, ~0.7em, normal weight, link colour; click jumps the text back')
const mark = await page.$(`${NOTES} [data-note-index="${landed.firstIndex}"] :is(.note-mark, .note-label)`)
const markStyle = await mark.evaluate(el => {
  const style = getComputedStyle(el), parent = getComputedStyle(el.closest('.line'))
  return { va: style.verticalAlign, ratio: parseFloat(style.fontSize) / parseFloat(parent.fontSize), weight: style.fontWeight, color: style.color, parentColor: parent.color }
})
console.log(`  ${JSON.stringify(markStyle)}`)
check(markStyle.va === 'super' && Math.abs(markStyle.ratio - 0.7) < 0.08 && Number(markStyle.weight) <= 400 && markStyle.color !== markStyle.parentColor, 'note number looks like a footnote reference')
const noteBlock = await page.evaluate(([selector, index]) => Number(document.querySelector(`${selector} [data-note-index="${index}"]`).dataset.blockIndex), [NOTES, landed.firstIndex])
await hook('return t.scrollToBlock(a)', 0)
await page.waitForTimeout(400)
await mark.click()
await page.waitForTimeout(800)
const back = (await hook('return t.state()')).firstVisible?.index
check(Math.abs(back - noteBlock) <= 3, `note number jumped the text to its reference block (${noteBlock}, first visible ${back})`)

step('notes column = document column')
const widths = await page.evaluate(selector => {
  const note = document.querySelector(`${selector} .line`), block = document.querySelector('.fast-doc-view-block')
  const noteRect = note.getBoundingClientRect(), blockRect = block.getBoundingClientRect()
  return { note: [Math.round(noteRect.left), Math.round(noteRect.right)], block: [Math.round(blockRect.left), Math.round(blockRect.right)], notePad: getComputedStyle(note).paddingInlineStart, blockPad: getComputedStyle(block).paddingInlineStart }
}, NOTES)
console.log(`  ${JSON.stringify(widths)}`)
check(Math.abs(widths.note[0] - widths.block[0]) <= 12 && Math.abs(widths.note[1] - widths.block[1]) <= 12 && widths.notePad === widths.blockPad, 'notes share the document column edges (scrollbar width aside)')

console.log(`  all notes fetches: ${JSON.stringify(noteFetches)}`)
check(consoleErrors.length === 0, `no console errors ${consoleErrors.length ? JSON.stringify(consoleErrors.slice(0, 5)) : ''}`)
await browser.close()
console.log(failures ? `${failures} FAILED` : 'ALL PASS')
process.exit(failures ? 1 : 0)
