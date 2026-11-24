// Live check of the APP's FastDocView page (the demo page IS the real page; only the transport is the
// demo's). DOM + console + the page's DEV test hook (window.__fastDocViewTest) - no screenshots.
//   node playwright/appCheck.mjs <file> [--headed]
// Cold first open (index rebuilt), count grows while the index builds, completes, exact jumps to the
// middle and the last block, footnote links both ways (when the file has notes), reload shows the
// same block at the same pixel, no console errors.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
const headed = process.argv.includes('--headed')
if (!filePath) { console.error('usage: node playwright/appCheck.mjs <file> [--headed]'); process.exit(2) }

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch({ headless: !headed })
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const consoleErrors = []
page.on('console', message => { if (message.type() === 'error') consoleErrors.push(message.text().slice(0, 200)) })
page.on('pageerror', error => consoleErrors.push(String(error).slice(0, 200)))
// The hook is registered on mount; a hot reload (the app's files being edited) remounts the page, so wait for it each time.
const hookReady = () => page.waitForFunction(() => !!window.__fastDocViewTest, null, { timeout: 60_000 })
const state = async () => { await hookReady(); return page.evaluate(() => window.__fastDocViewTest.state()) }
const scrollToBlock = async index => { await hookReady(); await page.evaluate(i => window.__fastDocViewTest.scrollToBlock(i), index) }
const waitFor = (predicate, argument, timeout = 60_000) => page.waitForFunction(predicate, argument, { timeout })
// a text row's top relative to the text viewport's top, sub-pixel (getBoundingClientRect); null when the row is not mounted
const rowTopInViewport = index => page.evaluate(i => {
  const row = [...document.querySelectorAll(`.windowed-scroller-viewport [data-index="${i}"]`)].find(element => !element.closest('.fast-doc-view-notes-pane'))
  const viewport = row?.closest('.windowed-scroller-viewport')
  return row && viewport ? row.getBoundingClientRect().top - viewport.getBoundingClientRect().top : null
}, index)

step(`open ${DEMO_URL}`)
await page.goto(`${DEMO_URL}`, { waitUntil: 'domcontentloaded' })
await waitFor(() => !!window.__fastDocViewDemo, null, 120_000)

step(`cold first open (index rebuilt) ${filePath}`)
const openStarted = Date.now()
await page.evaluate(path => window.__fastDocViewDemo.open(path, { rebuildIndex: true, forceSidecar: true }), filePath)
// content shown: a row with text, or with a drawing (a document of charts or pictures only has rows with no text at all - waiting for
// text there timed out although the page had drawn every chart)
await waitFor(() => [...document.querySelectorAll('.windowed-scroller-viewport [data-index]')].some(row => (row.textContent ?? '').trim().length > 0 || !!row.querySelector('svg, img')), null, 120_000)
const textShown = Date.now() - openStarted
const opened = await state()
console.log(`  text shown after ${textShown} ms; count ${opened?.blockCount}, complete ${opened?.isIndexComplete}`)
check(!!opened, 'the page exposes its DEV test hook')
check(textShown < 2000, 'text shown in under 2 s wall')
check(opened.blockCount === opened.knownBlockCount, 'the count is the known blocks, never an estimate')

if (!opened.isIndexComplete) {
  await page.waitForTimeout(1500)
  const building = await state()
  console.log(`  after 1.5 s: count ${building.blockCount}, complete ${building.isIndexComplete}, firstVisible ${building.firstVisible?.index}`)
  check(building.blockCount >= opened.blockCount, 'the count only grows while the index builds')
  check((building.firstVisible?.index ?? 0) <= 2, 'the reader stays at the top while the index builds')
}

step('the index completes')
await waitFor(() => window.__fastDocViewTest?.state().isIndexComplete, null, 180_000)
const complete = await state()
console.log(`  complete: count ${complete.blockCount}, headings ${complete.headingCount}, notes ${complete.footnoteCount}+${complete.endnoteCount}`)
check(complete.blockCount === complete.knownBlockCount, 'count exact once complete')

for (const target of [Math.floor(complete.blockCount / 2), complete.blockCount - 1]) {
  step(`scrollToBlock(${target})`)
  await scrollToBlock(target)
  await waitFor(index => !!document.querySelector(`.windowed-scroller-viewport [data-index="${index}"]`), target, 30_000).catch(() => {})
  await page.waitForTimeout(300)
  const landed = await state()
  const onScreen = await page.$eval(`.windowed-scroller-viewport [data-index="${target}"]`, row => {
    const view = row.closest('.windowed-scroller-viewport').getBoundingClientRect(), rect = row.getBoundingClientRect()
    return rect.bottom > view.top && rect.top < view.bottom
  }).catch(() => false)
  check(Math.abs((landed.firstVisible?.index ?? -1) - target) <= 2 || onScreen, `block ${target} is on screen (first visible ${landed.firstVisible?.index})`)
}

if (complete.footnoteCount + complete.endnoteCount > 0) {
  step('footnote links both ways')
  await scrollToBlock(0)
  // straight to the first note reference: the index lists every reference's block (no window-by-window scan, which took
  // ~90 s on the 87 MB sample whose first reference sits deep in the text)
  const references = await page.evaluate(async id => (await fetch(`/api/documents/${id}/note-references`)).json(), complete.documentId)
  console.log(`  ${references.length} note references; first in block ${references[0]?.blockIndex}`)
  let found = false
  // a note whose reference sits in a field code has no mark in the text by design (Word draws none): try the referencing blocks
  // in order until one draws a mark
  const blocks = [...new Set(references.map(reference => reference.blockIndex))]
  let markSelector = '.fast-doc-view-content .note-reference[data-note-kind]'
  for (let i = 0; i < blocks.length && !found; i++) {
    await scrollToBlock(blocks[i])
    const selector = `.windowed-scroller-viewport [data-index="${blocks[i]}"] .note-reference[data-note-kind]`
    found = await page.waitForSelector(selector, { state: 'attached', timeout: i === 0 ? 30_000 : 5_000 }).then(Boolean, () => false)
    if (found) markSelector = selector
  }
  check(found, 'a note reference is rendered in the text')
  if (found) {
    // What the page does (FastDocViewPage openNoteFromReference + FastDocViewNotesPane landNote): a click on a reference
    // opens the notes pane and LANDS its note at the pane's top (no highlight class); the note's label / own mark jumps
    // the text back to the reference and flashes it (.note-reference--current).
    const reference = await page.$(markSelector)
    const referenceLabel = ((await reference.textContent()) ?? '').trim()
    // a DOM click: a mark inside an EQ (MathML) may have no box Playwright counts as clickable
    await reference.evaluate(element => element.click())
    // wait for the page's own outcome - the pane open and the reference's note landed at its top - never a fixed delay:
    // (a reference clicked before the note list caught up once found no note - the page now waits for it, useFastDocViewNotes)
    await page.waitForFunction(label => {
      const viewport = document.querySelector('.fast-doc-view-notes-pane .notes-scroller')
      if (!window.__fastDocViewTest.state().notesPaneOpen || !viewport) return false
      const top = viewport.getBoundingClientRect().top
      return [...viewport.querySelectorAll('.fast-doc-view-note')].some(row => {
        const mark = row.querySelector('.note-label') ?? row.querySelector('.note-mark')
        const atEnd = viewport.scrollTop + viewport.clientHeight >= viewport.scrollHeight - 2
        const offset = row.getBoundingClientRect().top - top
        return (mark?.textContent ?? '').trim() === label && (Math.abs(offset) <= 12 || atEnd && offset >= -4 && offset < viewport.clientHeight)
      })
    }, referenceLabel, { timeout: 60000 }).catch(() => {})
    // the clicked reference's own note (not just the pane's top row: notes whose reference sits in a field code have no mark in
    // the text by design - Word draws none - so only a note with a mark is a valid jump target). A note near the end lands as high
    // as the pane can scroll: at the top, or anywhere visible with the pane scrolled to its end.
    const landed = await page.evaluate(label => {
      const viewport = document.querySelector('.fast-doc-view-notes-pane .notes-scroller')
      if (!viewport) return null
      const top = viewport.getBoundingClientRect().top
      const atEnd = viewport.scrollTop + viewport.clientHeight >= viewport.scrollHeight - 2
      const row = [...viewport.querySelectorAll('.fast-doc-view-note')].find(candidate => {
        const mark = candidate.querySelector('.note-label') ?? candidate.querySelector('.note-mark')
        return (mark?.textContent ?? '').trim() === label
      })
      if (!row) return null
      const offset = Math.round(row.getBoundingClientRect().top - top)
      return { index: Number(row.dataset.noteIndex), offset, label, atTop: Math.abs(offset) <= 12 || atEnd && offset >= -4 && offset < viewport.clientHeight }
    }, referenceLabel)
    console.log(`  reference "${referenceLabel}" -> note pane top: ${JSON.stringify(landed)}`)
    check((await state()).notesPaneOpen && !!landed, 'clicking a reference opens the notes pane')
    check(!!landed && landed.atTop,
          "the clicked reference's note lands at the notes pane's top")
    if (landed) {
      // sync off first: with it on, the notes pane rightly follows the text to the end, and a document with notes
      // throughout (synthetic RTF: one every ~12 blocks) then no longer draws the landed note
      await page.evaluate(() => window.__fastDocViewTest.setNotesSync(false))
      await scrollToBlock(Math.max(0, (await state()).blockCount - 1))
      await page.waitForTimeout(400)
      const back = await page.$(`.fast-doc-view-notes-pane .fast-doc-view-note[data-note-index="${landed.index}"] :is(.note-label, .note-mark)`)
      if (back) {
        await back.click()
        await page.waitForTimeout(600)
        check(!!(await page.$('.note-reference--current')), 'clicking the note label jumps back and flashes its reference')
      } else check(false, 'the landed note keeps its label after the text scrolled away')
      await page.evaluate(() => window.__fastDocViewTest.setNotesSync(true))
    }
  }
}

step('reload restores the position')
const middle = Math.floor(complete.blockCount / 3)
await scrollToBlock(middle)
await page.waitForTimeout(800)                        // the page saves its position debounced
const before = await state()
const beforeTop = await rowTopInViewport(before.firstVisible?.index)
await page.reload({ waitUntil: 'domcontentloaded' })
await waitFor(() => window.__fastDocViewTest?.state()?.isIndexComplete && window.__fastDocViewTest.state().firstVisible, null, 120_000)
await page.waitForTimeout(1500)
const after = await state()
const afterTop = await rowTopInViewport(before.firstVisible?.index)
console.log(`  before: block ${before.firstVisible?.index} @ ${before.firstVisible?.offsetWithinRow?.toFixed(1)} px   after: block ${after.firstVisible?.index} @ ${after.firstVisible?.offsetWithinRow?.toFixed(1)} px`)
console.log(`  block ${before.firstVisible?.index}'s top in the viewport: before ${beforeTop?.toFixed(2)} px, after ${afterTop?.toFixed(2)} px`)
// compared in pixels, not by "first visible block": rows have fractional heights and readFirstVisible works on the rounded
// offsetTop / offsetHeight, so a row edge within a pixel of the viewport's top reads as either neighbour
check(beforeTop != null && afterTop != null && Math.abs(afterTop - beforeTop) <= 1, 'the same block at the same pixel after reload (1 px)')

check(consoleErrors.length === 0, `no console errors${consoleErrors.length ? ': ' + consoleErrors.slice(0, 3).join(' | ') : ''}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
