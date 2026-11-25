// Live check: the NOTES search typed during a COLD first open (the index still building). Notes whose text depends on what
// only the final index knows are reported provisional and searched again once the index completes (the same signal that
// refetches provisional blocks and notes): the notes search count must settle to exactly the service's notes count over the
// complete index, and Enter must step note match by note match, landing the notes pane on each. DOM + the page's DEV hook,
// no screenshots. ASCII output.
//   node playwright/provisionalNotesSearchCheck.mjs <large .docx> [query]
import { chromium } from 'playwright'
import path from 'node:path'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const file = process.argv[2], query = process.argv[3] ?? 'שבת'
if (!file) { console.error('usage: node playwright/provisionalNotesSearchCheck.mjs <large .docx> [query]'); process.exit(2) }

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
const waitFor = (predicate, argument, timeout = 60_000) => page.waitForFunction(predicate, argument, { timeout, polling: 100 })
const state = () => page.evaluate(() => window.__fastDocViewTest.state())
const label = () => page.$eval('.search-bar .match-count', el => el.textContent.trim()).catch(() => '')
/** The service's own notes count over the whole (complete) document. */
const serviceTotal = async q => page.evaluate(async ({ id, q }) => {
  let total = 0, notes = 0, from = 0
  while (true) {
    const batch = await (await fetch(`/api/documents/${id}/search-notes?query=${encodeURIComponent(q)}&matchCase=false&fromNote=${from}&maxNotes=2000`)).json()
    total += batch.hitCounts.reduce((a, b) => a + b, 0)
    notes += batch.hitBlocks.length
    if (batch.reachedEnd || batch.nextBlock < 0) return { total, notes }
    from = batch.nextBlock
  }
}, { id: (await state()).documentId, q })

await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await waitFor(() => !!window.__fastDocViewDemo, null, 180_000)
step(`cold first open of ${path.basename(file)}, notes search while the index builds`)
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true, forceSidecar: true }), path.resolve(file))
await waitFor(() => !!window.__fastDocViewTest?.state()?.documentId && document.querySelector('.word-document'), null, 120_000)
const completeAtOpen = (await state()).isIndexComplete
await page.click('.windowed-scroller-viewport')
await waitFor(() => window.__fastDocViewTest.state().footnoteCount + window.__fastDocViewTest.state().endnoteCount > 0, null, 60_000)
await page.evaluate(() => window.__fastDocViewTest.setNotesPlace('bottom'))
await waitFor(() => window.__fastDocViewTest.state().notesPaneOpen, null, 30_000)
await page.keyboard.press('Control+f')
await waitFor(() => !!document.querySelector('.search-bar input'))
await page.evaluate(() => window.__fastDocViewTest.searchNotes(true))
await waitFor(() => window.__fastDocViewTest.state().notesSearch.active)
await page.fill('.search-bar input', query)
const partial = await page.evaluate(() => new Promise(resolve => {
  let searched = false, count = 0
  const look = () => {
    const s = window.__fastDocViewTest.state()
    if (s.isIndexComplete) return resolve({ searched, count })
    searched ||= s.notesSearch.isSearching || s.notesSearch.matchCount > 0
    count = Math.max(count, s.notesSearch.matchCount)
    requestAnimationFrame(look)
  }
  look()
}))
console.log(`  index complete at open: ${completeAtOpen}; while building: searched ${partial.searched}, highest count ${partial.count}`)
check(!completeAtOpen && partial.searched, 'the notes search ran while the index was still being built')
await waitFor(() => window.__fastDocViewTest.state().isIndexComplete, null, 600_000)
await waitFor(() => !window.__fastDocViewTest.state().notesSearch.isSearching, null, 600_000)
const service = await serviceTotal(query)
let settled = true
try { await waitFor(n => window.__fastDocViewTest.state().notesSearch.matchCount === n, service.total, 120_000) } catch { settled = false }
const s = (await state()).notesSearch
console.log(`  settled count ${s.matchCount}; service ${service.total} matches in ${service.notes} notes; label "${await label()}"`)
check(settled && s.matchCount === service.total && service.total > 0, `the notes count settles to exactly the service's count over the complete index (${s.matchCount} vs ${service.total})`)
check((await label()) === `${s.currentMatch + 1} / ${s.matchCount}`, 'the bar shows "current / total"')

const seen = []
for (let i = 0; i < 6; i++) {
  const before = (await state()).notesSearch.currentMatch
  await page.keyboard.press('Enter')
  await waitFor(([b, first]) => { const n = window.__fastDocViewTest.state().notesSearch; return first || n.currentMatch !== b }, [before, i === 0], 10_000).catch(() => {})
  const n = (await state()).notesSearch
  const shown = await waitFor(note => !!document.querySelector(`.fast-doc-view-notes-pane [data-index="${note}"] mark`), n.current?.note ?? n.current?.block, 10_000).then(() => true, () => false)
  seen.push({ index: n.currentMatch, note: n.current?.block, occurrence: n.current?.occurrence, shown, label: await label(), total: n.matchCount })
}
for (const entry of seen) console.log(`  ${JSON.stringify(entry)}`)
check(seen.slice(1).every((e, i) => e.index === seen[i].index + 1 || e.index === 0), 'each Enter advances exactly one notes match')
check(seen.every(e => e.shown), 'the notes pane shows the current note with its marks')
check(seen.every(e => e.label === `${e.index + 1} / ${e.total}`), 'the label follows every step')
check(errors.length === 0, `no console errors${errors.length ? ': ' + errors.slice(0, 3).join(' | ') : ''}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
