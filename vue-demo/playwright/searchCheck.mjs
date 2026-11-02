// Live check of the APP's FastDocView in-document search through the demo: the count is MATCHES
// (every occurrence, not one per block), Enter / Shift+Enter step match by match inside a block
// and then on to the next block, the current match carries the `.current` mark and the others the
// plain mark, the bar shows "n / total". DOM + console + the page's DEV hook - no screenshots.
//   node playwright/searchCheck.mjs <file> [<query>]
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
// default: a frequent Hebrew word (written as escapes: the tooling's traffic stays ASCII)
const query = process.argv[3] ?? 'שבת'
if (!filePath) { console.error('usage: node playwright/searchCheck.mjs <file> [<query>]'); process.exit(2) }

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const consoleErrors = []
page.on('console', message => { if (message.type() === 'error') consoleErrors.push(message.text().slice(0, 200)) })
page.on('pageerror', error => consoleErrors.push(String(error).slice(0, 200)))
const waitFor = (predicate, argument, timeout = 60_000) => page.waitForFunction(predicate, argument, { timeout })
const state = () => page.evaluate(() => window.__fastDocViewTest.state())
/** The rendered rows' marks: per block index, the number of marks and which one (if any) is current. */
const marks = () => page.evaluate(() => {
  const rows = [...document.querySelectorAll('.windowed-scroller-viewport [data-index]')]
  const result = {}
  for (const row of rows) {
    const all = [...row.querySelectorAll('mark.search-match')]
    if (all.length) result[row.dataset.index] = { count: all.length, current: all.findIndex(m => m.classList.contains('current')) }
  }
  return result
})
const label = () => page.$eval('.search-bar .match-count', el => el.textContent.trim())

step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await waitFor(() => !!window.__fastDocViewDemo, null, 120_000)
step(`open ${filePath}`)
await page.evaluate(path => window.__fastDocViewDemo.open(path), filePath)
await waitFor(() => window.__fastDocViewTest?.state().isIndexComplete, null, 180_000)
await page.evaluate(() => window.__fastDocViewTest.scrollToBlock(0))
await page.waitForTimeout(400)

step('Ctrl+F, type the query, wait for the scan')
await page.click('.windowed-scroller-viewport')
await page.keyboard.press('Control+f')
await waitFor(() => !!document.querySelector('.search-bar input'))
await page.fill('.search-bar input', query)
await waitFor(() => { const s = window.__fastDocViewTest.state().search; return s.matchCount > 0 && !s.isSearching }, null, 120_000)
let s = (await state()).search
console.log(`  matches ${s.matchCount}, current ${s.currentMatch} -> ${JSON.stringify(s.current)}, label "${await label()}"`)
check(s.matchCount > 0, 'matches found')
check((await label()) === `${s.currentMatch + 1} / ${s.matchCount}`, 'the bar shows "current / total" over matches')

// the service's own count of matches vs blocks, for the record (the dev server speaks JSON)
const documentId = (await state()).documentId
let blocks = 0, matches = 0, from = 0
while (true) {
  const batch = await page.evaluate(async ({ id, q, f }) => (await fetch(`/api/documents/${id}/search?query=${encodeURIComponent(q)}&matchCase=false&fromBlock=${f}&maxBlocks=2000`)).json(), { id: documentId, q: query, f: from })
  blocks += batch.hitBlocks.length
  matches += batch.hitCounts.reduce((a, b) => a + b, 0)
  if (batch.reachedEnd || batch.nextBlock < 0) break
  from = batch.nextBlock
}
console.log(`  service: ${blocks} hit blocks, ${matches} matches`)
check(matches > blocks, 'the sample has blocks with several matches (the case under test)')
check(s.matchCount === matches, `the page counts matches (${s.matchCount}), not blocks (${blocks})`)

step('rendered marks per block equal the service counts')
const first = await marks()
const visibleRows = Object.keys(first).map(Number)
let markTotal = 0
for (const block of visibleRows) markTotal += first[block].count
console.log(`  ${visibleRows.length} rendered rows with marks, ${markTotal} marks`)
const perBlock = await page.evaluate(async ({ id, q, rows }) => {
  const result = {}
  for (const row of rows) {
    const batch = await (await fetch(`/api/documents/${id}/search?query=${encodeURIComponent(q)}&matchCase=false&fromBlock=${row}&maxBlocks=1`)).json()
    result[row] = batch.hitBlocks.length ? batch.hitCounts[0] : 0
  }
  return result
}, { id: documentId, q: query, rows: visibleRows })
check(visibleRows.every(block => perBlock[block] === first[block].count), `the highlighter's marks per block equal the service's counts (${visibleRows.map(b => `${b}:${first[b].count}/${perBlock[b]}`).join(' ')})`)

step('Enter steps match by match within a block, then to the next block')
const seen = []
for (let i = 0; i < 6; i++) {
  await page.keyboard.press('Enter')
  await page.waitForTimeout(350)
  s = (await state()).search
  const m = await marks()
  const row = m[s.current.block]
  seen.push({ index: s.currentMatch, block: s.current.block, occurrence: s.current.occurrence, marksInRow: row?.count, currentMark: row?.current, label: await label() })
}
for (const entry of seen) console.log(`  ${JSON.stringify(entry)}`)
check(seen.every(e => e.currentMark === e.occurrence), 'the .current mark is the nth mark of the current match\'s block')
check(seen.every(e => e.label === `${e.index + 1} / ${s.matchCount}`), 'the label follows every step')
const consecutive = seen.slice(1).every((e, i) => e.index === seen[i].index + 1 || e.index === 0)
check(consecutive, 'each Enter advances exactly one match')
const withinBlock = seen.slice(1).some((e, i) => e.block === seen[i].block && e.occurrence === seen[i].occurrence + 1)
const acrossBlocks = seen.slice(1).some((e, i) => e.block > seen[i].block && e.occurrence === 0)
check(withinBlock, 'a step moved to the next occurrence inside the same block')
check(acrossBlocks, 'a step moved on to the first occurrence of a later block')
check(await page.evaluate(() => document.querySelectorAll('.windowed-scroller-viewport mark.search-match.current').length) === 1, 'exactly one current mark in the view')

step('Shift+Enter steps back one match')
const before = (await state()).search.currentMatch
await page.keyboard.press('Shift+Enter')
await page.waitForTimeout(350)
s = (await state()).search
check(s.currentMatch === before - 1, `Shift+Enter goes back one match (${before} -> ${s.currentMatch})`)
const m = await marks()
check(m[s.current.block]?.current === s.current.occurrence, 'the .current mark followed the step back')

check(consoleErrors.length === 0, `no console errors${consoleErrors.length ? ': ' + consoleErrors.slice(0, 3).join(' | ') : ''}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
