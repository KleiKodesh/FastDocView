// Live check: in-document search typed during a COLD first open (the index still building). Blocks searched before the facts
// their text depends on are indexed are reported provisional and searched again once the index completes (the same signal
// that refetches provisional blocks): the page's count must settle to exactly the service's count over the complete index,
// and Enter must step match by match. DOM + the page's DEV hook, no screenshots. ASCII output.
//   node playwright/provisionalSearchCheck.mjs <output folder> [<large .docx>] [paragraphs]
// Writes provisional-search-live-v1.mht there: one long div; every paragraph but the last gets ::before "Zebra " (N-2 shown) through
// :nth-last-of-type(n+2), and paragraph 2 ("Quokka") is hidden by :nth-last-of-type(N-1) - both need the div's end.
import { chromium } from 'playwright'
import fs from 'node:fs'
import path from 'node:path'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const folder = process.argv[2], docx = process.argv[3], paragraphs = Number(process.argv[4] ?? 400000)
if (!folder) { console.error('usage: node playwright/provisionalSearchCheck.mjs <output folder> [<large .docx>] [paragraphs]'); process.exit(2) }
const archive = path.resolve(folder, 'provisional-search-live-v1.mht')
const boundary = '----=_NextPart_01DB0000.00000004'
const lines = [`MIME-Version: 1.0`, `Content-Type: multipart/related; boundary="${boundary}"`, ``, `--${boundary}`,
  `Content-Location: file:///C:/provisional-search-live.htm`, `Content-Transfer-Encoding: 8bit`, `Content-Type: text/html; charset="us-ascii"`, ``,
  `<html>`, `<head>`, `<style>`, `<!--`,
  `div.long > p:nth-last-of-type(n+2)::before { content:"Zebra "; }`,
  `div.long > p:nth-last-of-type(${paragraphs - 1}) { display:none; }`,
  `-->`, `</style>`, `</head>`, `<body>`, `<div class=long>`]
for (let i = 1; i <= paragraphs; i++) lines.push(`<p>${i === 2 ? 'Quokka' : 'Paragraph'} ${i} of the long div</p>`)
lines.push(`</div>`, `</body>`, `</html>`, ``, `--${boundary}--`, ``)
fs.writeFileSync(archive, lines.join('\r\n'))

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
const marks = () => page.evaluate(() => {
  const result = {}
  for (const row of document.querySelectorAll('.windowed-scroller-viewport [data-index]')) {
    const all = [...row.querySelectorAll('mark.search-match')]
    if (all.length) result[row.dataset.index] = { count: all.length, current: all.findIndex(m => m.classList.contains('current')) }
  }
  return result
})
/** The service's own count over the whole (complete) document. */
const serviceTotal = async query => page.evaluate(async ({ id, q }) => {
  let total = 0, blocks = 0, from = 0
  while (true) {
    const batch = await (await fetch(`/api/documents/${id}/search?query=${encodeURIComponent(q)}&matchCase=false&fromBlock=${from}&maxBlocks=2000`)).json()
    total += batch.hitCounts.reduce((a, b) => a + b, 0)
    blocks += batch.hitBlocks.length
    if (batch.reachedEnd || batch.nextBlock < 0) return { total, blocks }
    from = batch.nextBlock
  }
}, { id: (await state()).documentId, q: query })

await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await waitFor(() => !!window.__fastDocViewDemo, null, 180_000)

/** Cold open, type the query at once, watch the count while the index builds, then let it settle. */
async function coldSearch(name, file, query, expectChange) {
  step(`${name}: cold first open of ${path.basename(file)}, search "${query.replace(/[^\x20-\x7e]/g, '?')}" while the index builds`)
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true, forceSidecar: true }), file)
  await waitFor(() => !!window.__fastDocViewTest?.state()?.documentId && document.querySelector('.word-document'), null, 120_000)
  const completeAtOpen = (await state()).isIndexComplete
  await page.click('.windowed-scroller-viewport')
  await page.keyboard.press('Control+f')
  await waitFor(() => !!document.querySelector('.search-bar input'))
  await page.fill('.search-bar input', '')
  await page.fill('.search-bar input', query)
  // while the index is still incomplete: did the search run, and the highest count it showed (test-side frame watch)
  const partial = await page.evaluate(() => new Promise(resolve => {
    let searched = false, count = 0
    const look = () => {
      const s = window.__fastDocViewTest.state()
      if (s.isIndexComplete) return resolve({ searched, count })
      searched ||= s.search.isSearching || s.search.matchCount > 0
      count = Math.max(count, s.search.matchCount)
      requestAnimationFrame(look)
    }
    look()
  }))
  console.log(`  index complete at open: ${completeAtOpen}; while building: searched ${partial.searched}, highest count ${partial.count}`)
  check(!completeAtOpen && partial.searched, `${name}: the search ran while the index was still being built`)
  await waitFor(() => window.__fastDocViewTest.state().isIndexComplete, null, 600_000)
  await waitFor(() => !window.__fastDocViewTest.state().search.isSearching, null, 600_000)
  const service = await serviceTotal(query)
  // the recount of provisional blocks runs after completion: wait (DOM poll, test side only) for the count to reach the service's
  let settled = true
  try { await waitFor(n => window.__fastDocViewTest.state().search.matchCount === n, service.total, 120_000) } catch { settled = false }
  const s = (await state()).search
  console.log(`  settled count ${s.matchCount}; service ${service.total} matches in ${service.blocks} blocks; label "${await label()}"`)
  check(settled && s.matchCount === service.total, `${name}: the count settles to exactly the service's count over the complete index (${s.matchCount} vs ${service.total})`)
  if (expectChange) check(partial.count !== service.total, `${name}: the count while building differed (the case under test: highest ${partial.count} -> ${service.total})`)
  if (service.total > 0) check((await label()) === `${s.currentMatch + 1} / ${s.matchCount}`, `${name}: the bar shows "current / total"`)
  return service.total
}

async function stepping(name, checksMarks = true) {
  await page.evaluate(() => window.__fastDocViewTest.scrollToBlock(0))
  await waitFor(() => window.__fastDocViewTest.state().firstVisible?.index === 0 || window.__fastDocViewTest.state().firstVisible === 0, null, 30_000).catch(() => {})
  const seen = []
  for (let i = 0; i < 5; i++) {
    const before = (await state()).search.currentMatch
    await page.keyboard.press('Enter')
    await waitFor(([b, first]) => { const s = window.__fastDocViewTest.state().search; return first || s.currentMatch !== b }, [before, i === 0], 10_000).catch(() => {})
    const s = (await state()).search
    await waitFor(block => !!document.querySelector(`.windowed-scroller-viewport [data-index="${block}"] mark.search-match.current`), s.current?.block, 10_000).catch(() => {})
    const row = (await marks())[s.current.block]
    seen.push({ index: s.currentMatch, block: s.current.block, occurrence: s.current.occurrence, currentMark: row?.current, label: await label(), total: s.matchCount })
  }
  for (const entry of seen) console.log(`  ${JSON.stringify(entry)}`)
  // generated ::before text is CSS, not DOM text: the highlighter has nothing to mark there, so marks are checked on the .docx
  if (checksMarks) check(seen.every(e => e.currentMark === e.occurrence), `${name}: the .current mark is the current match's occurrence in its block`)
  check(seen.slice(1).every((e, i) => e.index === seen[i].index + 1 || e.index === 0), `${name}: each Enter advances exactly one match`)
  check(seen.every(e => e.label === `${e.index + 1} / ${e.total}`), `${name}: the label follows every step`)
}

const zebra = await coldSearch('web archive ::before', archive, 'Zebra', true)
check(zebra === paragraphs - 2, `web archive: every shown paragraph but the last has "Zebra" (the hidden one not: ${zebra})`)
await stepping('web archive', false)
const quokka = await coldSearch('web archive display:none', archive, 'Quokka', false)   // whether block 1 is scanned before the build ends is timing
check(quokka === 0, 'web archive: the hidden paragraph is not found once the index is complete')
if (docx) {
  await coldSearch('large .docx', path.resolve(docx), process.env.PROVISIONAL_SEARCH_QUERY ?? '\u05e9\u05d1\u05ea', false)
  await stepping('large .docx')
}
check(errors.length === 0, `no console errors${errors.length ? ': ' + errors.slice(0, 3).join(' | ') : ''}`)
await browser.close()
fs.rmSync(archive, { force: true })
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
