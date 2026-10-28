// Live check against the running demo (npm run dev first, or pass --start to spawn it):
//   node playwright/liveCheck.mjs "<path to .doc/.docx>" [--headed]
// Opens the file through the page's __wordDemo hooks, then: waits for rows, jumps to the middle,
// the end, a heading, reads the scroller state, counts rendered rows, reports server memory and
// every console error. Exit code 1 on any failure. No screenshots — DOM and console only.
import { chromium } from 'playwright'

const DEMO_URL = 'http://localhost:5189/'
const filePath = process.argv[2]
const headed = process.argv.includes('--headed')
if (!filePath) { console.error('usage: node playwright/liveCheck.mjs <file> [--headed]'); process.exit(2) }

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = message => console.log(`[${new Date().toLocaleTimeString()}] ${message}`)

const browser = await chromium.launch({ headless: !headed })
const page = await browser.newPage({ viewport: { width: 1300, height: 850 } })
const consoleErrors = []
page.on('console', message => { if (message.type() === 'error') consoleErrors.push(message.text()) })
page.on('pageerror', error => consoleErrors.push(`pageerror: ${error.message}`))

step(`open ${DEMO_URL}`)
await page.goto(`${DEMO_URL}?file=`, { waitUntil: 'networkidle' })   // empty ?file= — the check opens its own file
await page.evaluate(() => fetch('/api/documents', { method: 'DELETE' }))   // leftovers from an earlier run

step(`open document ${filePath}`)
const openStarted = Date.now()
await page.evaluate(path => window.__wordDemo.open(path), filePath)
await page.waitForSelector('[data-test="scroller"]', { timeout: 180_000 })
await page.waitForSelector('[data-block]', { timeout: 60_000 })
const state0 = await page.evaluate(() => window.__wordDemo.state())
console.log(`  opened in ${Date.now() - openStarted} ms wall, server open ${state0.document.openMilliseconds.toFixed(0)} ms, ${state0.document.blockCount.toLocaleString()} blocks, ${state0.document.mode}`)
check(state0.document.blockCount > 0, 'document has blocks')

async function renderedRows() {
  return page.$$eval('[data-block]', rows => rows.map(r => ({ index: Number(r.dataset.block), text: r.textContent.trim().length, height: r.getBoundingClientRect().height })))
}
async function waitForRowsAround(index) {
  await page.waitForFunction(i => document.querySelector(`[data-block="${i}"]`)?.textContent !== undefined, index, { timeout: 30_000 })
}

const rows0 = await renderedRows()
console.log(`  rendered rows at open: ${rows0.length} (${rows0[0]?.index}–${rows0.at(-1)?.index}), with text: ${rows0.filter(r => r.text > 0).length}`)
check(rows0.length > 0 && rows0.length < 400, 'a bounded window is rendered, not the whole document')
check(rows0.some(r => r.text > 0), 'rendered rows have text')

for (const fraction of [0.5, 0.98, 0.1]) {
  const target = Math.floor(state0.document.blockCount * fraction)
  step(`scrollToBlock(${target})`)
  const started = Date.now()
  await page.evaluate(i => window.__wordDemo.scrollToBlock(i), target)
  await waitForRowsAround(target)
  const state = await page.evaluate(() => window.__wordDemo.state())
  const rows = await renderedRows()
  console.log(`  landed in ${Date.now() - started} ms: firstVisible ${state.firstVisible?.index}, rendered ${rows.length} rows (${rows[0]?.index}–${rows.at(-1)?.index}), cached ${state.cachedBlocks}, ${state.lastFetch}`)
  check(Math.abs((state.firstVisible?.index ?? -1) - target) <= 2, `first visible row is ${target} (±2)`)
  check(rows.length < 400, 'window stays bounded after a jump')
  check(state.cachedBlocks <= 3200, `block cache bounded (${state.cachedBlocks})`)
}

const headingCount = await page.$$eval('[data-test="heading"]', items => items.length)
if (headingCount > 1) {
  step('click the second heading')
  const headingBlock = await page.$$eval('[data-test="heading"]', items => Number(items[1].title.replace('block ', '')))
  await page.locator('[data-test="heading"]').nth(1).click()
  await waitForRowsAround(headingBlock)
  const state = await page.evaluate(() => window.__wordDemo.state())
  check(Math.abs((state.firstVisible?.index ?? -1) - headingBlock) <= 2, `heading click lands on block ${headingBlock}`)
}

step('wheel-scroll 30 notches')
await page.locator('[data-test="scroller"]').hover()
for (let i = 0; i < 30; i++) await page.mouse.wheel(0, 600)
await page.waitForTimeout(800)
const afterWheel = await page.evaluate(() => window.__wordDemo.state())
console.log(`  after wheel: firstVisible ${afterWheel.firstVisible?.index}, rendered ${afterWheel.renderedRange?.start}–${afterWheel.renderedRange?.end}`)

step('reload — must come back to the same block at the same pixel (the lines view contract)')
const beforeReload = await page.evaluate(() => window.__wordDemo.state())
const anchorText = await page.evaluate(i => document.querySelector(`[data-block="${i}"]`)?.textContent?.trim().slice(0, 40), beforeReload.firstVisible.index)
const reloadStarted = Date.now()
await page.locator('[data-test="reload-button"]').click()
await page.waitForFunction(() => !document.querySelector('[data-test="status"]')?.textContent?.startsWith('Opening'), null, { timeout: 180_000 })
await page.waitForSelector('[data-block]', { timeout: 60_000 })
await page.waitForTimeout(300)
const reloaded = await page.evaluate(() => window.__wordDemo.state())
const reloadedAnchorText = await page.evaluate(i => document.querySelector(`[data-block="${i}"]`)?.textContent?.trim().slice(0, 40), beforeReload.firstVisible.index)
console.log(`  reload: ${Date.now() - reloadStarted} ms wall, server ${reloaded.document.openMilliseconds.toFixed(0)} ms, mode ${reloaded.document.mode}`)
console.log(`  before: block ${beforeReload.firstVisible.index} @ ${beforeReload.firstVisible.offsetWithinRow.toFixed(1)} px  →  after: block ${reloaded.firstVisible?.index} @ ${reloaded.firstVisible?.offsetWithinRow.toFixed(1)} px`)
check(reloaded.document.openMilliseconds < 1500, 'server reopen under 1.5 s')
check(reloaded.firstVisible?.index === beforeReload.firstVisible.index, 'same first visible block after reload')
check(Math.abs((reloaded.firstVisible?.offsetWithinRow ?? 1e9) - beforeReload.firstVisible.offsetWithinRow) <= 2, 'same offset within the block (±2 px)')
check(anchorText === reloadedAnchorText, `anchor row text unchanged: "${anchorText}"`)

const memory = await page.evaluate(() => fetch('/api/memory?collect=true').then(r => r.json()))
console.log(`  server after GC: heap ${memory.managedHeapMegabytes} MB, process ${memory.workingSetMegabytes} MB, open documents ${memory.openDocuments}`)
check(memory.openDocuments === 1, 'exactly one document open on the server after reload')

const status = await page.locator('[data-test="status"]').textContent()
console.log(`  status: ${status}`)
check(consoleErrors.length === 0, consoleErrors.length ? `console errors:\n    ${consoleErrors.join('\n    ')}` : 'no console errors')

await browser.close()
console.log(failures ? `\n${failures} FAILURE(S)` : '\nALL PASSED')
process.exit(failures ? 1 : 0)
