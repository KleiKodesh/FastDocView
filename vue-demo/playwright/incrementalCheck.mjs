// Live check of the on-demand index (first open of a big .docx with no sidecar), against the running demo:
//   node playwright/incrementalCheck.mjs "<path to big .docx>" [--headed]
// Opens with rebuildIndex + forceSidecar (sidecar deleted), then: text shows before the index is complete,
// the scroller count is the server's estimate, a jump to the middle corrects the count from the blocks
// response, a jump to the end completes the index, after which headings + weights are refetched and the
// count is exact. No timers in the page: every update rides on a blocks response. DOM + console only.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? 'http://localhost:5189/'
const filePath = process.argv[2]
const headed = process.argv.includes('--headed')
if (!filePath) { console.error('usage: node playwright/incrementalCheck.mjs <file.docx> [--headed]'); process.exit(2) }

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = message => console.log(`[${new Date().toLocaleTimeString()}] ${message}`)

const browser = await chromium.launch({ headless: !headed })
const page = await browser.newPage({ viewport: { width: 1300, height: 850 } })
const consoleErrors = []
page.on('console', message => { if (message.type() === 'error') consoleErrors.push(message.text()) })
page.on('pageerror', error => consoleErrors.push(`pageerror: ${error.message}`))
const state = () => page.evaluate(() => window.__wordDemo.state())

step(`open ${DEMO_URL}`)
await page.goto(`${DEMO_URL}?file=`, { waitUntil: 'networkidle' })
await page.evaluate(() => fetch('/api/documents', { method: 'DELETE' }))

step(`first open (sidecar deleted) ${filePath}`)
const openStarted = Date.now()
await page.evaluate(path => window.__wordDemo.open(path, { rebuildIndex: true, forceSidecar: true }), filePath)
await page.waitForFunction(() => [...document.querySelectorAll('[data-block]')].some(row => row.textContent.trim().length > 0), null, { timeout: 60_000 })
const textShownMilliseconds = Date.now() - openStarted
const opened = await state()
console.log(`  text shown after ${textShownMilliseconds} ms wall; server open ${opened.document.openMilliseconds.toFixed(0)} ms; ` +
            `count ${opened.document.blockCount.toLocaleString()} (known ${opened.document.knownBlockCount.toLocaleString()}), complete ${opened.document.isIndexComplete}`)
check(!opened.document.isIndexComplete, 'index is still being built when text shows')
check(opened.document.knownBlockCount === opened.document.blockCount, 'the count is the known blocks, never an estimate')
check(textShownMilliseconds < 2000, 'text shown in under 2 s wall')
check(await page.isVisible('[data-test="index-state"]'), 'toolbar shows the on-demand index state')

step('while reading the start, the index keeps building and the count grows')
await page.waitForTimeout(1500)
const building = await state()
console.log(`  after 1.5 s: count ${building.document.blockCount.toLocaleString()}, complete ${building.document.isIndexComplete}, firstVisible ${building.firstVisible?.index}`)
check(building.document.blockCount > opened.document.blockCount, `the count grew (${opened.document.blockCount.toLocaleString()} -> ${building.document.blockCount.toLocaleString()})`)
check((building.firstVisible?.index ?? -1) <= 2, 'the reader stays where it was (top)')

step('the index completes on its own')
let started = Date.now()
await page.waitForFunction(() => window.__wordDemo.state().document.isIndexComplete, null, { timeout: 180_000 })
await page.waitForFunction(() => window.__wordDemo.state().hasWeights, null, { timeout: 30_000 })
await page.waitForTimeout(500)    // headings + weights refetch lands
const atEnd = await state()
console.log(`  completed after ${Date.now() - started} ms more; count ${atEnd.document.blockCount.toLocaleString()}, headings ${atEnd.headingCount}`)
check(atEnd.document.blockCount === atEnd.document.knownBlockCount, 'count is exact once complete')
check(!(await page.isVisible('[data-test="index-state"]')), 'index-state badge gone when complete')
check(atEnd.headingCount >= opened.headingCount, `headings refreshed on completion (${opened.headingCount} -> ${atEnd.headingCount})`)

const middle = Math.floor(atEnd.document.blockCount / 2)
step(`scrollToBlock(${middle})`)
await page.evaluate(i => window.__wordDemo.scrollToBlock(i), middle)
await page.waitForFunction(i => !!document.querySelector(`[data-block="${i}"]`), middle, { timeout: 30_000 })
await page.waitForTimeout(300)
const atMiddle = await state()
check(Math.abs((atMiddle.firstVisible?.index ?? -1) - middle) <= 2, `first visible row is ${middle} (+-2)`)
const last = atEnd.document.blockCount - 1
await page.evaluate(i => window.__wordDemo.scrollToBlock(i), last)
await page.waitForFunction(i => !!document.querySelector(`[data-block="${i}"]`), last, { timeout: 30_000 }).catch(() => {})
check(await page.$eval(`[data-block="${last}"]`, () => true).catch(() => false), 'the last block is rendered')

step('reopen - uses the saved sidecar')
await page.evaluate(path => window.__wordDemo.open(path), filePath)
await page.waitForFunction(() => window.__wordDemo.state().document?.isIndexComplete === true, null, { timeout: 60_000 })
const reopened = await state()
check(reopened.document.blockCount === atEnd.document.blockCount, `reopen count ${reopened.document.blockCount.toLocaleString()} matches`)

check(consoleErrors.length === 0, `no console errors${consoleErrors.length ? ': ' + consoleErrors.slice(0, 5).join(' | ') : ''}`)
await page.evaluate(() => fetch('/api/documents', { method: 'DELETE' }))
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
