// Live probe (no screenshots): cold first open of a .docx, jump to the estimated end, and print
// what the scroller holds after the index completes — rendered range, first visible, count, and
// whether the real last block is mounted / on screen.
//   node playwright/probeEndOfIncrementalOpen.mjs <docx>
import { chromium } from 'playwright'

const filePath = process.argv[2]
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
await page.goto('http://localhost:5189/?file=', { waitUntil: 'networkidle' })
await page.evaluate(() => fetch('/api/documents', { method: 'DELETE' }))
await page.evaluate(path => window.__wordDemo.open(path, { rebuildIndex: true }), filePath)
await page.waitForSelector('[data-block]', { timeout: 120_000 })
const state = () => page.evaluate(() => window.__wordDemo.state())
for (let attempt = 0; attempt < 20 && !(await state()).document.isIndexComplete; attempt++) {
  await page.evaluate(() => window.__wordDemo.scrollToBlock(window.__wordDemo.state().document.blockCount - 1))
  await page.waitForTimeout(1500)
}
await page.waitForTimeout(1500)
const s = await state()
const last = s.document.blockCount - 1
const lastRow = await page.$eval(`[data-block="${last}"]`, row => {
  const rect = row.getBoundingClientRect()
  return { top: Math.round(rect.top), bottom: Math.round(rect.bottom), text: row.textContent.trim().length }
}).catch(() => null)
const viewport = await page.$eval('[data-test="scroller"]', el => ({ scrollTop: Math.round(el.scrollTop), scrollHeight: el.scrollHeight, clientHeight: el.clientHeight }))
console.log(JSON.stringify({ count: s.document.blockCount, known: s.document.knownBlockCount, complete: s.document.isIndexComplete,
  renderedRange: s.renderedRange, firstVisible: s.firstVisible, lastRow, viewport, lastFetch: s.lastFetch }, null, 1))
await page.evaluate(() => window.__wordDemo.scrollToBlock(window.__wordDemo.state().document.blockCount - 1))
await page.waitForTimeout(1000)
const after = await state()
const lastRowAfter = await page.$eval(`[data-block="${after.document.blockCount - 1}"]`, row => row.textContent.trim().length).catch(() => null)
console.log(JSON.stringify({ afterSecondJump: { renderedRange: after.renderedRange, firstVisible: after.firstVisible, lastRowTextLength: lastRowAfter } }))
await browser.close()
