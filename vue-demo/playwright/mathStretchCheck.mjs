// Stretched brackets and growing integrals against Word's PDF: opens a brackets-<font>-v1.docx (Word 16, 3-, 5- and 10-row
// matrices in ( [ { |, growing integrals over 3-, 5- and 10-level fractions; make-math-brackets.ps1) and prints, per
// "case <label>:" paragraph, the widths (pt) of its stretched operators as laid out, as JSON for
// scripts/compare-math-stretch.py. DOM only, no screenshots.   node playwright/mathStretchCheck.mjs <file> <out.json>
import { chromium } from 'playwright'
import { writeFileSync } from 'node:fs'
const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [filePath, outPath] = process.argv.slice(2)
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), filePath)
await page.waitForFunction(() => window.__fastDocViewTest?.state()?.isIndexComplete && window.__fastDocViewTest.state().blockCount > 0, null, { timeout: 60_000 })
const count = await page.evaluate(() => window.__fastDocViewTest.state().blockCount)
const result = {}
for (let i = 0; i < count; i++) {
  await page.evaluate(index => window.__fastDocViewTest.scrollToBlock(index), i)
  await page.waitForTimeout(150)
  const found = await page.evaluate(index => {
    const row = document.querySelector(`.windowed-scroller-viewport [data-index="${index}"]`)
    const label = /case ([^:]+):/.exec(row?.textContent ?? '')?.[1]
    if (!label) return null
    return { label, widths: [...row.querySelectorAll('mo.omath-stretch')].map(o => { const m = getComputedStyle(o); return +((o.getBoundingClientRect().width + parseFloat(m.marginLeft) + parseFloat(m.marginRight)) * 0.75).toFixed(1) }) }
  }, i)
  if (found) result[found.label] = found.widths
}
writeFileSync(outPath, JSON.stringify(result))
console.log(`${Object.keys(result).length} equations measured`)
await browser.close()
