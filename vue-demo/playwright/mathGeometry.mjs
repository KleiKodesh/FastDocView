// Equation sizes as the browser lays them out, for comparison with Word's PDF (scripts/compare-math-geometry.py).
// Opens a file in the live rig, visits every block, and writes {label: {width, height, baselineFromTop}} in points for each
// "case <label>:" paragraph's first equation (CSS px * 0.75 at zoom 100 %). DOM only, no screenshots.
//   node playwright/mathGeometry.mjs <file> <out.json>
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
  await page.waitForTimeout(60)
  const found = await page.evaluate(index => {
    const row = document.querySelector(`.windowed-scroller-viewport [data-index="${index}"]`)
    const text = row?.textContent ?? ''
    const label = /case ([^:]+):/.exec(text)?.[1]
    const math = row?.querySelector('math')
    if (!label || !math) return null
    const box = math.getBoundingClientRect()
    // the baseline: a zero-height inline-block probe after the math sits on the same baseline
    const probe = document.createElement('span')
    probe.style.cssText = 'display:inline-block;width:0;height:0'
    math.after(probe)
    const baseline = probe.getBoundingClientRect().bottom - box.top
    probe.remove()
    return { label, width: box.width * 0.75, height: box.height * 0.75, baselineFromTop: baseline * 0.75 }
  }, i)
  if (process.env.MATH_GEOMETRY_DEBUG && i < 3) console.log(i, JSON.stringify(found))
  if (found && !result[found.label]) result[found.label] = found
}
writeFileSync(outPath, JSON.stringify(result, null, 1))
console.log(`${Object.keys(result).length} equations measured`)
await browser.close()
