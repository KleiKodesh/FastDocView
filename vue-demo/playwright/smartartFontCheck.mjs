// Live check: SmartArt without a drawing part draws its text at Word's autofit size (font size of the first run of each
// shape, as Word saved it in the original's drawing part). DOM + console only.
//   node playwright/smartartFontCheck.mjs <name>=<size pt> ...   (files: test-data/smartart-layouts/<name>-v1-nodrawing.docx)
import { chromium } from 'playwright'
import path from 'node:path'
const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const folder = path.resolve(import.meta.dirname, '..', '..', 'test-data', 'smartart-layouts')
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const browser = await chromium.launch({ headless: true })
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
for (const spec of process.argv.slice(2)) {
  const [name, want] = spec.split('=')
  const file = path.join(folder, `${name}-v1-nodrawing.docx`)
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), file)
  await page.waitForFunction(() => document.querySelectorAll('.windowed-scroller-viewport svg path').length > 0, null, { timeout: 60_000 })
  await page.waitForTimeout(300)
  // the text over the shapes: the font sizes of elements holding the diagram's text
  const sizes = await page.evaluate(() => {
    const found = new Set()
    for (const el of document.querySelectorAll('.windowed-scroller-viewport *')) {
      if (el.children.length || !(el.textContent ?? '').trim()) continue
      const fs = el.style?.fontSize || el.getAttribute?.('font-size') || ''
      if (fs) found.add(fs)
    }
    return [...found]
  })
  check(sizes.some(s => parseFloat(s) === Number(want)), `${name}: text at ${want} pt (sizes in the page: ${sizes.join(', ')})`)
}
check(errors.length === 0, `no console errors${errors.length ? ': ' + errors.join(' | ') : ''}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
