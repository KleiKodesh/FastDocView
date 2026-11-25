// Matrix row pitch against Word's: opens a document with Office Math matrices (test-data/word-math/brackets-*-v1.docx,
// math-v2.docx) and prints, per matrix, the distance between its rows' baselines in em of the matrix's font size, plus the
// stretched operators' widths (pt). Word 16's PDFs: one math line, 1.18 em in Cambria Math (brackets-cambria-v1.pdf),
// 1.2 em in DejaVu Math TeX Gyre. DOM only, no screenshots.   node playwright/matrixPitchCheck.mjs <file>
import { chromium } from 'playwright'
const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [filePath] = process.argv.slice(2)
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', message => { if (message.type() === 'error') errors.push(message.text()) })
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), filePath)
await page.waitForFunction(() => window.__fastDocViewTest?.state()?.isIndexComplete && window.__fastDocViewTest.state().blockCount > 0, null, { timeout: 60_000 })
const count = await page.evaluate(() => window.__fastDocViewTest.state().blockCount)
const pitches = []
for (let i = 0; i < count; i++) {
  await page.evaluate(index => window.__fastDocViewTest.scrollToBlock(index), i)
  await page.waitForTimeout(150)
  const found = await page.evaluate(index => {
    const row = document.querySelector(`.windowed-scroller-viewport [data-index="${index}"]`)
    if (!row) return []
    const canvas = document.createElement('canvas').getContext('2d')
    // a row's baseline: its first token's box top + the font's ascent share of the box
    const baselineOf = mtr => {
      const token = mtr.querySelector('mi, mn, mtext')
      if (!token) return null
      const style = getComputedStyle(token)
      canvas.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`
      const m = canvas.measureText(token.textContent)
      const box = token.getBoundingClientRect()
      const line = m.fontBoundingBoxAscent + m.fontBoundingBoxDescent
      return Math.abs(box.height - line) < Math.abs(box.height - m.actualBoundingBoxAscent - m.actualBoundingBoxDescent) ? box.top + m.fontBoundingBoxAscent * box.height / line : box.top + m.actualBoundingBoxAscent
    }
    return [...row.querySelectorAll('mtable.omath-matrix')].map(table => {
      const size = parseFloat(getComputedStyle(table).fontSize)
      const baselines = [...table.children].map(baselineOf).filter(b => b !== null)
      const gaps = baselines.slice(1).map((b, k) => +((b - baselines[k]) / size).toFixed(3))
      const label = /case ([^:]+):/.exec(row.textContent ?? '')?.[1] ?? `block ${index}`
      return { label, rows: baselines.length, gaps, height: +(table.getBoundingClientRect().height / size).toFixed(3) }
    })
  }, i)
  pitches.push(...found)
}
for (const p of pitches) console.log(`${p.label}: rows ${p.rows}, baseline gaps (em) ${p.gaps.join(' ')}, table ${p.height} em`)
console.log(`${pitches.length} matrices; console errors: ${errors.length}`)
for (const e of errors) console.log('  ' + e)
await browser.close()
