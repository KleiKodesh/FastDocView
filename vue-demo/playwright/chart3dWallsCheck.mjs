// Live check of the 3-D chart walls in the app's FastDocView page (2026-11-24 Word probes): horizontal 3-D bars paint the floor's
// fill on the face under the bars' bases and the side wall's own fill on the bottom face; at rotY 90 the back wall stands on the
// floor's front edge (left of the floor). DOM only, no screenshots, no console errors.
//   node playwright/chart3dWallsCheck.mjs <test-data/word-charts-scene-types/types-006.docx> <test-data/word-charts-scene-grid/grid-008.docx>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [typesFile, gridFile] = process.argv.slice(2)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
// faces of every drawn chart: class, fill (the page wraps it in light-dark(), so colours are matched as substrings) and the x range of the path's points
const faces = () => page.evaluate(() => [...document.querySelectorAll('svg.word-chart')].map(svg => [...svg.querySelectorAll('path.word-chart-wall, path.word-chart-floor')].map(p => {
  const xs = (p.getAttribute('d').match(/-?[0-9.]+/g) ?? []).map(Number).filter((_, i) => i % 2 === 0)
  return { cls: p.getAttribute('class'), fill: p.getAttribute('fill'), minX: Math.min(...xs), maxX: Math.max(...xs) }
})))
async function open(file) {
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), file)
  await page.waitForFunction(() => document.querySelectorAll('svg.word-chart').length >= 2, null, { timeout: 60_000 })
}
// types-006: the chart list starts at probe 144; probes 144..147 are horizontal clustered bars (make-word-charts-scene-types.py)
await open(typesFile)
let all = await faces()
const bars = all.slice(0, 4).filter(f => f.length === 3)
check(bars.length > 0 && bars.every(f => f[1].fill.toUpperCase().includes('#FF0000') && f[1].cls === 'word-chart-floor' && f[2].fill.toUpperCase().includes('#0000FF')),
  `horizontal 3-D bars: floor fill on the bars' base face, side wall fill underneath (${bars.length} charts)`)
check(bars.every(f => f[1].minX < f[2].minX && f[1].minX < f[0].minX), 'horizontal 3-D bars: the floor-coloured face is the leftmost one')
// grid-008: probes 192..215; rotY 90 ones are back wall left of the floor
await open(gridFile)
all = await faces()
const rot90 = all.filter(f => f.length === 3 && f[0].fill.toUpperCase().includes('#00FF00'))
  .filter(f => f[0].maxX - f[0].minX < (f[2].maxX - f[2].minX) * 0.6)   // the back wall seen nearly edge-on
check(rot90.length > 0 && rot90.every(f => f[0].minX <= f[2].minX + 1), `rotY 90: back wall on the floor's front (left) edge (${rot90.length} charts)`)
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
