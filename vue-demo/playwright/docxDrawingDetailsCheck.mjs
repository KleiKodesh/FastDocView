// Live check of .docx drawing details in the app's FastDocView page: open the --fidelity fixture
// (TestOutput/fidelity-v1-drawing-details.docx) and read the drawn SVG from the DOM: the browser parses every path
// (getTotalLength works), line-end markers sit at the line ends, opacity / dash / gradient attributes reach the page,
// the gradient ids resolve, and there are no console errors. DOM only, no screenshots.
//   node playwright/docxDrawingDetailsCheck.mjs <path to fidelity-v1-drawing-details.docx>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 2400 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
step(`open ${filePath}`)
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), filePath)
await page.waitForFunction(() => document.querySelectorAll('.word-shape svg.word-shape-geometry').length >= 18, null, { timeout: 60_000 })
step('read the drawn SVG')
const facts = await page.evaluate(() => {
  const svgs = [...document.querySelectorAll('.word-shape svg.word-shape-geometry')]
  const paths = svgs.flatMap(svg => [...svg.querySelectorAll('path')])
  const broken = paths.filter(p => { try { return !(p.getTotalLength() >= 0) } catch { return true } }).length
  const urls = paths.map(p => p.getAttribute('fill') ?? '').filter(f => f.startsWith('url(#')).map(f => f.slice(5, -1))
  const missing = urls.filter(id => !document.getElementById(id)).length
  const lineEnd = svgs[3] ? [...svgs[3].querySelectorAll('path')].map(p => { const b = p.getBBox(); return [p.getAttribute('fill'), Math.round(b.x), Math.round(b.y), Math.round(b.width), Math.round(b.height)] }) : []
  return {
    svgs: svgs.length, paths: paths.length, broken, gradients: urls.length, missing, lineEnd,
    opacity: document.querySelectorAll('.word-shape path[fill-opacity], .word-shape path[stroke-opacity], .word-shape stop[stop-opacity]').length,
    dashes: document.querySelectorAll('.word-shape path[stroke-dasharray]').length,
    radial: document.querySelectorAll('.word-shape radialGradient').length,
  }
})
console.log('  facts', JSON.stringify(facts))
check(facts.svgs >= 18 && facts.broken === 0, `every drawn path parses (${facts.paths} paths in ${facts.svgs} drawings)`)
check(facts.gradients >= 4 && facts.missing === 0 && facts.radial >= 1, 'gradient fills resolve to their definitions, the radial one included')
check(facts.opacity >= 4, 'alpha reaches the page as fill / stroke / stop opacity')
check(facts.dashes >= 5, 'dash patterns reach the page')
// drawing 3: 2pt line 72pt long, oval head (4 x 4 pt) centred on the start, triangle tail 6 x 6 pt with its tip on the end
const px = pt => pt * 96 / 72
const markers = facts.lineEnd.filter(m => m[0] !== 'none')
check(markers.length === 2 && markers.some(m => Math.abs(m[1] + 2) <= 1 && Math.abs(m[3] - 4) <= 1) && markers.some(m => Math.abs(m[1] - 66) <= 1 && Math.abs(m[3] - 6) <= 1),
      `line ends measured in the SVG: oval at the start, triangle at the end (${JSON.stringify(facts.lineEnd)}; px(72) = ${px(72)})`)
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
