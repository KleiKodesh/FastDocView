// Live check of RTF / .doc text boxes and drawn shapes in the app's FastDocView page: open a --fidelity fixture
// (TestOutput/fidelity-v1-rtf-drawn-shapes.rtf, or a .doc with text boxes), check the shape frames' SVG fills its frame,
// each expected text shows exactly once and no console errors. DOM only, no screenshots.
//   node playwright/rtfDocShapesCheck.mjs <file> <minimum frames> <text> [<text> ...]
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [filePath, minimumFrames, ...texts] = process.argv.slice(2)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
step(`open ${filePath}`)
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), filePath)
const wanted = Number(minimumFrames)
await page.waitForFunction(n => document.querySelectorAll('.word-shape svg.word-shape-geometry, .text-box').length >= n, wanted, { timeout: 60_000 })
const frames = await page.evaluate(() => [...document.querySelectorAll('.word-shape')].filter(f => f.querySelector('svg.word-shape-geometry')).map(frame => {
  const f = frame.getBoundingClientRect(), s = frame.querySelector('svg.word-shape-geometry').getBoundingClientRect()
  return { fw: f.width, fh: f.height, sw: s.width, sh: s.height }
}))
console.log('  frames', JSON.stringify(frames.map(f => [Math.round(f.sw), Math.round(f.sh)])))
const near = (a, b) => Math.abs(a - b) <= 1.5
check(frames.every(f => near(f.sw, f.fw) && near(f.sh, f.fh)), 'every SVG fills its frame')
const boxes = await page.evaluate(() => document.querySelectorAll('.text-box').length)
check(frames.length + boxes >= wanted, `${frames.length} shape frames + ${boxes} text boxes, at least ${wanted}`)
const body = await page.evaluate(() => document.querySelector('.windowed-scroller-viewport')?.textContent ?? '')
for (const t of texts) check(body.split(t).length - 1 === 1, `"${t}" shows once`)
check(errors.length === 0, `no console errors ${JSON.stringify(errors)}`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
