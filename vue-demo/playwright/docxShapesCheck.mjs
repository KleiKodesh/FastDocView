// Live check of .docx drawn shapes in the app's FastDocView page: open the --fidelity fixture
// (TestOutput/fidelity-v1-shapes.docx), measure each shape frame's SVG box in the DOM, check every shape
// text shows once and no console errors. DOM only, no screenshots.
//   node playwright/docxShapesCheck.mjs <path to fidelity-v1-shapes.docx>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
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
await page.waitForFunction(() => document.querySelectorAll('.word-shape svg.word-shape-geometry').length >= 9, null, { timeout: 60_000 })
const frames = await page.evaluate(() => [...document.querySelectorAll('.word-shape')].map(frame => {
  const f = frame.getBoundingClientRect(), s = frame.querySelector('svg.word-shape-geometry').getBoundingClientRect()
  return { fw: f.width, fh: f.height, sw: s.width, sh: s.height, text: frame.textContent.trim() }
}))
const px = pt => pt * 96 / 72
const zoom = await page.evaluate(() => { const d = document.querySelector('.word-document'); return d ? d.getBoundingClientRect().width / d.offsetWidth : 1 })
console.log('  frames', JSON.stringify(frames.map(f => [Math.round(f.sw), Math.round(f.sh), f.text])))
const near = (a, b) => Math.abs(a - b) <= 1.5
check(frames.slice(0, 6).every(f => near(f.sw, px(72) * zoom) && near(f.sh, px(36) * zoom)), 'six 72 x 36 pt shapes: the SVG box is the frame (wp:extent)')
check(frames.every(f => near(f.sw, f.fw) && near(f.sh, f.fh)), 'every SVG fills its frame')
const vml = frames.find(f => f.text.includes('vmltext'))
check(!!vml && near(vml.sw, px(144) * zoom) && near(vml.sh, px(72) * zoom), 'VML roundrect: 2in x 1in')
const body = await page.evaluate(() => document.querySelector('.windowed-scroller-viewport')?.textContent ?? '')
for (const t of ['ovaltext', 'grouptext', 'nestedtext', 'vmltext', 'choicetext'])
  check(body.split(t).length - 1 === 1, `"${t}" shows once`)
check(!body.includes('fallbacktext'), 'the mc:Fallback copy does not show')
const groupText = await page.evaluate(() => {
  const span = [...document.querySelectorAll('.word-shape-text')].find(s => s.textContent.includes('grouptext'))
  const frame = span.closest('.word-shape').getBoundingClientRect(), r = span.getBoundingClientRect()
  return { dx: r.left - frame.left, w: r.width, h: r.height }
})
check(near(groupText.dx, px(36) * zoom) && near(groupText.w, px(36) * zoom) && near(groupText.h, px(18) * zoom), `group text box at its scaled child box (${JSON.stringify(groupText)})`)
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
