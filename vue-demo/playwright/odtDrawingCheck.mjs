// Live check of OpenDocument drawing in the app's FastDocView page: open the --fidelity fixture
// (TestOutput/fidelity-v1-odt-drawing.fodt), check every shape frame has its SVG at its size, gradients / masks /
// patterns resolve, shape texts show once, the text box's table is not split out of its paragraph, the contour float
// has its shape-outside, and no console errors. DOM only, no screenshots.
//   node playwright/odtDrawingCheck.mjs <path to fidelity-v1-odt-drawing.fodt>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), filePath)
await page.waitForFunction(() => document.querySelectorAll('.word-shape svg.word-shape-geometry').length >= 10, null, { timeout: 60_000 })
const zoom = await page.evaluate(() => { const d = document.querySelector('.word-document'); return d ? d.getBoundingClientRect().width / d.offsetWidth : 1 })
const px = pt => pt * 96 / 72 * zoom
const near = (a, b) => Math.abs(a - b) <= 1.5
const frames = await page.evaluate(() => [...document.querySelectorAll('.word-shape')].map(frame => {
  const f = frame.getBoundingClientRect(), s = frame.querySelector('svg.word-shape-geometry')?.getBoundingClientRect()
  return { fw: f.width, fh: f.height, sw: s?.width ?? 0, sh: s?.height ?? 0, text: frame.textContent.trim().slice(0, 30) }
}))
console.log('  frames', frames.length)
check(frames.every(f => near(f.sw, f.fw) && near(f.sh, f.fh)), 'every SVG fills its frame')
check(near(frames[0].fw, px(144)) && near(frames[0].fh, px(72)), 'the enhanced-path shape is 2 x 1 in')
const refs = await page.evaluate(() => {
  const ids = new Set([...document.querySelectorAll('[id]')].map(e => e.id))
  const used = [...document.querySelectorAll('.word-shape [fill^="url(#"], .word-shape [mask^="url(#"]')].flatMap(e => [e.getAttribute('fill'), e.getAttribute('mask')]).filter(v => v && v.startsWith('url(#')).map(v => v.slice(5, -1))
  return { used: used.length, missing: used.filter(id => !ids.has(id)) }
})
check(refs.used >= 9 && refs.missing.length === 0, `every gradient / pattern / mask reference resolves (${refs.used} used, missing ${refs.missing})`)
const body = await page.evaluate(() => document.querySelector('.windowed-scroller-viewport')?.textContent ?? '')
for (const t of ['owntext', 'fontworkword', 'turnedtext', 'listone', 'cellone', 'areatext', 'pagetext', 'after box'])
  check(body.split(t).length - 1 === 1, `"${t}" shows once`)
check(await page.evaluate(() => !!document.querySelector('.word-paragraph-block table')), 'the text box table stays inside its paragraph block')
check(await page.evaluate(() => [...document.querySelectorAll('.word-shape')].some(f => getComputedStyle(f).shapeOutside.startsWith('polygon'))), 'contour wrap: shape-outside applied')
const textPath = await page.evaluate(() => { const t = [...document.querySelectorAll('textPath')].find(e => e.textContent === 'fontworkword'); return t ? t.getComputedTextLength() : 0 })
check(textPath > 0, `Fontwork text laid on its path (${Math.round(textPath)} px)`)
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
