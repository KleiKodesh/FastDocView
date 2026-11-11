// Live check of .doc / .rtf charts (embedded Excel chart objects) in the app's FastDocView page: open
// test-data/word-charts-binary/word-charts-binary-v2.doc or .rtf, check the charts are drawn as SVG filling their frame (not
// the preview picture), with titles, legend and Word's series colours, and no console errors. DOM only, no screenshots.
//   node playwright/binaryChartsCheck.mjs <path to word-charts-binary-v2.doc or .rtf>
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
await page.waitForFunction(() => document.querySelectorAll('svg.word-chart').length >= 2, null, { timeout: 60_000 })
const charts = await page.evaluate(() => [...document.querySelectorAll('svg.word-chart')].map(svg => {
  const s = svg.getBoundingClientRect(), f = svg.parentElement.getBoundingClientRect()
  return { sw: s.width, sh: s.height, fw: f.width, fh: f.height, text: svg.textContent, fills: [...svg.querySelectorAll('[fill]')].map(e => e.getAttribute('fill')) }
}))
console.log('  charts in view', charts.length)
check(charts.length >= 2 && charts.every(c => c.sw > 100 && Math.abs(c.sw - c.fw) <= 1.5 && Math.abs(c.sh - c.fh) <= 1.5), 'each chart SVG fills its frame')
check(charts[0].text.includes('Column Chart') && charts[0].text.includes('North') && charts[0].text.includes('Q1'), 'first chart: title, legend, categories')
check(charts[0].fills.some(f => f.includes('#4472C4')) && charts[0].fills.some(f => f.includes('#ED7D31')), "first chart: Word's series colours")
check(await page.evaluate(() => !document.querySelector('img.word-picture[data-word-picture="0"]')), 'no preview picture for the first drawn chart')
await page.evaluate(() => document.querySelector('.windowed-scroller-viewport')?.scrollTo(0, 1e7))
await page.waitForFunction(() => document.querySelectorAll('.word-chart-box').length > 0, null, { timeout: 60_000 }).catch(() => {})
check(await page.evaluate(() => document.querySelectorAll('.word-chart-box').length > 0), '3-D column chart drawn as boxes at the end')
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
