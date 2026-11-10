// Live check of .odt charts in the app's FastDocView page: open test-data/odt-charts/odt-charts-v1.odt, check every chart
// is drawn as SVG filling its frame (not the preview picture), titles and series colours show, no console errors.
// Chart search is checked by the harness (OdtChartsFidelityTest). DOM only, no screenshots.
//   node playwright/odtChartsCheck.mjs <path to odt-charts-v1.odt>
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
await page.waitForFunction(() => document.querySelectorAll('svg.word-chart').length >= 3, null, { timeout: 60_000 })
const charts = await page.evaluate(() => [...document.querySelectorAll('svg.word-chart')].map(svg => {
  const s = svg.getBoundingClientRect(), f = svg.parentElement.getBoundingClientRect()
  return { sw: s.width, sh: s.height, fw: f.width, fh: f.height, text: svg.textContent, fills: [...svg.querySelectorAll('[fill]')].map(e => e.getAttribute('fill')) }
}))
console.log('  charts in view', charts.length)
check(charts.length >= 3 && charts.every(c => c.sw > 100 && Math.abs(c.sw - c.fw) <= 1.5 && Math.abs(c.sh - c.fh) <= 1.5), 'each chart SVG fills its frame')
check(charts[0].text.includes('Column Sales') && charts[0].text.includes('Quarter Axis') && charts[0].text.includes('Alpha'), 'first chart: title, axis title, legend')
check(charts[0].fills.some(f => f.includes('#004586')) && charts[0].fills.some(f => f.includes('#FF420E')), 'first chart: LibreOffice series colours ' + [...new Set(charts[0].fills)].join(' '))
check(await page.evaluate(() => ![...document.querySelectorAll('img')].some(i => (i.getAttribute('src') ?? '').includes('ObjectReplacements'))), 'no preview picture for drawn charts')
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
