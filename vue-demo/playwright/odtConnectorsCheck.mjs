// Live check of ODT connectors glued to shapes in other blocks: open the --fidelity fixture TestOutput/cross-block-connectors-v1.fodt
// in the app's FastDocView page and check both connectors leave their start shape sideways (routed around the recorded shapes, not as
// free points), and no console errors. DOM only, no screenshots.
//   node playwright/odtConnectorsCheck.mjs <fixture>
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
await page.waitForFunction(() => document.querySelectorAll('.word-shape svg.word-shape-geometry').length >= 6, null, { timeout: 60_000 })
const paths = await page.evaluate(() => [...document.querySelectorAll('.word-shape svg.word-shape-geometry')]
  .map(svg => [...svg.querySelectorAll('path')].map(p => p.getAttribute('d')).filter(d => d && /L.*L/.test(d) && !/Z\s*$/i.test(d)).at(-1)).filter(Boolean))
check(paths.length === 2, `two routed connectors (${paths.length})`)
for (const d of paths) {
  const n = d.match(/-?\d+(\.\d+)?/g).map(Number)
  check(Math.abs(n[3] - n[1]) < 0.01 && Math.abs(n[2] - n[0]) > 0.01, `connector leaves its start shape sideways (${d})`)
}
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
process.exit(failures ? 1 : 0)
