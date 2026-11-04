// Live check of 3-D extruded custom shapes in the app's FastDocView page: open the --fidelity fixture
// (TestOutput/fidelity-v1-odt-extrusion.fodt: parallel, perspective, rotated, draft, unfilled, shaded), check each shape
// has its SVG with at least two faces filling its frame, draft faces are outlines, and no console errors. DOM only.
//   node playwright/odtExtrusionCheck.mjs <path to fidelity-v1-odt-extrusion.fodt>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
const names = ['parallel', 'perspective', 'rotated', 'draft', 'unfilled', 'shaded']
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
await page.waitForFunction(n => document.querySelectorAll('.word-shape svg').length >= n, names.length, { timeout: 60_000 })
const found = await page.evaluate(() => [...document.querySelectorAll('.word-shape')].map(frame => {
  const f = frame.getBoundingClientRect(), svg = frame.querySelector('svg'), s = svg?.getBoundingClientRect()
  const paths = svg ? [...svg.querySelectorAll('path')] : []
  return { fw: f.width, fh: f.height, sw: s?.width ?? 0, sh: s?.height ?? 0, faces: paths.length,
           filled: paths.filter(p => (p.getAttribute('fill') ?? '').startsWith('#')).length,
           block: frame.closest('p, [data-index]')?.textContent.trim().slice(0, 12) ?? '' }
}))
// match each shape by its paragraph's text (the shape's name), not by DOM order
const shapes = names.map(name => found.find(s => s.block.startsWith(name)))
check(found.length === names.length && shapes.every(Boolean), `${found.length} extruded shapes, one per mode`)
shapes.forEach((s, i) => s && check(s.faces >= 2 && s.sw > 0 && Math.abs(s.sw - s.fw) <= 1.5 && Math.abs(s.sh - s.fh) <= 1.5,
  `${names[i]}: SVG fills its frame with ${s.faces} faces`))
check(shapes[3]?.filled === 0, 'draft: faces are outlines')
check(shapes.filter((_, i) => i !== 3).every(s => s.filled >= 2), 'other modes: faces are filled')
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
