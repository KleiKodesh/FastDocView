// Live check of 3-D shapes (a:scene3d / a:sp3d, OfficeArt f3D) in the app's FastDocView page: open
// test-data/word-shapes-3d/shapes-3d-v1.docx or shapes-3d-on-v1.rtf, check the 12 shapes each draw several lit faces inside an
// SVG that fills its frame, bevelled shapes many, and no console errors. DOM only.
//   node playwright/shapes3dCheck.mjs <path to shapes-3d-v1.docx | shapes-3d-on-v1.rtf>
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
await page.waitForFunction(() => document.querySelectorAll('.word-shape svg').length >= 12, null, { timeout: 60_000 })
const shapes = await page.evaluate(() => [...document.querySelectorAll('.word-shape')].map(frame => {
  const f = frame.getBoundingClientRect(), svg = frame.querySelector('svg'), s = svg?.getBoundingClientRect()
  const fills = new Set(svg ? [...svg.querySelectorAll('path')].map(p => p.getAttribute('fill')) : [])
  return { fw: f.width, fh: f.height, sw: s?.width ?? 0, sh: s?.height ?? 0, faces: svg?.querySelectorAll('path').length ?? 0, colours: fills.size }
}))
check(shapes.length === 12, `${shapes.length} shapes`)
shapes.forEach((s, i) => check(s.faces >= 1 && s.colours >= 1 && Math.abs(s.sw - s.fw) <= 1.5 && Math.abs(s.sh - s.fh) <= 1.5,
  `shape ${i}: ${s.faces} faces, ${s.colours} colours, SVG fills its frame`))
check(shapes.some(s => s.faces >= 20), 'a bevelled shape draws its bands')
check(shapes.filter(s => s.faces >= 2).length >= 10, 'turned and extruded shapes draw several faces (a front perspective view may show only its front)')
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
