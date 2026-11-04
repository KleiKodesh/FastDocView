// Live check of ODT gradient angle units and automatic steps in the app's FastDocView page: open a --fidelity fixture
// (TestOutput/fidelity-v1-odt-angle-ooo.fodt / .odt: ODF 1.2 by OpenOffice.org 3.2, draw:angle="900" = 90 degrees; or -lo / -none:
// LibreOffice 7.3 / no generator = 900 degrees = 180) and check the turned gradient's rotation, the automatic-step gradient's
// 128 bands (under transparency), the plain one smooth, every paint reference resolves, and no console errors. DOM only, no screenshots.
//   node playwright/odtGradientCheck.mjs <fixture> <expected rotation, e.g. -90>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2], rotation = process.argv[3] ?? '-90'
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
await page.waitForFunction(() => document.querySelectorAll('.word-shape svg.word-shape-geometry').length >= 4, null, { timeout: 60_000 })
const shapes = await page.evaluate(() => [...document.querySelectorAll('.word-shape svg.word-shape-geometry')].map(svg => ({
  transforms: [...svg.querySelectorAll('[transform], [gradientTransform], [patternTransform]')].map(e => e.getAttribute('transform') ?? e.getAttribute('gradientTransform') ?? e.getAttribute('patternTransform')).join(' '),
  stops: svg.querySelectorAll('stop').length,
  colors: [...svg.querySelectorAll('stop')].map(s => getComputedStyle(s).stopColor),
})))
check(shapes.length === 4, `four gradient shapes (${shapes.length})`)
check(shapes[0]?.transforms.includes(`rotate(${rotation})`), `draw:angle="900" drawn as rotate(${rotation}) (${shapes[0]?.transforms.slice(0, 200)})`)
check(shapes[1]?.stops === 256, `automatic steps under transparency: 128 bands = 256 stops (${shapes[1]?.stops})`)
check(shapes[3]?.stops === 3 && shapes[3]?.colors.join() === 'rgb(0, 0, 255),rgb(0, 0, 255),rgb(255, 0, 0)', `radial with a 20 % border in Skia's layout (${shapes[3]?.colors})`)
check(shapes[2]?.stops === 2, `automatic steps, plain two-colour linear: smooth as Skia draws it (${shapes[2]?.stops} stops)`)
check(shapes[1]?.colors.includes('rgb(253, 0, 2)') && shapes[1]?.colors.at(-1) === 'rgb(0, 0, 255)', 'band 1 is #fd0002, the last band the end colour')
const refs = await page.evaluate(() => {
  const ids = new Set([...document.querySelectorAll('[id]')].map(e => e.id))
  const used = [...document.querySelectorAll('.word-shape [fill^="url(#"], .word-shape [mask^="url(#"]')].flatMap(e => [e.getAttribute('fill'), e.getAttribute('mask')]).filter(v => v && v.startsWith('url(#')).map(v => v.slice(5, -1))
  return { used: used.length, missing: used.filter(id => !ids.has(id)) }
})
check(refs.used >= 2 && refs.missing.length === 0, `every paint reference resolves (${refs.used} used, missing ${refs.missing})`)
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
