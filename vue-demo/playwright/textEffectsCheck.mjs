// Live check of character scale and Word 2010 text effects in the app's FastDocView page, DOM measurements only:
//   node playwright/textEffectsCheck.mjs <TestOutput/fidelity-v1-char-scale.docx> <test-data/word-text-effects/text-effects-v1.doc> <TestOutput/fidelity-v1-text-fill.docx> <TestOutput/fidelity-v1-glow-shadow.docx>
// Scale: every scaled word box is exactly factor x its text's natural width, the scaled text fills the box (not shifted),
// and a long scaled paragraph still wraps onto several lines inside the page. Effects: the computed gradient clip and text-shadow.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [scaleDocx, scaleDoc, fillDocx, glowDocx] = process.argv.slice(2)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })

async function open(path, ready) {
  step(`open ${path}`)
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), path)
  await page.waitForFunction(ready, null, { timeout: 60_000 })
  await page.waitForTimeout(400)
}

const measureScale = () => page.evaluate(() => [...document.querySelectorAll('.char-scale')].map(box => {
  const inner = box.firstElementChild
  const factor = Number(/scaleX\(([\d.]+)\)/.exec(inner.getAttribute('style'))[1])
  const outer = box.getBoundingClientRect(), drawn = inner.getBoundingClientRect()
  // the page zooms its text (CSS zoom): the natural width in screen pixels = layout width x the zoom the box's rect shows
  const zoom = outer.width / parseFloat(getComputedStyle(box).width)
  return { text: inner.textContent, factor, boxWidth: outer.width, natural: parseFloat(getComputedStyle(inner).width) * zoom, drawnLeft: drawn.left, drawnWidth: drawn.width, left: outer.left, top: outer.top }
}))

for (const path of [scaleDocx, scaleDoc]) {
  await open(path, () => document.querySelectorAll('.char-scale').length > 0)
  const boxes = await measureScale()
  const exact = boxes.filter(b => Math.abs(b.boxWidth - b.natural * b.factor) < 0.6 && Math.abs(b.drawnWidth - b.boxWidth) < 0.6 && Math.abs(b.drawnLeft - b.left) < 0.6)
  check(boxes.length > 0 && exact.length === boxes.length, `${path.split(/[\\/]/).pop()}: ${exact.length}/${boxes.length} word boxes = factor x natural width, text drawn over the box (e.g. ${JSON.stringify(boxes[0])})`)
  const factors = [...new Set(boxes.map(b => b.factor))]
  check(factors.includes(0.8) && (path.endsWith('.doc') ? factors.includes(1.5) : true), `factors seen: ${factors}`)
  if (path === scaleDocx) {
    const long = boxes.filter(b => /^word\d+$/.test(b.text))
    const lines = new Set(long.map(b => Math.round(b.top))).size
    const pageRight = await page.evaluate(() => Math.max(...[...document.querySelectorAll('.char-scale')].map(b => b.closest('p,div,h1,h2,h3')?.getBoundingClientRect().right ?? 0)))
    check(long.length === 80 && lines >= 3 && long.every(b => b.left + b.boxWidth <= pageRight + 1), `long scaled paragraph wraps: ${long.length} words on ${lines} lines, none past the paragraph's edge`)
  }
}

await open(fillDocx, () => [...document.querySelectorAll('span')].some(s => s.textContent === 'linear'))
const fills = await page.evaluate(() => Object.fromEntries(['alpha', 'linear', 'radial'].map(word => {
  const span = [...document.querySelectorAll('span[style]')].find(s => s.textContent === word)
  const cs = span && getComputedStyle(span)
  return [word, cs && { color: cs.color, clip: cs.backgroundClip || cs.webkitBackgroundClip, image: cs.backgroundImage }]
})))
check(fills.alpha?.color === 'rgba(255, 0, 0, 0.6)', `solid fill with alpha: ${JSON.stringify(fills.alpha)}`)
check(fills.linear?.clip === 'text' && fills.linear.image.startsWith('linear-gradient(rgb(255, 0, 0)') && fills.linear.color === 'rgba(0, 0, 0, 0)', `linear gradient clipped to the text: ${JSON.stringify(fills.linear)}`)
check(fills.radial?.clip === 'text' && fills.radial.image.startsWith('radial-gradient(circle'), `radial gradient clipped to the text: ${JSON.stringify(fills.radial)}`)

await open(glowDocx, () => [...document.querySelectorAll('span')].some(s => s.textContent === 'glow'))
const shadows = await page.evaluate(() => Object.fromEntries(['glow', 'shadow', 'legacy'].map(word => {
  const span = [...document.querySelectorAll('span[style]')].find(s => s.textContent === word)
  return [word, span && getComputedStyle(span).textShadow]
})))
check(/rgba\(0, 176, 240, 0\.4\) 0px 0px 6\.6\d*px, rgba\(0, 176, 240, 0\.4\) 0px 0px 13\.3\d*px/.test(shadows.glow ?? ''), `glow: ${shadows.glow}`)
check(/rgba\(0, 0, 0, 0\.4\) 1\.8\d*px 1\.8\d*px 4px/.test(shadows.shadow ?? ''), `shadow: ${shadows.shadow}`)
check((shadows.legacy ?? 'none') !== 'none', `legacy shadow: ${shadows.legacy}`)

check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILED`)
process.exit(failures === 0 ? 0 : 1)
