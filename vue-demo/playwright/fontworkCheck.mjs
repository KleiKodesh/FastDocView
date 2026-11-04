// Live check of WordArt (text effects) in the app's FastDocView page, DOM only (no screenshots):
//   node playwright/fontworkCheck.mjs <TestOutput/fidelity-v1-rtf-drawn-shapes.rtf> [<TestOutput/lo-converted/fidelity-v1-fontwork-curved.doc>]
// RTF fixture: the single-outline line (ArchText) got LibreOffice's fit factor (its size changed from the line height and its
// measured length fills the arc); the paired line (WordArtText) is bent - a mask of the bent text over a rectangle in the
// text's paint - with the text kept, transparent, once in the page.
// Curved .doc (optional): InflateText / WaveText are bent between curved outlines, the inflate keeps its gradient paint, and
// a device-pixel-ratio change re-bends them at the new resolution.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [rtfPath, curvedPath] = process.argv.slice(2)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 1800 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })

async function open(path) {
  step(`open ${path}`)
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), path)
  await page.waitForFunction(() => document.querySelector('g[data-fontwork-shape][data-fontwork-done]') !== null, null, { timeout: 60_000 })
  await page.waitForTimeout(500)
}

const bent = word => page.evaluate(w => {
  const text = [...document.querySelectorAll('text[data-fontwork="pair"]')].find(t => t.textContent === w)
  const group = text?.previousElementSibling
  const image = group?.querySelector('mask image')
  const rect = group?.querySelector('rect[mask]')
  return text && {
    opacity: text.getAttribute('fill-opacity'),
    isBent: !!group?.classList.contains('fast-doc-view-fontwork-bent'),
    dataLength: (image?.getAttribute('href') ?? '').length,
    width: Number(image?.getAttribute('width') ?? 0),
    height: Number(image?.getAttribute('height') ?? 0),
    paint: rect?.getAttribute('fill') ?? '',
    pixels: Number(text.getAttribute('data-fontwork-pixels') ?? 0),
  }
}, word)
const countInPage = word => page.evaluate(w => (document.querySelector('.windowed-scroller-viewport')?.textContent ?? '').split(w).length - 1, word)

await open(rtfPath)
const single = await page.evaluate(() => {
  const text = [...document.querySelectorAll('text[data-fontwork="single"]')].find(t => t.textContent === 'ArchText')
  return text && { height: Number(text.getAttribute('data-fontwork-height')), size: Number(text.getAttribute('font-size')),
                   length: Number(text.getAttribute('data-fontwork-length')), measured: text.getComputedTextLength() }
})
const pair = await bent('WordArtText')
console.log('  ' + JSON.stringify({ single, pair }))
check(!!single && single.size !== single.height && Math.abs(single.measured - single.length) / single.length < 0.02, 'single outline: font size scaled by the fit factor, the line fills its outline')
check(!!pair && pair.isBent && pair.opacity === '0' && pair.dataLength > 1000 && Math.abs(pair.width - 144) < 1 && Math.abs(pair.height - 36) < 1 && pair.paint === '#ff0000',
      'paired outlines: the bent text masks a rectangle in the text paint, the text kept transparent')
check(await countInPage('ArchText') === 1 && await countInPage('WordArtText') === 1, 'each WordArt text once in the page')

if (curvedPath) {
  await open(curvedPath)
  const inflate = await bent('InflateText'), wave = await bent('WaveText')
  console.log('  ' + JSON.stringify({ inflate, wave }))
  check(!!inflate && inflate.isBent && inflate.dataLength > 1000 && inflate.paint.startsWith('url(#word-officeart-gradient-'), 'inflate: bent between its curved outlines, painted with its gradient')
  // the gradient spans the bent text's own bounds (LibreOffice: all glyph outlines in one object with the shape's fill,
  // EnhancedCustomShapeFontWork.cxx:1096-1130): angle 0 runs from the ink's top to its bottom
  const span = await page.evaluate(() => {
    const text = [...document.querySelectorAll('text[data-fontwork="pair"]')].find(t => t.textContent === 'InflateText')
    const [x, y, w, h] = (text?.getAttribute('data-fontwork-ink') ?? '').split(' ').map(Number)
    const id = /url\(#([^)]+)\)/.exec(text?.previousElementSibling?.querySelector('rect[mask]')?.getAttribute('fill') ?? '')?.[1]
    const gradient = id ? document.getElementById(id) : null
    const band = text?.previousElementSibling?.querySelector('mask image')
    return { ink: { x, y, w, h }, id, y1: Number(gradient?.getAttribute('y1')), y2: Number(gradient?.getAttribute('y2')), bandHeight: Number(band?.getAttribute('height') ?? 0) }
  })
  console.log('  ' + JSON.stringify(span))
  check(!!span.id && span.id.includes('-text-') && Math.abs(span.y1 - span.ink.y) < 0.01 && Math.abs(span.y2 - (span.ink.y + span.ink.h)) < 0.01,
        'inflate: its gradient spans the bent ink (top to bottom), not the shape box')
  check(!!wave && wave.isBent && wave.dataLength > 1000 && wave.paint === '#0000ff', 'wave: bent between its curved outlines, in its colour')
  check(await countInPage('InflateText') === 1 && await countInPage('WaveText') === 1, 'curved WordArt texts once in the page')
  // a resolution change (browser zoom, another screen): the bent text is drawn again at the new density
  const session = await page.context().newCDPSession(page)
  await session.send('Emulation.setDeviceMetricsOverride', { width: 1300, height: 1800, deviceScaleFactor: 2, mobile: false })
  await page.waitForFunction(p => {
    const text = [...document.querySelectorAll('text[data-fontwork="pair"]')].find(t => t.textContent === 'InflateText')
    return text && Number(text.getAttribute('data-fontwork-pixels')) > p * 1.5
  }, inflate?.pixels ?? 1, { timeout: 10_000 }).catch(() => {})
  const after = await bent('InflateText')
  console.log('  ' + JSON.stringify({ before: inflate?.pixels, after: after?.pixels }))
  check(!!after && !!inflate && after.pixels > inflate.pixels * 1.5 && after.isBent && after.dataLength > inflate.dataLength,
        'device pixel ratio 1 -> 2: re-bent at the new resolution')
  // the page's own zoom (Ctrl + =, useZoom): the block grows, the bent text is drawn again at the new size
  await page.locator('.fast-doc-view-scroller').first().click({ position: { x: 5, y: 5 } }).catch(() => {})
  for (let i = 0; i < 4; i++) await page.keyboard.press('Control+Equal')
  await page.waitForFunction(p => {
    const text = [...document.querySelectorAll('text[data-fontwork="pair"]')].find(t => t.textContent === 'InflateText')
    return text && Number(text.getAttribute('data-fontwork-pixels')) > p * 1.02
  }, after?.pixels ?? 1, { timeout: 10_000 }).catch(() => {})
  const zoomed = await bent('InflateText')
  console.log('  ' + JSON.stringify({ beforeZoom: after?.pixels, afterZoom: zoomed?.pixels }))
  check(!!zoomed && !!after && zoomed.pixels > after.pixels * 1.02 && zoomed.isBent, 'page zoom in: re-bent at the larger size')
}

check(errors.length === 0, `no console errors ${JSON.stringify(errors)}`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
