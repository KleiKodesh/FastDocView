// Live check of WordArt (text effects) in the app's FastDocView page: open the RTF shapes fixture
// (TestOutput/fidelity-v1-rtf-drawn-shapes.rtf) and check, in the DOM, that
// - the single-outline line (ArchText) got LibreOffice's fit factor: its font size changed from the line height and its
//   measured length now fills the arc (the widest line of the shape fills its outline),
// - the paired-outline line (WordArtText) was bent between its outlines: a bent image of the band's size replaces its paint,
//   the text itself stays (transparent) once in the page,
// - no console errors. DOM only, no screenshots.
//   node playwright/fontworkCheck.mjs <path to fidelity-v1-rtf-drawn-shapes.rtf>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
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
step(`open ${filePath}`)
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), filePath)
await page.waitForFunction(() => document.querySelector('g[data-fontwork-shape][data-fontwork-done]') !== null, null, { timeout: 60_000 })
await page.waitForTimeout(500)
const result = await page.evaluate(() => {
  const texts = [...document.querySelectorAll('text[data-fontwork]')]
  const single = texts.find(t => t.textContent === 'ArchText')
  const pair = texts.find(t => t.textContent === 'WordArtText')
  const singleOut = single && {
    height: Number(single.getAttribute('data-fontwork-height')),
    size: Number(single.getAttribute('font-size')),
    length: Number(single.getAttribute('data-fontwork-length')),
    measured: single.getComputedTextLength(),
  }
  const image = pair?.parentNode?.querySelector('image.fast-doc-view-fontwork-bent')
  const pairOut = pair && {
    opacity: pair.getAttribute('fill-opacity'),
    hasImage: !!image,
    dataLength: (image?.getAttribute('href') ?? '').length,
    width: Number(image?.getAttribute('width') ?? 0),
    height: Number(image?.getAttribute('height') ?? 0),
  }
  const body = document.querySelector('.windowed-scroller-viewport')?.textContent ?? ''
  return { singleOut, pairOut, archCount: body.split('ArchText').length - 1, wordArtCount: body.split('WordArtText').length - 1 }
})
console.log('  ' + JSON.stringify(result))
const s = result.singleOut
check(!!s && s.size !== s.height && Math.abs(s.measured - s.length) / s.length < 0.02, 'single outline: font size scaled by the fit factor, the line now fills its outline')
const p = result.pairOut
check(!!p && p.hasImage && p.opacity === '0' && p.dataLength > 1000 && Math.abs(p.width - 144) < 1 && Math.abs(p.height - 36) < 1, 'paired outlines: the text bent between them as an image of the band, the text kept transparent')
check(result.archCount === 1 && result.wordArtCount === 1, 'each WordArt text once in the page')
check(errors.length === 0, `no console errors ${JSON.stringify(errors)}`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
