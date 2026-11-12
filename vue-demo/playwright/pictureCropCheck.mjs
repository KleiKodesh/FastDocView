// Live check of picture cropping in the app's FastDocView page, DOM measurements only:
//   node playwright/pictureCropCheck.mjs <test-data/word-picture-crop/picture-crop-v1.docx> [.doc] [.rtf] [.odt]
// Each cropped picture: its box keeps the shown size (Word's PDF: 104.93 / 164.9 / 119.95 pt wide), the picture inside has
// loaded and is drawn at the whole picture's width (about 150 pt) at its crop offset; the box hides what overflows.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const expected = [[104.93, -14.99, -7.49], [164.9, 14.99, 7.49], [119.95, -29.99, 0]]

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1200, height: 1400 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
for (const path of process.argv.slice(2)) {
  console.log(`open ${path.split(/[\/]/).pop()}`)
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), path)
  await page.waitForFunction(() => [...document.querySelectorAll('.word-picture-crop img')].filter(i => i.complete && i.naturalWidth > 0).length >= 3, null, { timeout: 60_000 })
  const boxes = await page.evaluate(() => [...document.querySelectorAll('.word-picture-crop')].map(box => {
    const zoom = box.getBoundingClientRect().width / parseFloat(getComputedStyle(box).width)
    const b = box.getBoundingClientRect(), i = box.querySelector('img').getBoundingClientRect()
    const pt = px => px / zoom * 0.75
    return { box: pt(b.width), whole: pt(i.width), left: pt(i.left - b.left), top: pt(i.top - b.top), overflow: getComputedStyle(box).overflow, natural: box.querySelector('img').naturalWidth }
  }))
  check(boxes.length === 3, `three cropped pictures (${boxes.length})`)
  boxes.forEach((r, n) => {
    const [box, left, top] = expected[n]
    check(Math.abs(r.box - box) < 2 && Math.abs(r.whole - 150) < 2 && Math.abs(r.left - left) < 2 && Math.abs(r.top - top) < 2 && r.overflow === 'hidden' && r.natural === 200,
          `picture ${n + 1}: box ${r.box.toFixed(2)} pt, whole ${r.whole.toFixed(2)} pt at ${r.left.toFixed(2)} / ${r.top.toFixed(2)} pt, overflow ${r.overflow}, loaded ${r.natural} px`)
  })
}
check(errors.length === 0, `no console errors (${errors.join(' | ')})`)
await browser.close()
process.exit(failures ? 1 : 0)
