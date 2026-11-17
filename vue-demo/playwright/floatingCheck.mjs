// Live check of .doc / RTF floating text boxes, groups and shapes in the app's FastDocView page: open
// test-data/word-shapes/floating-v1.doc / .rtf or floating-group-v1.* and check each marker's outer frame sits where Word
// puts it in the page's own layout (left / right / centred in its block, or at its offset from the column), the text it
// floats beside goes around it, each marker shows once, no console errors. DOM only, no screenshots.
//   node playwright/floatingCheck.mjs <file> <marker>=<left|right|center|offset:N:W|down:N:H> [...]   (offset N pt for a frame W pt wide, or N pt below its block's top for a frame H pt high: zoom-free)
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [filePath, ...wanted] = process.argv.slice(2)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
step(`open ${filePath}`)
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), filePath)
const markers = wanted.map(w => w.split('='))
await page.waitForFunction(names => names.every(n => (document.querySelector('.windowed-scroller-viewport')?.textContent ?? '').includes(n)),
  markers.map(m => m[0]), { timeout: 60_000 })
for (const [marker, expected] of markers) {
  const placed = await page.evaluate(name => {
    const walker = document.createTreeWalker(document.querySelector('.windowed-scroller-viewport'), NodeFilter.SHOW_TEXT)
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (!node.textContent.includes(name)) continue
      let frame = null
      for (let e = node.parentElement; e; e = e.parentElement)
        if ((e.classList.contains('word-shape') || e.classList.contains('text-box')) && getComputedStyle(e).position !== 'absolute') frame = e
      if (!frame) return null
      const block = frame.parentElement.closest('p, div, li') ?? frame.parentElement
      const f = frame.getBoundingClientRect(), b = block.getBoundingClientRect()
      return { width: f.width, height: f.height, top: f.top - b.top, left: f.left - b.left, right: b.right - f.right, float: getComputedStyle(frame).float, display: getComputedStyle(frame).display }
    }
    return null
  }, marker)
  if (!placed) { check(false, `${marker}: no frame`); continue }
  const near = (a, b) => Math.abs(a - b) <= 2
  const ok = expected === 'left' ? near(placed.left, 0)
    : expected === 'right' ? near(placed.right, 0)
    : expected === 'center' ? near(placed.left, placed.right)
    : expected.startsWith('down') ? near(placed.top, Number(expected.split(':')[1]) * placed.height / Number(expected.split(':')[2]))
    : near(placed.left, Number(expected.split(':')[1]) * placed.width / Number(expected.split(':')[2]))
  check(ok, `${marker}: ${expected} (frame ${Math.round(placed.left)} px from the block's left, ${Math.round(placed.right)} px from its right, ${Math.round(placed.top)} px below its top, float ${placed.float})`)
  const body = await page.evaluate(() => document.querySelector('.windowed-scroller-viewport')?.textContent ?? '')
  check(body.split(marker).length - 1 === 1, `"${marker}" shows once`)
}
check(errors.length === 0, `no console errors ${JSON.stringify(errors)}`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
