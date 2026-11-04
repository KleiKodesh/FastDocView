// Live check of two-sided and contour wrap in the app's FastDocView page: open a file, find its floating objects, check each
// floats (never a centred block for a two-sided wrap), that a contour reaches the browser as a shape-outside polygon, that the
// text of a float's paragraph sits on the float's free side, and no console errors. DOM only, no screenshots.
//   node playwright/wrapCheck.mjs <file> <minimum floats> <minimum contours>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [filePath, minimumFloats, minimumContours] = process.argv.slice(2)
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
await page.waitForFunction(n => [...document.querySelectorAll('.word-shape, .text-box, img.word-picture')].filter(e => getComputedStyle(e).float !== 'none').length >= n,
  Number(minimumFloats), { timeout: 60_000 })
const floats = await page.evaluate(() => [...document.querySelectorAll('.word-shape, .text-box, img.word-picture')]
  .filter(e => getComputedStyle(e).float !== 'none' && !e.parentElement.closest('.word-shape, .text-box'))
  .map(e => {
    const style = getComputedStyle(e), box = e.getBoundingClientRect()
    // the paragraph's own text after the float: where its first line box starts
    const paragraph = e.closest('p')
    let textLeft = null, textRight = null
    if (paragraph) {
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT)
      for (let node; (node = walker.nextNode());) {
        if (node.parentElement.closest('.word-shape, .text-box') || !node.textContent.trim()) continue
        const range = document.createRange(); range.selectNodeContents(node)
        const r = range.getClientRects()[0]
        if (r && r.top < box.bottom) { textLeft = r.left; textRight = r.right; break }
      }
    }
    return { float: style.float, shape: style.shapeOutside, left: box.left, right: box.right, textLeft, textRight }
  }))
console.log('  floats', JSON.stringify(floats.map(f => [f.float, f.shape.slice(0, 30), Math.round(f.left), Math.round(f.right), f.textLeft && Math.round(f.textLeft)])))
check(floats.length >= Number(minimumFloats), `${floats.length} floating object(s), at least ${minimumFloats}`)
const contours = floats.filter(f => f.shape.startsWith('polygon('))
check(contours.length >= Number(minimumContours), `${contours.length} contour(s) as shape-outside polygon, at least ${minimumContours}`)
const beside = floats.filter(f => f.textLeft !== null && f.shape === 'none')
check(beside.every(f => f.float === 'left' ? f.textLeft >= f.right - 1 : f.textRight <= f.left + 1),
  `the text beside each box float runs along its free side (${beside.length} checked)`)
const contoured = floats.filter(f => f.textLeft !== null && f.shape.startsWith('polygon('))
console.log('  contour text', JSON.stringify(contoured.map(f => [f.float, Math.round(f.left), Math.round(f.right), Math.round(f.textLeft), Math.round(f.textRight)])))
check(errors.length === 0, `no console errors (${errors.join(' | ')})`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
