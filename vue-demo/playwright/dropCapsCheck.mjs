// Live check of drop caps and automatic hyphenation in the app's FastDocView page, DOM measurements only:
//   node playwright/dropCapsCheck.mjs <test-data/drop-caps-hyphenation/drop-caps-hyphenation-v1.docx> [.doc] [.rtf] [.odt]
// The drop cap spans N text lines (its box is at least N-1 line heights tall and the paragraph text flows beside it, not
// below it), and the hyphenated paragraph has hyphens:auto with its lang.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
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

for (const path of process.argv.slice(2)) {
  step(`open ${path}`)
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), path)
  await page.waitForFunction(() => document.querySelectorAll('.drop-cap').length > 0, null, { timeout: 60_000 })
  await page.waitForTimeout(400)
  const result = await page.evaluate(() => {
    const cap = document.querySelector('.drop-cap')
    const paragraph = cap.closest('p, h1, h2, h3, h4, h5, h6')
    const style = getComputedStyle(cap)
    const capBox = cap.getBoundingClientRect()
    const lineHeight = parseFloat(getComputedStyle(paragraph).lineHeight) || parseFloat(getComputedStyle(paragraph).fontSize) * 1.15
    const zoom = paragraph.getBoundingClientRect().height / parseFloat(getComputedStyle(paragraph).height)
    // the second text line starts beside the initial, not under it: a text range just after the initial
    const range = document.createRange()
    const text = cap.nextSibling ?? cap.parentNode.nextSibling
    const rects = [...paragraph.getClientRects()]
    const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT)
    let node, lineTops = new Set(), besideCount = 0
    while ((node = walker.nextNode())) {
      if (cap.contains(node)) continue
      range.selectNodeContents(node)
      for (const r of range.getClientRects()) {
        lineTops.add(Math.round(r.top))
        if (r.top < capBox.bottom - 2 && (r.left >= capBox.right - 1 || r.right <= capBox.left + 1)) besideCount++
      }
    }
    return {
      initialLetter: style.initialLetter ?? style.getPropertyValue('initial-letter'), float: style.float,
      capHeight: capBox.height, lineHeight: lineHeight * zoom, besideCount, lines: lineTops.size,
      hyphens: getComputedStyle(paragraph).hyphens, lang: paragraph.getAttribute('lang'), rects: rects.length,
    }
  })
  console.log('   ', JSON.stringify(result))
  check(result.capHeight >= result.lineHeight * 1.8, `${path}: the initial is taller than two lines (${result.capHeight.toFixed(1)} px, line ${result.lineHeight.toFixed(1)} px)`)
  check(result.besideCount >= 2, `${path}: at least two text lines flow beside the initial (${result.besideCount})`)
  check(result.hyphens === 'auto' && /^en/.test(result.lang ?? ''), `${path}: hyphens:auto with lang ${result.lang}`)
}
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
