// Live check of the tab layout (the page's fastDocViewTabs pass) against Word's and LibreOffice's PDFs, DOM measurements only:
//   node playwright/tabStopsCheck.mjs <reference.json> <file> [<reference.json> <file> ...]
// The reference is scripts/measure-tab-stops-pdf.py's output for test-data/tab-stops/tab-stops-v2 (Word's PDF for .docx / .doc /
// .rtf, LibreOffice's for .odt): per paragraph, the x (pt from the left margin) where the text after each tab starts. Each must
// match within 2 pt; the leaders are drawn; the text keeps its tab characters.
import { chromium } from 'playwright'
import { readFileSync } from 'node:fs'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const TOLERANCE = 2
const TEXT_WIDTH = Number(process.env.TAB_TEXT_WIDTH ?? 468)       // the fixture's page text width (Letter, 1 in margins)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1200, height: 1600 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })

const args = process.argv.slice(2)
for (let a = 0; a + 1 < args.length; a += 2) {
  const { paragraphs: reference, bars: referenceBars } = JSON.parse(readFileSync(args[a], 'utf8'))
  const path = args[a + 1]
  step(`open ${path}`)
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), path)
  await page.waitForFunction(n => document.querySelectorAll('.fast-doc-view-block-text p').length >= n && document.querySelectorAll('.fast-doc-view-block-text .tab-space').length > 0, 20, { timeout: 60_000 })
  // the column as wide as the page's text area (the demo column follows the window): lines wrap where Word wraps them; the
  // width change also re-runs the tab pass (its ResizeObserver, next frame)
  await page.evaluate(async (textWidth) => {
    for (const block of document.querySelectorAll('.fast-doc-view-block-text')) block.style.width = `${textWidth}pt`
    await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(() => requestAnimationFrame(r))))
  }, TEXT_WIDTH)
  const measured = await page.evaluate((textWidth) => {
    const paragraphs = [...document.querySelectorAll('.fast-doc-view-block-text p, .fast-doc-view-block-text h1, .fast-doc-view-block-text h2')]
    return paragraphs.map(p => {
      const probe = document.createElement('span')
      probe.style.cssText = 'display:inline-block;width:100pt'
      p.appendChild(probe)
      const pixelsPerPoint = probe.getBoundingClientRect().width / 100
      probe.remove()
      const style = getComputedStyle(p)
      const box = p.getBoundingClientRect()
      const start = Number((p.dataset.tabs ?? '36 0').split(' ')[1]) || 0
      const rtl = style.direction === 'rtl'
      const textAreaLeft = box.left + parseFloat(style.paddingLeft) + parseFloat(style.borderLeftWidth) - (rtl ? 0 : start * pixelsPerPoint)
      // right-anchored lines (a right-to-left paragraph, a right-aligned absolute tab) are compared from the text area's end: the
      // flow column may be wider than the page's text width (TEXT_WIDTH pt)
      const fromEnd = rtl || !!p.querySelector('[data-ptab^="r"]')
      const contentRight = box.right - parseFloat(style.paddingRight) - parseFloat(style.borderRightWidth)
      const xOf = (node, at) => {
        const range = document.createRange()
        range.setStart(node, at); range.setEnd(node, at + 1)
        const left = range.getBoundingClientRect().left
        return +(fromEnd ? textWidth - (contentRight - left) / pixelsPerPoint : (left - textAreaLeft) / pixelsPerPoint).toFixed(2)
      }
      // after each tab character (laid out in a .tab-space, or a plain one the browser's tab-size places), the first visible character
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT)
      const xs = []
      let waiting = 0
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const text = node.nodeValue ?? ''
        for (let i = 0; i < text.length; i++) {
          if (text[i] === '\t') { waiting++; continue }
          if (waiting && /\S/.test(text[i])) { const x = xOf(node, i); while (waiting) { xs.push(x); waiting-- } }
        }
      }
      const leaders = [...p.querySelectorAll('.tab-space[data-leader]')].map(s => s.dataset.leader)
      const bars = [...p.querySelectorAll(':scope > .tab-bar')].map(b => +((b.getBoundingClientRect().left - textAreaLeft) / pixelsPerPoint).toFixed(2))
      return { text: p.textContent.split('\t').length - 1, xs, leaders, bars, width: +(box.width / pixelsPerPoint).toFixed(1) }
    })
  }, TEXT_WIDTH)
  measured.forEach((m, i) => {
    const expected = reference[i]?.x ?? []
    const pairs = expected.map((x, k) => [x, m.xs[k]])
    const ok = pairs.every(([x, got]) => x === null || (got !== null && got !== undefined && Math.abs(x - got) <= TOLERANCE))
    check(ok, `${path} paragraph ${i + 1} (${reference[i]?.first}): ${pairs.map(([x, got]) => `${x} vs ${got}`).join(', ')}${m.leaders.length ? ' leaders ' + m.leaders.join(',') : ''}`)
  })
  check(measured[0]?.text === 1, `${path}: the text keeps its tab characters`)
  // bar stops (Word draws them; LibreOffice drops them, so the .odt has none)
  const bars = measured.flatMap(m => m.bars)
  if (!path.endsWith('.odt'))
    check(bars.length === referenceBars.length && bars.every((x, i) => Math.abs(x - referenceBars[i]) <= TOLERANCE), `${path}: bar lines at ${bars.join(', ')} vs ${referenceBars.join(', ')}`)
}
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
