// Live check of decimal stops and underlined tabs (the page's fastDocViewTabs pass) against Word's PDF, DOM measurements only:
//   node playwright/tabDecimalUnderlineCheck.mjs <tab-stops-decimal-v1.json> <tab-stops-decimal-v1.docx> [<decimal-char.fodt>]
// The reference is scripts/measure-tab-stops-decimal-pdf.py's output: per decimal paragraph the x (pt from the left margin) of
// each '.' / ',' and of the text's end; per underlined paragraph the x ranges Word underlines. Each must match within 2 pt.
// The .fodt (FastDocView.Tests writes it: tab-stops-decimal-char-v1.fodt) checks a stop's own decimal character (',' at 144 pt).
import { chromium } from 'playwright'
import { readFileSync } from 'node:fs'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const TOLERANCE = 2
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)
const near = (a, b) => a !== null && b !== null && a !== undefined && b !== undefined && Math.abs(a - b) <= TOLERANCE

const browser = await chromium.launch()
const page = await (await browser.newContext({ viewport: { width: 1200, height: 1600 } })).newPage()
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })

async function open(path, paragraphs) {
  step(`open ${path}`)
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), path)
  await page.waitForFunction(n => document.querySelectorAll('.fast-doc-view-block-text p').length >= n && document.querySelectorAll('.fast-doc-view-block-text .tab-space').length > 0, paragraphs, { timeout: 60_000 })
  await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))))
}

// per paragraph: x (pt from the text area's start) of each '.' / ',', of the text's end, and the underlined ranges (text runs that
// draw an underline, and laid-out tabs with data-underline, whose ::after draws it under the tab's whole width)
const measure = () => page.evaluate(() => [...document.querySelectorAll('.fast-doc-view-block-text p')].map(p => {
  const probe = document.createElement('span'); probe.style.cssText = 'display:inline-block;width:100pt'; p.appendChild(probe)
  const ppt = probe.getBoundingClientRect().width / 100; probe.remove()
  const style = getComputedStyle(p), box = p.getBoundingClientRect()
  const left = box.left + parseFloat(style.paddingLeft) + parseFloat(style.borderLeftWidth)
  const x = px => +((px - left) / ppt).toFixed(2)
  const separators = [], ranges = []
  let end = null
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT)
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.nodeValue ?? ''
    for (let i = 0; i < text.length; i++) {
      // a tab the page laid out is measured as its span below; a plain one (default stops: the browser's tab-size) as text
      if (text[i] === ' ' || (text[i] === '\t' && node.parentElement?.classList.contains('tab-space'))) continue
      const range = document.createRange(); range.setStart(node, i); range.setEnd(node, i + 1)
      const r = range.getBoundingClientRect()
      if (text[i] === '.' || text[i] === ',') separators.push(x(r.left))
      if (text[i] !== '\t') end = x(r.right)
      // a character drawn with an underline (text-decoration on an ancestor inside the paragraph)
      for (let e = node.parentElement; e && e !== p.parentElement; e = e.parentElement)
        if (getComputedStyle(e).textDecorationLine.includes('underline')) { ranges.push([x(r.left), x(r.right)]); break }
    }
  }
  for (const span of p.querySelectorAll('.tab-space[data-underline]')) {
    const r = span.getBoundingClientRect()
    const after = getComputedStyle(span, '::after')
    if (after.textDecorationLine.includes('underline') && r.width > 0) ranges.push([x(r.left), x(r.right)])
  }
  // join touching ranges (a gap under 1 pt: characters of one run)
  ranges.sort((a, b) => a[0] - b[0])
  const joined = []
  for (const r of ranges) { const last = joined[joined.length - 1]; if (last && r[0] - last[1] < 1) last[1] = Math.max(last[1], r[1]); else joined.push([...r]) }
  return { text: p.textContent, separators, end, underline: joined, tabsDrawUnderline: [...p.querySelectorAll('.tab-space')].map(s => s.hasAttribute('data-underline')) }
}))

const [referencePath, docx, fodt] = process.argv.slice(2)
const { paragraphs: reference } = JSON.parse(readFileSync(referencePath, 'utf8'))
await open(docx, reference.length)
const measured = await measure()
reference.forEach((expected, i) => {
  const got = measured[i]
  if (!got) { check(false, `paragraph ${i + 1} (${expected.first}) rendered`); return }
  if (expected.separators) {
    const ok = expected.separators.every((s, k) => near(s, got.separators[k])) && near(expected.end, got.end)
    check(ok, `decimal paragraph ${i + 1} (${expected.text}): separators ${expected.separators.join(',')} vs ${got.separators.join(',')}, end ${expected.end} vs ${got.end}`)
  } else {
    const ok = expected.underline.length === got.underline.length && expected.underline.every(([a, b], k) => near(a, got.underline[k][0]) && near(b, got.underline[k][1]))
    check(ok, `underline paragraph ${i + 1} (${expected.text}): ${JSON.stringify(expected.underline)} vs ${JSON.stringify(got.underline)}`)
  }
})
if (fodt) {
  await open(fodt, 3)
  const odt = await measure()
  // Comma (style:char ','): the comma at 144 pt; Point ('.'): the point at 144; None (LibreOffice's ','): the comma at 144
  const at = odt.map(m => m.separators[0])
  check(odt.length === 3 && at.every(x => near(144, x)), `.fodt: each paragraph's own decimal character sits on the 144 pt stop (${at.join(', ')})`)
}
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
