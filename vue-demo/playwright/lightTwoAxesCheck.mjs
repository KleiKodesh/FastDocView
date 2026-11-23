// Live check of Word's light rig on pictures turned on both axes in the app's FastDocView page, DOM only:
//   node playwright/lightTwoAxesCheck.mjs <test-data/word-picture-effects/effects-light-two-v1.docx>
// Every turned picture is drawn with its SVG filter's light transfer (slope 1.043, intercept = Word's lift / 255): the 37
// pictures' lifts match Word's PDF within 1 level, each picture's filter is in the page and applied, no console errors.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [docx] = process.argv.slice(2)
const word = [4.4, 3.4, 13.2, 35.4, 82.4, 80.1, 56.3, 0.4, 6.0, 23.7, 72.7, 69.6, 45.2, 0, 1.7, 9.0, 35.6, 33.2, 20.3,
              1.3, 2.7, 1.6, 0, 0, 0, 10.2, 13.3, 11.6, 2.7, 0.4, 0, 26.2, 29.2, 27.5, 12.3, 5.3, 1.4]
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), docx)
await page.waitForFunction(() => document.querySelectorAll('img[style*="filter:url(#fdv-picture-effect"]').length > 0, null, { timeout: 60_000 })
const lifts = []
for (let pass = 0; pass < 200 && lifts.length < word.length; pass++) {
  const seen = await page.evaluate(() => [...document.querySelectorAll('img[style*="filter:url(#fdv-picture-effect"]')].map(img => {
    const id = /filter:url\(#([^)]+)\)/.exec(img.getAttribute('style'))[1]
    const fn = document.getElementById(id)?.querySelector('feFuncR[slope="1.043"]')
    // the label is the last 'two-...' text before the picture in document order
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT)
    let label = ''
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (img.compareDocumentPosition(n) & Node.DOCUMENT_POSITION_FOLLOWING) break
      if (/^two-/.test(n.textContent.trim())) label = n.textContent.trim()
    }
    return { label, applied: getComputedStyle(img).filter.includes(id), lift: fn ? Number(fn.getAttribute('intercept')) * 255 : null }
  }))
  for (const s of seen) if (s.label.startsWith('two-') && !lifts.some(l => l.label === s.label)) lifts.push(s)
  await page.mouse.move(500, 700); await page.mouse.wheel(0, 400); await page.waitForTimeout(250)
}
check(lifts.length === word.length, `${word.length} turned pictures seen while scrolling (${lifts.length})`)
check(lifts.every(l => l.applied), 'every picture has its filter applied')
const order = ['two-0-0', ...[-45, -30, -15, 15, 30, 45].flatMap(a => [-45, -30, -15, 15, 30, 45].map(b => `two-lat${a}-lon${b}`))]
const off = lifts.filter(l => l.lift === null || Math.abs(l.lift - word[order.indexOf(l.label)]) > 1)
check(off.length === 0, `lifts within 1 level of Word's PDF (${off.length} off: ${JSON.stringify(off.slice(0, 3))})`)
check(errors.length === 0, `no console errors (${errors.slice(0, 3).join(' | ')})`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
