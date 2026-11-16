// Live check of vertical cell text, RTF border groups and ODT drop caps over spans in the app's FastDocView page (DOM only):
//   node playwright/bordersCellsCheck.mjs <test-data/borders-cell-direction/borders-cell-direction-v1.rtf> <.../drop-cap-spans-v1.fodt>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const [rtf, fodt] = process.argv.slice(2)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })

await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), rtf)
await page.waitForFunction(() => [...document.querySelectorAll('p')].some(p => p.textContent.startsWith('Alone')), null, { timeout: 60_000 })
const rtfResult = await page.evaluate(() => {
  const cell = text => [...document.querySelectorAll('td')].find(td => td.textContent.trim() === text)
  const para = start => [...document.querySelectorAll('p')].find(p => p.textContent.startsWith(start))
  const box = p => p.getBoundingClientRect()
  const one = para('Boxed one'), two = para('Boxed two')
  const s1 = getComputedStyle(one), s2 = getComputedStyle(two)
  return {
    down: getComputedStyle(cell('Down')).writingMode, up: getComputedStyle(cell('Up')).writingMode, flat: getComputedStyle(cell('Flat')).writingMode,
    downTall: cell('Down').querySelector('*')?.getBoundingClientRect().height ?? 0,
    gap: box(two).top - box(one).bottom, oneBottom: s1.borderBottomStyle, twoTop: s2.borderTopStyle, shadow: s2.boxShadow,
  }
})
console.log(JSON.stringify(rtfResult))
check(rtfResult.down === 'vertical-rl' && rtfResult.up === 'sideways-lr' && rtfResult.flat === 'horizontal-tb', 'rtf: cell writing modes drawn')
check(Math.abs(rtfResult.gap) < 0.5 && rtfResult.oneBottom === 'none' && rtfResult.twoTop === 'none', 'rtf: the two boxed paragraphs touch with no line between (one box)')
check(rtfResult.shadow !== 'none', 'rtf: the box has its shadow')

await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), fodt)
await page.waitForFunction(() => document.querySelectorAll('.drop-cap').length >= 3, null, { timeout: 60_000 })
const caps = await page.evaluate(() => [...document.querySelectorAll('.drop-cap')].map(cap => ({ text: cap.textContent, height: cap.getBoundingClientRect().height })))
console.log(JSON.stringify(caps))
check(caps[0].text === 'Abc' && caps[1].text === 'Word' && caps[2].text === 'X', 'fodt: drop caps over spans: Abc, Word, X')
check(caps.every(cap => cap.height > 25), 'fodt: every initial is drawn large')
check(errors.length === 0, `no console errors (${errors.join(' | ')})`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
