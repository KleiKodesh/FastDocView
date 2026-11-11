// Live check of display equations that break (the app's fastDocViewMath.ts on the library's breakable equations): opens
// test-data/word-math/math-settings-v2.docx (m:brkBin repeat, m:brkBinSub "-+", m:wrapRight) in a narrow window so the
// 40-term equation wraps, then checks every continuation line starts with its repeated operator (only those are shown) and
// ends at the block's right content edge, also after the window is widened. DOM only, no screenshots.
//   node playwright/mathBreaksCheck.mjs <path to math-settings-v2.docx>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 700, height: 900 } })
const errors = []
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), filePath)
await page.waitForFunction(() => window.__fastDocViewTest?.state()?.isIndexComplete && window.__fastDocViewTest.state().blockCount > 0, null, { timeout: 60_000 })

async function measure() {
  const state = await page.evaluate(() => window.__fastDocViewTest.state())
  for (let i = 0; i < state.blockCount; i++) {
    await page.evaluate(index => window.__fastDocViewTest.scrollToBlock(index), i)
    await page.waitForTimeout(80)
    const found = await page.evaluate(index => {
      const row = document.querySelector(`.windowed-scroller-viewport [data-index="${index}"]`)
      const block = [...(row?.querySelectorAll('.omath-para[data-math-break]') ?? [])].find(b => b.querySelectorAll(':scope > math').length >= 30)
      if (!block) return null
      const pieces = [...block.querySelectorAll(':scope > math')]
      const lines = []
      for (const piece of pieces) {
        const box = piece.getBoundingClientRect()
        const line = lines[lines.length - 1]
        if (line && box.top < line.bottom - 1) { line.right = box.right; line.count++; line.bottom = Math.max(line.bottom, box.bottom) }   // overlapping = the same line (shared baseline)
        else lines.push({ top: box.top, bottom: box.bottom, right: box.right, count: 1, startsWithRepeat: getComputedStyle(piece.querySelector('mo.omath-repeat') ?? piece).display !== 'none' && !!piece.querySelector('mo.omath-repeat') })
      }
      const shown = block.querySelectorAll('mo.omath-repeat:not([style*="none"])').length
      const right = block.getBoundingClientRect().right - parseFloat(getComputedStyle(block).paddingRight || '0')
      return { lines: lines.length, shown, continuationRepeat: lines.slice(1).every(l => l.startsWithRepeat), rightGaps: lines.slice(1).map(l => +(right - l.right).toFixed(1)) }
    }, i)
    if (found) return found
  }
  return null
}

let found = await measure()
console.log('  narrow', JSON.stringify(found))
check(!!found && found.lines >= 2, 'the long display equation wraps into several lines')
check(!!found && found.shown === found.lines - 1 && found.continuationRepeat, 'brkBin repeat: each continuation line starts with the repeated operator, no other copy shown')
check(!!found && found.rightGaps.every(g => Math.abs(g) <= 1.5), 'wrapRight: each continuation line ends at the right edge (within 1.5 px)')
await page.setViewportSize({ width: 1000, height: 900 })
await page.waitForTimeout(500)
found = await measure()
console.log('  wider', JSON.stringify(found))
check(!!found && found.shown === found.lines - 1 && found.rightGaps.every(g => Math.abs(g) <= 1.5), 'after a resize: repeats and right edges follow the new lines')
const timings = await page.evaluate(() => window.__fastDocViewMathMs ?? [])
const operators = await page.evaluate(() => Math.max(0, ...[...document.querySelectorAll('.omath-para[data-math-break]')].map(b => b.querySelectorAll('mo.omath-repeat').length)))
console.log(`  layout: ${timings.length} runs, max ${Math.max(0, ...timings).toFixed(2)} ms, last ${(timings[timings.length - 1] ?? 0).toFixed(2)} ms; ${operators} repeated operators in the block`)
check(operators >= 50 && timings.length > 0 && Math.max(...timings) < 16, 'batched measurement (at most three reads per render, never one per operator): a block of 50 repeated operators lays out within a frame (16 ms)')
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
