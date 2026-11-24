// Live check of EQ field layout in the APP's FastDocView page through the demo (Word 16's drawing, test-data/word-fields/eq-layout-v1.pdf):
// \o centres its elements by default and aligns them left / right with \al / \ar, on one baseline, as wide as the widest; \i limits
// above and below; \b with only \lc one bracket. DOM + console + the page's DEV hook - no screenshots.
//   node playwright/eqLayoutCheck.mjs <eq-layout-v1.docx / .doc / .rtf> [...]
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const files = process.argv.slice(2)
if (files.length === 0) { console.error('usage: node playwright/eqLayoutCheck.mjs <file> [<file> ...]'); process.exit(2) }

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const consoleErrors = []
page.on('console', message => { if (message.type() === 'error') consoleErrors.push(message.text().slice(0, 200)) })
page.on('pageerror', error => consoleErrors.push(String(error).slice(0, 200)))
const waitFor = (predicate, argument, timeout = 60_000) => page.waitForFunction(predicate, argument, { timeout })

await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await waitFor(() => !!window.__fastDocViewDemo, null, 120_000)
for (const file of files) {
  console.log(`[${new Date().toLocaleTimeString()}] ${file}`)
  await page.evaluate(path => window.__fastDocViewDemo.open(path), file)
  await waitFor(() => window.__fastDocViewTest?.state().isIndexComplete, null, 120_000)
  await page.evaluate(() => window.__fastDocViewTest.scrollToBlock(0))
  await page.waitForTimeout(800)
  const cases = await page.evaluate(() => {
    const result = {}
    for (const block of document.querySelectorAll('.windowed-scroller-viewport [data-block-index], .windowed-scroller-viewport p, .windowed-scroller-viewport div')) {
      const name = /case ([a-z0-9]+):/.exec(block.textContent ?? '')?.[1]
      const math = block.querySelector('math')
      if (!name || !math || result[name]) continue
      const box = element => { const r = element.getBoundingClientRect(); return { left: r.left, right: r.right, bottom: r.bottom, width: r.width } }
      const grid = math.querySelector('.eq-overstrike')
      result[name] = {
        cells: grid ? [...grid.children].map(box) : [],
        grid: grid ? box(grid) : null,
        brackets: [...math.querySelectorAll('mo[stretchy="true"]')].map(o => o.textContent).join(''),
        under: !!math.querySelector('munderover'),
        display: grid ? getComputedStyle(grid).display : '',
      }
    }
    return result
  })
  const centre = c => (c.left + c.right) / 2
  const o = cases.o, oal = cases.oal, oar = cases.oar, oi = cases.oi
  check(/grid/.test(o?.display ?? ''), `\\o drawn as one grid cell, blockified by the math row (display ${o?.display})`)
  check(o && Math.abs(centre(o.cells[0]) - centre(o.cells[1])) < 1, '\\o: W and i centred on each other')
  check(oal && Math.abs(oal.cells[0].left - oal.cells[1].left) < 1, '\\o\\al: one left edge')
  check(oar && Math.abs(oar.cells[0].right - oar.cells[1].right) < 1, '\\o\\ar: one right edge')
  check(oi && Math.abs(oi.grid.width - Math.max(...oi.cells.map(c => c.width))) < 1, '\\o(i,W): as wide as the widest element')
  check(o && Math.abs(o.cells[0].bottom - o.cells[1].bottom) < 2, '\\o: elements on one baseline')
  check(cases.int?.under && cases.sum?.under && !cases.intin?.under, 'integral and sum limits above and below; \\in to the right')
  check(cases.blc?.brackets === '[' && cases.brc?.brackets === ']' && cases.bfrac?.brackets === '()', `one-sided brackets ([${cases.blc?.brackets}] [${cases.brc?.brackets}] [${cases.bfrac?.brackets}])`)
}
check(consoleErrors.length === 0, `no console errors${consoleErrors.length ? ': ' + consoleErrors.join(' | ') : ''}`)
await browser.close()
console.log(failures === 0 ? 'EQ LAYOUT LIVE CHECK PASSED' : `${failures} FAILED`)
process.exit(failures === 0 ? 0 : 1)
