// Live check of table STRUCTURE in the app's FastDocView page, computed styles and DOM geometry only:
//   node playwright/tablesCheck.mjs <test-data/word-tables/tables-v1.rtf> [<test-data/word-tables/tables-v1.doc>]
// The Word 16 fixture (make-tables-fixtures.ps1): C2:C3 merged down is drawn as one joined area (the first cell has no bottom
// line, the cell below no top line, the two touch), tables nested 3 deep sit inside their cells, a cell's own borders (2.25 pt
// red top, double blue box, no line elsewhere) and vertical alignment (top / middle / bottom in a 40 pt row).
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1100, height: 3000 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })

for (const path of process.argv.slice(2)) {
  step(`open ${path}`)
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), path)
  const cellText = ['m23', 'deep1', 'red-top', 'blue-box', 'top', 'center', 'bottom']
  await page.waitForFunction(texts => {
    const cells = [...document.querySelectorAll('.windowed-scroller-viewport td')]
    return texts.every(text => cells.some(td => !td.querySelector('td') && (td.textContent ?? '').trim().startsWith(text)))
  }, cellText, { timeout: 60_000 })
  await page.waitForTimeout(400)
  const r = await page.evaluate(() => {
    const cells = [...document.querySelectorAll('.windowed-scroller-viewport td')].filter(td => !td.querySelector('td'))
    const cell = text => cells.find(td => (td.textContent ?? '').trim() === text) ?? cells.find(td => (td.textContent ?? '').trim().startsWith(text))
    const side = (td, name) => {
      const s = getComputedStyle(td)
      return { style: s[`border${name}Style`], width: parseFloat(s[`border${name}Width`]), color: s[`border${name}Color`] }
    }
    const merged = cell('m23'), below = document.querySelector('.windowed-scroller-viewport td.merged-up')
    const mergedBox = merged.getBoundingClientRect(), belowBox = below?.getBoundingClientRect()
    const deep = cell('deep1')
    let depth = 0
    for (let node = deep; node; node = node.parentElement) if (node.tagName === 'TABLE') depth++
    const inner21 = cell('inner21')
    const innerTable = inner21.closest('table'), deepTable = deep.closest('table')
    const textTop = td => { const range = document.createRange(); range.selectNodeContents(td); const rects = [...range.getClientRects()].filter(x => x.height > 0); return rects.length ? rects[0].top - td.getBoundingClientRect().top : -1 }
    return {
      mergedBottom: side(merged, 'Bottom').style, belowTop: below ? side(below, 'Top').style : null,
      gap: belowBox ? Math.abs(belowBox.top - mergedBox.bottom) : 999, sameColumn: belowBox ? Math.abs(belowBox.left - mergedBox.left) < 1.5 && Math.abs(belowBox.width - mergedBox.width) < 1.5 : false,
      belowText: below ? (below.textContent ?? '').trim() : null, mergedHolds: (merged.textContent ?? '').includes('m33'),
      depth, innerRows: innerTable.rows.length, deepInsideInner: innerTable.contains(deepTable) && innerTable !== deepTable,
      innerInsideOuterCell: !!innerTable.parentElement.closest('td'),
      redTop: side(cell('red-top'), 'Top'), redLeft: side(cell('red-top'), 'Left'),
      blue: ['Top', 'Bottom', 'Left', 'Right'].map(name => side(cell('blue-box'), name)),
      align: ['top', 'center', 'bottom'].map(text => ({ vertical: getComputedStyle(cell(text)).verticalAlign, top: textTop(cell(text)) })),
    }
  })
  console.log('   ', JSON.stringify(r))
  check(r.mergedBottom === 'none' && r.belowTop === 'none' && r.gap < 1.5 && r.sameColumn, `${path}: the cell merged down and the one below have no line between them and touch (gap ${r.gap?.toFixed?.(2)})`)
  check(r.mergedHolds && r.belowText === '', `${path}: the merged area's text is in its first cell, the continued cell is empty`)
  check(r.depth === 3 && r.innerRows === 2 && r.deepInsideInner && r.innerInsideOuterCell, `${path}: nesting depth 3 (deep1 in ${r.depth} tables), the inner table has its 2 rows inside the outer cell`)
  check(r.redTop.style === 'solid' && r.redTop.color === 'rgb(255, 0, 0)' && r.redTop.width >= 2.5 && r.redLeft.style === 'none', `${path}: the cell's own red 2.25 pt top border, no left line`)
  check(r.blue.every(b => b.style === 'double' && b.color === 'rgb(0, 0, 255)'), `${path}: the double blue box on all four sides`)
  check(r.align[0].vertical !== 'middle' && r.align[0].vertical !== 'bottom' && r.align[1].vertical === 'middle' && r.align[2].vertical === 'bottom'
        && r.align[0].top < r.align[1].top - 5 && r.align[1].top < r.align[2].top - 5, `${path}: vertical alignment top / middle / bottom in the 40 pt row (text at ${r.align.map(a => a.top.toFixed(1)).join(' / ')} px from the cell top)`)
}
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
