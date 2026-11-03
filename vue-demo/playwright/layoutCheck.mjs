// Live layout check of the APP's FastDocView page: list label gaps, the column's margins, tables and pictures,
// measured from the DOM (boxes + computed styles) - no screenshots.
//   node playwright/layoutCheck.mjs <file> [--dump]
// Prints per check PASS / FAIL; --dump prints each rendered block's box summary.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
const dump = process.argv.includes('--dump')
if (!filePath) { console.error('usage: node playwright/layoutCheck.mjs <file> [--dump]'); process.exit(2) }

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 2400 } })
const consoleErrors = []
page.on('pageerror', error => consoleErrors.push(String(error).slice(0, 200)))
step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
step(`open ${filePath}`)
await page.evaluate(path => window.__fastDocViewDemo.open(path, { rebuildIndex: true }), filePath)
await page.waitForFunction(() => window.__fastDocViewTest?.state().isIndexComplete, null, { timeout: 120_000 })
await page.waitForFunction(() => document.querySelectorAll('.fast-doc-view-block').length > 0, null, { timeout: 60_000 })
await page.waitForTimeout(800)

const result = await page.evaluate(() => {
  const blocks = [...document.querySelectorAll('.fast-doc-view-block')]
  const out = { blocks: [], markers: [], tables: [], pictures: [] }
  for (const block of blocks) {
    const rect = block.getBoundingClientRect(), style = getComputedStyle(block)
    out.blocks.push({ text: (block.textContent ?? '').trim().slice(0, 30), left: rect.left, width: rect.width, padStart: style.paddingInlineStart, maxWidth: style.maxWidth })
  }
  for (const marker of document.querySelectorAll('.fast-doc-view-block .list-marker')) {
    const paragraph = marker.closest('p, h1, h2, h3, h4, h5, h6')
    const rtl = getComputedStyle(paragraph).direction === 'rtl'
    const box = marker.getBoundingClientRect()
    // the first text after the label
    const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT)
    let node, textBox = null
    while ((node = walker.nextNode())) {
      if (marker.contains(node) || !node.textContent.trim()) continue
      const range = document.createRange()
      const offset = node.textContent.search(/\S/)
      range.setStart(node, offset); range.setEnd(node, offset + 1)
      textBox = range.getBoundingClientRect(); break
    }
    // the label's own letters (not its box or the tab's room)
    const inner = []
    const labelWalker = document.createTreeWalker(marker, NodeFilter.SHOW_TEXT)
    for (let text; (text = labelWalker.nextNode());) { const r = document.createRange(); r.selectNodeContents(text); inner.push(...r.getClientRects()) }
    const labelEdge = inner.length ? (rtl ? Math.min(...inner.map(r => r.left)) : Math.max(...inner.map(r => r.right))) : (rtl ? box.left : box.right)
    const gap = textBox ? (rtl ? labelEdge - textBox.right : textBox.left - labelEdge) : null
    out.markers.push({ label: marker.textContent, rtl, display: getComputedStyle(marker).display, minWidth: getComputedStyle(marker).minWidth, gap })
  }
  for (const table of document.querySelectorAll('.fast-doc-view-block table')) {
    const block = table.closest('.fast-doc-view-block-text').getBoundingClientRect(), rect = table.getBoundingClientRect()
    const cell = table.querySelector('td'), cs = getComputedStyle(cell)
    out.tables.push({ startGap: rect.left - block.left, endGap: block.right - rect.right, width: rect.width, cellPadding: [cs.paddingTop, cs.paddingInlineStart, cs.paddingInlineEnd].join(' '), borderCollapse: getComputedStyle(table).borderCollapse })
  }
  for (const picture of document.querySelectorAll('.fast-doc-view-block .word-picture')) {
    const block = picture.closest('.fast-doc-view-block-text').getBoundingClientRect(), rect = picture.getBoundingClientRect(), cs = getComputedStyle(picture)
    const image = picture.querySelector('img') ?? (picture.tagName === 'IMG' ? picture : null)
    out.pictures.push({ float: cs.float, display: cs.display, startGap: rect.left - block.left, endGap: block.right - rect.right, width: rect.width, height: rect.height,
      margin: [cs.marginTop, cs.marginRight, cs.marginBottom, cs.marginLeft].join(' '), loaded: image ? image.complete && image.naturalWidth > 0 : false, position: cs.position })
  }
  return out
})
if (dump) console.log(JSON.stringify(result, null, 1))

step('column: the document text width plus the fixed gutter only')
const first = result.blocks[0]
check(first && first.padStart === '12px', `gutter 12px (${first?.padStart})`)
// A4 11906 - 2 * 1440 twips = 9026 twips = 601.7 px at 96 dpi, at the page's text scale; + 2 * 12px gutter
check(first && first.maxWidth !== 'none', `column max-width from the document (${first?.maxWidth})`)

step('list labels: a visible gap after every label with a tab or space suffix')
for (const marker of result.markers)
  check(marker.display === 'inline-block' && (marker.gap ?? 0) >= 2, `label "${marker.label}" ${marker.rtl ? 'RTL' : 'LTR'}: display ${marker.display}, min-width ${marker.minWidth}, gap ${marker.gap?.toFixed(1)}px`)

step('tables')
for (const table of result.tables) console.log(`  table: start gap ${table.startGap.toFixed(1)} end gap ${table.endGap.toFixed(1)} width ${table.width.toFixed(1)} cell padding ${table.cellPadding} ${table.borderCollapse}`)
step('pictures')
for (const picture of result.pictures) console.log(`  picture: ${picture.display} float ${picture.float} ${picture.position} start ${picture.startGap.toFixed(1)} end ${picture.endGap.toFixed(1)} ${picture.width.toFixed(1)}x${picture.height.toFixed(1)} margin ${picture.margin} loaded ${picture.loaded}`)

step('"ignore document margins": the text fills the pane, the gutter stays')
await page.click('.fast-doc-view-ignore-margins')
await page.waitForTimeout(300)
const ignored = await page.evaluate(() => { const s = getComputedStyle(document.querySelector('.fast-doc-view-block')); return [s.maxWidth, s.paddingInlineStart] })
check(ignored[0] === 'none' && ignored[1] === '12px', `max-width ${ignored[0]}, gutter ${ignored[1]}`)
await page.click('.fast-doc-view-ignore-margins')
step('pictures and floats')
check(result.pictures.every(picture => picture.loaded), 'every picture is drawn')

check(consoleErrors.length === 0, `no page errors (${consoleErrors.join(' | ')})`)
await browser.close()
console.log(failures ? `${failures} FAILED` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
