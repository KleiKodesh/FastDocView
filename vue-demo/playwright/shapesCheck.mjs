// Live check of .odt / .fodt text boxes and shapes in the APP's FastDocView page, measured from the DOM (boxes + computed
// styles) - no screenshots. Made for TestOutput/fidelity-v1-odt-shapes.fodt (and its LibreOffice .odt).
//   node playwright/shapesCheck.mjs <file>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
if (!filePath) { console.error('usage: node playwright/shapesCheck.mjs <file>'); process.exit(2) }

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

step('measure boxes')
const r = await page.evaluate(() => {
  const find = text => [...document.querySelectorAll('.text-box')].find(box => box.textContent.includes(text))
  const shapeWith = selector => [...document.querySelectorAll('.word-shape')].find(box => box.querySelector(selector))
  const box = (element) => {
    if (!element) return null
    const rect = element.getBoundingClientRect(), style = getComputedStyle(element)
    const paragraph = element.closest('p, h1, h2, h3, h4, h5, h6')
    const prect = paragraph?.getBoundingClientRect()
    return { left: rect.left, right: rect.right, top: rect.top, width: rect.width, height: rect.height, float: style.cssFloat, display: style.display,
             position: style.position, pLeft: prect?.left, pRight: prect?.right, pTop: prect?.top, pBottom: prect?.bottom,
             // the anchoring paragraph's own text node after the box
             textLeft: (() => { const n = [...(paragraph?.childNodes ?? [])].find(node => node.nodeType === 3 && node.textContent.trim()); if (!n) return null; const range = document.createRange(); range.selectNodeContents(n); return range.getBoundingClientRect().left })() }
  }
  const svgFill = element => { const svg = element?.querySelector('svg'); if (!svg) return null; const rect = svg.getBoundingClientRect(), outer = element.getBoundingClientRect(); return { dw: rect.width - outer.width, dh: rect.height - outer.height } }
  const group = shapeWith('.word-shape-part')
  const parts = group ? [...group.querySelectorAll(':scope > .word-shape-part')].map(part => { const a = part.getBoundingClientRect(), g = group.getBoundingClientRect(); return { dx: a.left - g.left, dy: a.top - g.top, w: a.width, h: a.height } }) : []
  return {
    boxLeft: box(find('boxtext one')), inline: box(find('inlinebox')), centre: box(find('centrebox')), rtl: box(find('rtlbox')),
    shapeText: box(shapeWith('.word-shape-text')), shapeTextSvg: svgFill(shapeWith('.word-shape-text')),
    ellipse: box(shapeWith('ellipse')), line: box(shapeWith('line')), group: box(group), parts,
    paragraphsInBox: find('boxtext one')?.querySelectorAll('.text-box-paragraph').length ?? 0,
    pInsideP: document.querySelectorAll('p p, p h1, p div').length,
  }
})
// the page zooms its text: one scale for all sizes, taken from the as-char box (fixed 72pt width in the file)
const scale = r.inline ? r.inline.width / (72 * 96 / 72) : 1
console.log(`  page zoom ${scale.toFixed(3)}`)
const px = pt => pt * 96 / 72 * scale
const near = (a, b, tolerance = Math.max(1.5, Math.abs(b) * 0.03)) => a !== null && a !== undefined && Math.abs(a - b) <= tolerance
check(r.boxLeft && r.boxLeft.float === 'left' && near(r.boxLeft.left, r.boxLeft.pLeft) && near(r.boxLeft.width, px(108)), `text box floats left at the paragraph's left, 108pt wide (${JSON.stringify(r.boxLeft)})`)
check(r.boxLeft && r.boxLeft.height >= px(36) - 1, 'text box at least its min-height')
check(r.boxLeft && r.boxLeft.textLeft !== null && r.boxLeft.textLeft >= r.boxLeft.right + px(18) - 1.5, 'the paragraph text runs right of the box after its 18pt gap')
check(r.paragraphsInBox === 2 && r.pInsideP === 0, `box paragraphs as block spans, no nested <p> (${r.paragraphsInBox}, ${r.pInsideP})`)
check(r.inline && r.inline.display === 'inline-block' && r.inline.float === 'none' && r.inline.top >= r.inline.pTop - 1 && r.inline.top < r.inline.pBottom, `as-char box in the line (${JSON.stringify(r.inline)})`)
check(r.centre && r.centre.float === 'none' && near((r.centre.left + r.centre.right) / 2, (r.centre.pLeft + r.centre.pRight) / 2, 2), `centred both-sides box: a centred block (${JSON.stringify(r.centre)})`)
check(r.rtl && (r.rtl.float === 'right' || r.rtl.float === 'left') && near(r.rtl.float === 'right' ? r.rtl.right : r.rtl.left, r.rtl.float === 'right' ? r.rtl.pRight : r.rtl.pLeft), `right-to-left box at its side (${JSON.stringify(r.rtl)})`)
// size: the hand-written file says 72 x 36pt; LibreOffice re-saves this shape smaller (0.66in), so only its SVG filling the box is checked there
check(r.shapeText && (process.argv[2].endsWith('.fodt') ? near(r.shapeText.width, px(72)) && r.shapeText.height >= px(36) - 1 : r.shapeText.width > 0) && r.shapeTextSvg && near(r.shapeTextSvg.dw, 0) && near(r.shapeTextSvg.dh, 0),
      `custom shape: 72 x 36pt box, the SVG fills it (${JSON.stringify(r.shapeText)} ${JSON.stringify(r.shapeTextSvg)})`)
check(r.ellipse && near(r.ellipse.width, px(72)), 'ellipse box 72pt')
check(r.line && near(r.line.width, px(144)) && near(r.line.height, px(36)), `line box 144 x 36pt (${JSON.stringify(r.line)})`)
check(r.group && near(r.group.width, px(108)) && r.parts.length === 2 && near(r.parts[1].dx, px(36)) && near(r.parts[1].dy, px(18)),
      `group: 108pt wide, second child at 36 / 18pt (${JSON.stringify(r.group)} ${JSON.stringify(r.parts)})`)
check(consoleErrors.length === 0, `no page errors (${consoleErrors.join(' | ')})`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
