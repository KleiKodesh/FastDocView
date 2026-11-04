// Live check of OpenDocument shape text in the app's FastDocView page (fastDocViewShapeText.ts): open the --fidelity fixture
// (TestOutput/fidelity-v2-odt-shape-text.fodt, OdtShapeTextFidelityTest) and check that shapes which grow with their text took
// the text's size plus their distances (LibreOffice svdoashp.cxx AdjustTextFrameWidthAndHeight) in the right direction - plain,
// turned, mirrored, in a group, with turned text, without the attribute (default on) - that the frame took the new bounds and a
// contour wrap the new outline, and that dimension-line labels and arrows sit where the measured label puts them, inside the
// frame. DOM only, no screenshots.
//   node playwright/shapeTextCheck.mjs <path to fidelity-v2-odt-shape-text.fodt>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 4000 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), filePath)
await page.waitForFunction(() => document.querySelectorAll('.word-shape[data-shape-page-done]').length >= 13, null, { timeout: 60_000 })

const r = await page.evaluate(() => {
  const frames = [...document.querySelectorAll('.word-shape')]
  const contained = (frame, element) => {
    const f = frame.getBoundingClientRect(), e = element.getBoundingClientRect()
    return e.left >= f.left - 1.5 && e.right <= f.right + 1.5 && e.top >= f.top - 1.5 && e.bottom <= f.bottom + 1.5
  }
  const shape = i => {
    const frame = frames[i], boxes = [...frame.querySelectorAll('.word-shape-text[data-autogrow-id]')], box = boxes[boxes.length - 1]
    const content = box?.firstElementChild
    const g = box ? frame.querySelector(`g[data-autogrow-id="${box.dataset.autogrowId}"]`) : null
    return {
      width: Number.parseFloat(frame.style.width), height: Number.parseFloat(frame.style.height), marginTop: frame.style.marginTop,
      contentHeight: content ? content.offsetHeight * (Number.parseFloat(frame.style.width) / frame.offsetWidth) : 0,
      overflows: box ? box.scrollHeight > box.clientHeight + 1 || box.scrollWidth > box.clientWidth + 1 : true,
      textInside: box ? contained(frame, box) : false,
      matrix: g?.getAttribute('transform') ?? '', shapeOutside: getComputedStyle(frame).shapeOutside,
    }
  }
  const measure = i => {
    const frame = frames[i], m = frame.querySelector('.fast-doc-view-measure')
    const label = m.querySelector(':scope > .word-shape-text').getBoundingClientRect()
    const helps = frame.querySelector(':scope > svg g').getBoundingClientRect()
    const line = m.querySelector('svg g').getBoundingClientRect()
    const f = frame.getBoundingClientRect()
    return { labelLeft: label.left, labelRight: label.right, helpLeft: helps.left, helpRight: helps.right, lineInside: line.left >= f.left - 1.5 && line.right <= f.right + 1.5,
             markers: m.querySelectorAll('svg path[stroke="none"]').length, helpMarkers: frame.querySelectorAll(':scope > svg path[stroke="none"]').length }
  }
  return {
    grow: shape(0), up: shape(1), wide: shape(2), shrink: shape(3), fixed: shape(4), long: measure(5), short: measure(6),
    turned: shape(7), mirrored: shape(8), group: shape(9), textTurned: shape(10), noAttribute: shape(11), contour: shape(12), arrows: measure(13),
  }
})
console.log(JSON.stringify(r))
const near = (a, b, tolerance = 1) => Math.abs(a - b) <= tolerance
// a margin as the page writes it (calc of px and pt) in points
const marginPt = text => [...text.matchAll(/(-?[\d.]+)(px|pt)/g)].reduce((sum, m) => sum + Number(m[1]) * (m[2] === 'px' ? 0.75 : 1), 0)
check(r.grow.height > 36 && near(r.grow.height, r.grow.contentHeight + 7.2) && !r.grow.overflows && r.grow.textInside, `grows down: height = text + distances (${r.grow.height.toFixed(1)} pt)`)
check(near(r.up.height, r.grow.height) && near(marginPt(r.up.marginTop), -(r.up.height - 36), 0.1), `grows up: same height, moved up by the growth (${r.up.marginTop})`)
check(r.wide.width > 108 && !r.wide.overflows, `no wrap: grows right to the text (${r.wide.width.toFixed(1)} pt)`)
check(near(r.shrink.height, 43.2) && near(marginPt(r.shrink.marginTop), 50.4, 0.1), `shrinks to the minimum + distances, both ways (${r.shrink.height} pt, ${r.shrink.marginTop})`)
check(r.fixed.height === 36, 'auto-grow off: keeps 0.5 in')
check(r.long.labelRight <= r.long.helpLeft + 1 && r.long.lineInside, 'long label: outside on the left; the frame holds the whole dimension line')
check(r.short.labelLeft > r.short.helpLeft && r.short.labelRight < r.short.helpRight && near((r.short.labelLeft + r.short.labelRight) / 2, (r.short.helpLeft + r.short.helpRight) / 2, 1.5), 'short label: centred between the help lines')
check(near(r.turned.width, r.grow.height, 1.5) && near(r.turned.height, 108, 0.5) && !r.turned.overflows && r.turned.textInside,
  `turned 90: the growth runs across the page (frame ${r.turned.width.toFixed(1)} x ${r.turned.height.toFixed(1)} pt)`)
check(near(r.mirrored.height, r.grow.height, 0.5) && r.mirrored.matrix.startsWith('matrix(-1') && r.mirrored.textInside, `mirrored: grows the same, still mirrored (${r.mirrored.matrix})`)
check(near(r.group.height, 36 + r.grow.height, 1) && !r.group.overflows && r.group.textInside, `group: the frame takes the grown member (${r.group.height.toFixed(1)} pt)`)
// LibreOffice lays the text in the unturned text area, then turns it (viewcontactofsdrobjcustomshape.cxx:155-190): it wraps at the narrow side, the shape grows along the long one
check(r.textTurned.height > 108 && near(r.textTurned.width, 36, 0.5) && !r.textTurned.overflows, `turned text: wraps in its unturned area, grows in height (${r.textTurned.height.toFixed(1)} pt)`)
check(r.noAttribute.height > 36 && near(r.noAttribute.height, r.noAttribute.contentHeight, 1), `auto-grow absent: grows, LibreOffice's default on (${r.noAttribute.height.toFixed(1)} pt)`)
const ys = [...r.contour.shapeOutside.matchAll(/(-?[\d.]+)(pt|px) (-?[\d.]+)(pt|px)/g)].map(m => Number(m[3]) * (m[4] === 'px' ? 0.75 : 1))
check(r.contour.shapeOutside.startsWith('polygon') && near(Math.max(...ys), r.contour.height, 1.5), `contour wrap follows the grown outline (to ${Math.max(...ys).toFixed(1)} of ${r.contour.height.toFixed(1)} pt)`)
check(r.arrows.markers === 2 && r.arrows.helpMarkers === 0 && r.arrows.lineInside, `arrows: two markers on the dimension line, none on the help lines (${r.arrows.markers}, ${r.arrows.helpMarkers})`)
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
