// Live check of OpenDocument shape text in the app's FastDocView page (fastDocViewShapeText.ts): open the --fidelity fixture
// (TestOutput/fidelity-v1-odt-shape-text.fodt, OdtShapeTextFidelityTest), check that shapes which grow with their text took
// the text's height / width plus their distances (LibreOffice svdoashp.cxx AdjustTextFrameWidthAndHeight), grew in the right
// direction, that a fixed shape kept its size, and that dimension-line labels sit where the measured label puts them
// (outside on the left when wider than the line, centred above otherwise). DOM only, no screenshots.
//   node playwright/shapeTextCheck.mjs <path to fidelity-v1-odt-shape-text.fodt>
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 1400 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), filePath)
await page.waitForFunction(() => document.querySelectorAll('.word-shape[data-autogrow-done]').length >= 4 && document.querySelectorAll('.fast-doc-view-measure[data-measure-done]').length >= 2, null, { timeout: 60_000 })

const result = await page.evaluate(() => {
  const frames = [...document.querySelectorAll('.word-shape')]
  const ptPerPx = frame => Number.parseFloat(frame.style.width) / frame.offsetWidth
  const shape = i => {
    const frame = frames[i], box = frame.querySelector(':scope > .word-shape-text'), content = box?.firstElementChild
    const k = ptPerPx(frame)
    return {
      width: Number.parseFloat(frame.style.width), height: Number.parseFloat(frame.style.height), translate: frame.style.translate,
      contentHeight: (content?.offsetHeight ?? 0) * k, contentWidth: (content?.offsetWidth ?? 0) * k,
      boxHeight: Number.parseFloat(box?.style.height ?? '0'), overflows: box ? box.scrollHeight > box.clientHeight + 1 : true,
      stretched: frame.querySelector('svg')?.getAttribute('preserveAspectRatio') === 'none',
    }
  }
  const measure = i => {
    const m = document.querySelectorAll('.fast-doc-view-measure')[i]
    const label = m.querySelector(':scope > .word-shape-text').getBoundingClientRect()
    const lines = m.querySelector('svg path').getAttribute('d')
    const helps = m.parentElement.querySelector(':scope > svg path').getBoundingClientRect()
    return { labelLeft: label.left, labelRight: label.right, labelBottom: label.bottom, helpLeft: helps.left, helpRight: helps.right, lines }
  }
  return { grow: shape(0), up: shape(1), wide: shape(2), shrink: shape(3), fixed: shape(4), long: measure(0), short: measure(1) }
})
console.log(JSON.stringify(result))
const near = (a, b, tolerance = 1) => Math.abs(a - b) <= tolerance
const { grow, up, wide, shrink, fixed, long, short } = result
const shift = t => { const [x = 0, y = 0] = (t || '0pt 0pt').split(' ').map(Number.parseFloat); return { x, y: Number.isNaN(y) ? 0 : y } }
check(grow.height > 36 && near(grow.height, grow.contentHeight + 7.2) && !grow.translate && grow.stretched && !grow.overflows,
  `grows down: height = text + distances (${grow.height.toFixed(1)} vs ${(grow.contentHeight + 7.2).toFixed(1)} pt), not moved, outline stretched`)
check(near(up.height, grow.height) && near(shift(up.translate).y, -(up.height - 36), 0.1), `grows up: same height, moved up by the growth (${up.translate})`)
check(wide.width > 108 && near(wide.width, wide.contentWidth + 14.4) && shift(wide.translate).x === 0, `no wrap: grows right to the text width (${wide.width.toFixed(1)} pt)`)
check(near(shrink.height, 36 + 7.2) && near(shift(shrink.translate).y, (144 - 43.2) / 2, 0.1), `shrinks to the minimum 0.5 in + distances, both ways (${shrink.height} pt, ${shrink.translate})`)
check(fixed.height === 36 && !fixed.stretched, 'auto-grow off: keeps 0.5 in')
check(long.labelRight <= long.helpLeft + 1, `long label: outside, left of the line (label right ${long.labelRight.toFixed(1)}, line start ${long.helpLeft.toFixed(1)})`)
check(short.labelLeft > short.helpLeft && short.labelRight < short.helpRight && near((short.labelLeft + short.labelRight) / 2, (short.helpLeft + short.helpRight) / 2, 1.5),
  'short label: centred between the help lines')
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
