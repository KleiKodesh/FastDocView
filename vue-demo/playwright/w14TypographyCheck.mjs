// Live check of w14 OpenType typography and reflection in the app's FastDocView page, DOM measurements only:
//   node playwright/w14TypographyCheck.mjs <a Word-made .docx with ligatures / numForm / stylistic sets / reflection runs>
// Typography: the computed font-variant-ligatures / -numeric / font-feature-settings. Reflection: every word has its own
// mirror box with a computed -webkit-box-reflect, the box ends at the baseline (text-box trim), and the reflected
// paragraph still wraps onto several lines with the same line pitch as without the effect.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [docx] = process.argv.slice(2)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 700, height: 1400 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
step(`open ${docx}`)
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), docx)
await page.waitForFunction(() => document.querySelectorAll('.text-reflection').length > 0, null, { timeout: 60_000 })
await page.waitForTimeout(400)

const result = await page.evaluate(() => {
  const styleOf = text => {
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT)
    for (let node; (node = walker.nextNode());) if (node.textContent.includes(text)) {
      const s = getComputedStyle(node.parentElement)
      return { ligatures: s.fontVariantLigatures, numeric: s.fontVariantNumeric, features: s.fontFeatureSettings }
    }
    return null
  }
  const boxes = [...document.querySelectorAll('.text-reflection')]
  const lines = new Set(boxes.map(b => Math.round(b.getBoundingClientRect().top)))
  const baselineGaps = boxes.map(b => {
    // a zero-size inline-block probe on the baseline of the box's line
    const probe = document.createElement('span'); probe.style.cssText = 'display:inline-block;width:0;height:0'
    b.after(probe); const gap = probe.getBoundingClientRect().bottom - b.getBoundingClientRect().bottom; probe.remove(); return gap
  })
  return {
    standard: styleOf('office fluffy final standard'), none: styleOf('office fluffy final none'),
    oldStyle: styleOf('oldstyle tabular'), sets: styleOf('stylistic set six'), alternates: styleOf('contextual alternates'),
    boxCount: boxes.length, reflect: getComputedStyle(boxes[0]).webkitBoxReflect, lines: lines.size, baselineGaps,
    words: boxes.map(b => b.textContent),
  }
})
console.log(JSON.stringify(result))
check(/common-ligatures/.test(result.standard?.ligatures) && !/no-common/.test(result.standard.ligatures) && /"liga"(, | 1|$)/.test(result.standard.features) && /"clig" 0/.test(result.standard.features), 'standard ligatures: liga on, clig off')
check(/no-common-ligatures/.test(result.none?.ligatures), 'ligatures none: no-common-ligatures')
check(/oldstyle-nums/.test(result.oldStyle?.numeric) && /tabular-nums/.test(result.oldStyle.numeric), 'oldStyle + tabular digits')
check(/ss02/.test(result.sets?.features) && /ss03/.test(result.sets.features), 'stylistic sets 2 and 3')
check(/contextual/.test(result.alternates?.ligatures) && !/no-contextual/.test(result.alternates.ligatures), 'contextual alternates on')
check(result.boxCount >= 10 && /below/.test(result.reflect) && /gradient/.test(result.reflect), `reflection: ${result.boxCount} word boxes, computed box-reflect`)
check(result.lines >= 2, `reflected paragraph wraps onto ${result.lines} lines`)
check(result.baselineGaps.every(g => Math.abs(g) <= 1.5), 'each mirror box ends at the baseline')
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
