// Live check of equations in the app's FastDocView page: open test-data/word-math/math-v1.* (Word 16) or
// test-data/odt-math/odt-math-v1.odt (LibreOffice), check the MathML is laid out by the browser with a non-zero size and the
// right structure (fractions, roots, scripts, n-ary operators, tables), no preview picture for an equation, no console
// errors. Structure per element and search are checked by the harness (OfficeMathFidelityTest). DOM only, no screenshots.
//   node playwright/mathCheck.mjs <file> [<file>...]
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const files = process.argv.slice(2)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
for (const filePath of files) {
  console.log(`[${new Date().toLocaleTimeString()}] ${filePath}`)
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), filePath)
  await page.waitForFunction(() => document.querySelectorAll('.windowed-scroller-viewport math').length >= 5, null, { timeout: 60_000 })
  const found = await page.evaluate(() => {
    const maths = [...document.querySelectorAll('.windowed-scroller-viewport math')]
    const boxes = maths.map(m => m.getBoundingClientRect())
    const has = tag => maths.some(m => m.querySelector(tag))
    // a stacked fraction: numerator above denominator
    const frac = document.querySelector('.windowed-scroller-viewport mfrac')
    const [num, den] = frac ? [...frac.children].map(c => c.getBoundingClientRect()) : []
    const sup = document.querySelector('.windowed-scroller-viewport msup')
    const [base, script] = sup ? [...sup.children].map(c => c.getBoundingClientRect()) : []
    return {
      count: maths.length,
      sized: boxes.filter(b => b.width > 2 && b.height > 4).length,
      mathDisplay: getComputedStyle(maths[0]).display,
      structure: { mfrac: has('mfrac'), root: has('msqrt') || has('mroot'), msup: has('msup'), nary: has('munderover') || has('msubsup'), table: has('mtable') },
      fractionStacked: !!num && num.bottom <= den.top + 1,
      superscriptRaised: !!base && script.top < base.top,
      previewPictures: [...document.querySelectorAll('.windowed-scroller-viewport img')].filter(i => (i.getAttribute('src') ?? '').includes('ObjectReplacements')).length,
    }
  })
  console.log('  ' + JSON.stringify(found))
  check(found.count >= 5 && found.sized === found.count, `every equation in view is laid out with a size (${found.sized}/${found.count})`)
  check(found.mathDisplay.includes('math'), `the browser lays out MathML (display: ${found.mathDisplay})`)
  check(found.structure.mfrac && found.structure.root && found.structure.msup && found.structure.nary, 'fractions, roots, scripts, n-ary operators present')
  check(found.fractionStacked, 'a fraction: numerator drawn above its denominator')
  check(found.superscriptRaised, 'a superscript: drawn above its base line')
  check(found.previewPictures === 0, 'no preview picture where a formula is drawn')
}
check(errors.length === 0, `no console errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
