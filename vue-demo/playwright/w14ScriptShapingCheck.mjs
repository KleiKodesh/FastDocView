// Live check of Word's shaping per script in the app's FastDocView page, computed styles only:
//   node playwright/w14ScriptShapingCheck.mjs <test-data/word-text-effects/w14-scripts-v1.docx> <a .doc or .rtf file>
// Hebrew text: no common ligatures, no contextual alternates (also where the run sets ligatures all + cntxtAlts, as Word
// ignores it); Arabic text: the browser's ligatures (Word always forms them); a .doc / RTF document: Word's default.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [docx, other] = process.argv.slice(2)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 900, height: 1200 } })
const errors = []
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
const open = async path => {
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), path)
  await page.waitForFunction(() => document.querySelectorAll('.word-document p').length > 0, null, { timeout: 60_000 })
  await page.waitForTimeout(400)
}
// the computed ligatures of every text node, keyed by its script
const scripts = () => page.evaluate(() => {
  const out = []
  for (const root of document.querySelectorAll('.word-document')) { const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  for (let node; (node = walker.nextNode());) {
    const t = node.textContent
    const script = /[\u0600-\u06FF]/.test(t) ? 'arabic' : /[\u0590-\u05FF]/.test(t) ? 'hebrew' : /[a-z]/i.test(t) ? 'latin' : null
    if (script) out.push({ script, ligatures: getComputedStyle(node.parentElement).fontVariantLigatures })
  } }
  return out
})
await open(docx)
let found = await scripts()
console.log(JSON.stringify(found))
const hebrew = found.filter(f => f.script === 'hebrew'), arabic = found.filter(f => f.script === 'arabic')
check(hebrew.length >= 2 && hebrew.every(f => f.ligatures === 'no-common-ligatures no-contextual'), 'Hebrew: no ligatures, no contextual alternates, whatever the run sets')
check(arabic.length >= 2 && arabic.every(f => f.ligatures === 'normal'), 'Arabic: the browser ligatures (Word always forms them)')
await open(other)
found = await scripts()
check(found.length > 0 && found.filter(f => f.script !== 'arabic').every(f => f.ligatures === 'no-common-ligatures no-contextual'), `${other}: Word's default`)
check(errors.length === 0, `no page errors ${errors.join(' | ')}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
