// Live check of .doc / RTF ruby (EQ phonetic guide) and ODT bullet fonts named by style:font-name, DOM only:
//   node playwright/rubyCheck.mjs <test-data/word-ruby/ruby-v1.doc> [.rtf] <test-data/odt-bullet-fonts/bullet-font-name-v1.odt>
// Each ruby guide sits above its base (the guide's box ends above the base's middle; line boxes overlap) in Arial at the EQ's hps / 2 pt; each Wingdings bullet draws in Wingdings.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
for (const path of process.argv.slice(2)) {
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), path)
  if (/\.odt$|\.fodt$/.test(path)) {
    await page.waitForFunction(() => document.body.innerText.includes('font-name Wingdings'), null, { timeout: 60_000 })
    const wingdings = await page.evaluate(() => [...document.querySelectorAll('[style*="Wingdings"]')].some(e => e.textContent.includes('')))
    check(wingdings, `${path}: the font-face bullet draws in Wingdings`)
    continue
  }
  await page.waitForFunction(() => document.querySelectorAll('ruby').length >= 5, null, { timeout: 60_000 })
  const rubies = await page.evaluate(() => [...document.querySelectorAll('ruby')].map(r => {
    const rt = r.querySelector('rt'), base = document.createRange(); base.selectNodeContents(r.firstChild)
    const s = getComputedStyle(rt)
    const b = base.getBoundingClientRect()
    return { rtBottom: rt.getBoundingClientRect().bottom, baseMiddle: (b.top + b.bottom) / 2, size: parseFloat(s.fontSize), family: s.fontFamily }
  }))
  const sizes = [10, 8, 10, 12, 10]
  rubies.slice(0, 5).forEach((r, n) => {
    check(r.rtBottom < r.baseMiddle, `${path}: ruby ${n} guide above its base`)
    check(Math.abs(r.size * 0.75 - sizes[n]) < 0.1 && r.family.includes('Arial'), `${path}: ruby ${n} guide ${sizes[n]} pt Arial (${(r.size * 0.75).toFixed(2)} pt)`)
  })
}
check(errors.length === 0, `no console errors (${errors.join(' | ')})`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
