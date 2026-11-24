// Live check of .docx / .doc / RTF ruby (EQ phonetic guide) and ODT bullet fonts named by style:font-name, DOM only:
//   node playwright/rubyCheck.mjs <test-data/word-ruby/ruby-v1.docx> [.doc] [.rtf] <test-data/odt-bullet-fonts/bullet-font-name-v1.odt>
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
  if (/ruby-edges-v1\.(docx|doc|rtf)$/.test(path)) {
    // Word's PDF: ")" dropped with the text after it kept, tab / note mark / bookmark kept, a comma guide's rest before the base
    await page.waitForFunction(() => document.querySelectorAll('ruby').length >= 7, null, { timeout: 60_000 })
    const got = await page.evaluate(() => [...document.querySelectorAll('ruby')].map(r => {
      const base = [...r.childNodes].filter(n => n.nodeName !== 'RT').map(n => n.textContent).join('').replace(/\s+/g, ' ').trim()
      return r.querySelector('rt').textContent + '|' + base
    }))
    const expected = ['g0|abc d', 'g1|abc1', 'g2|abcy', 'g|3,base2', 'g|4,basew', 'g5|end of linex', 'g;6|b;6']
    expected.forEach((e, n) => check(got[n] === e, `${path}: ruby edge ${n}: ${e} (got ${got[n]})`))
    if (/docx$/.test(path)) check(await page.evaluate(() => !!document.querySelector('[id*="EdgeTwo"]') && !!document.querySelector('[id*="EdgeFour"]')), `${path}: bookmarks inside the rubies kept`)
    // the note marks inside rubies 1 and 3 are links to notes 1 and 2; the tab in ruby 0 is a tab
    check(await page.evaluate(() => { const r = [...document.querySelectorAll('ruby')]; return r[1].querySelector('.note-reference')?.textContent === '1' && r[3].querySelector('.note-reference')?.textContent === '2' && r[0].textContent.includes(String.fromCharCode(9)) }), `${path}: note marks 1 and 2 inside the rubies, the tab kept`)
    continue
  }
  if (/ruby-escapes/.test(path)) {
    // guide / base as Word's PDF draws them (the guide ends at the first comma, every ")" dropped)
    const expected = /probe/.test(path) ? ['g(0|base(0', 'g|1,base,1', String.raw`a\|b,c\,d`, String.raw`e\f|g\h`, String.raw`i\(j|k\\l`, 'g 5 x|base 5']
                                       : ['g(0|base(0', 'g|1,base,1', 'g;2|base;2', String.raw`g\3|base\3`, 'g4(|base4(', 'g 5 x|base 5']
    await page.waitForFunction(() => document.querySelectorAll('ruby').length >= 6, null, { timeout: 60_000 })
    const got = await page.evaluate(() => [...document.querySelectorAll('ruby')].map(r => r.querySelector('rt').textContent + '|' + [...r.childNodes].filter(n => n.nodeName !== 'RT').map(n => n.textContent).join('')))
    expected.forEach((e, n) => check(got[n] === e, `${path}: ruby escape ${n}: ${e} (got ${got[n]})`))
    continue
  }
  await page.waitForFunction(() => document.querySelectorAll('ruby').length >= 5, null, { timeout: 60_000 })
  const rubies = await page.evaluate(() => [...document.querySelectorAll('ruby')].map(r => {
    const rt = r.querySelector('rt'), base = document.createRange(); base.selectNodeContents(r.firstChild)
    // the guide text's own box: .docx puts the w:rt run's size on a span inside rt, .doc / RTF put hps / 2 pt on rt itself
    const text = [...rt.querySelectorAll('*')].reverse().find(e => e.textContent.trim()) ?? rt
    const s = getComputedStyle(text)
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
