// Live check of the APP's in-document search over phonetic guides (ruby), DOM only:
//   node playwright/rubySearchCheck.mjs <test-data folder>
// Word's Find reads a guide before its base ("gdbase" is one match, never "basegd"), and a guide with a comma as written
// ("g,3": head, comma, rest; its base not at all). The page draws base then guide, so a match across them is a first mark
// plus continuation marks sharing its data-search-part; every part becomes current together, and the ruby stays whole.
import { chromium } from 'playwright'
import path from 'node:path'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const root = process.argv[2]
if (!root) { console.error('usage: node playwright/rubySearchCheck.mjs <test-data folder>'); process.exit(2) }
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })

// (file, query, matches, [first mark's text, continuation texts in page order] of the match stepped to, or null for none)
const cases = []
for (const ext of ['docx', 'doc', 'rtf']) {
  const eq = path.join(root, 'word-fields', `eq-search-text-v1.${ext}`)
  cases.push([eq, 'gdbase', 1, ['gd', ['base']]], [eq, 'basegd', 0, null], [eq, 'base tailruby', 1, ['base', [' tailruby']]], [eq, ': gd', 1, [': ', ['gd']]])
  const edges = path.join(root, 'word-ruby', `ruby-edges-v1.${ext}`)
  cases.push([edges, 'g,3 after', 1, ['g', ['3', ' after']]], [edges, '3,base', 0, null], [edges, 'before g,4', 1, ['before ', ['4', 'g']]])
}

let opened = ''
for (const [file, query, expected, parts] of cases) {
  if (file !== opened) {
    await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), file)
    await page.waitForFunction(() => window.__fastDocViewTest?.state().isIndexComplete && document.querySelectorAll('ruby').length > 0, null, { timeout: 120_000 })
    opened = file
    await page.click('.windowed-scroller-viewport')
    await page.keyboard.press('Control+f')
    await page.waitForFunction(() => !!document.querySelector('.search-bar input'))
  }
  const name = `${path.basename(file)} "${query}"`
  const rubies = await page.evaluate(() => document.querySelectorAll('ruby rt').length)
  await page.fill('.search-bar input', query)
  await page.waitForFunction(q => { const s = window.__fastDocViewTest.state().search; return !s.isSearching && document.querySelector('.search-bar input').value === q }, query, { timeout: 60_000 })
  await page.waitForTimeout(500)
  const count = await page.evaluate(() => window.__fastDocViewTest.state().search.matchCount)
  check(count === expected, `${name}: ${expected} match(es) counted (${count})`)
  const marked = await page.evaluate(() => document.querySelectorAll('.windowed-scroller-viewport mark.search-match').length)
  check(marked === expected, `${name}: ${expected} first mark(s) on the page (${marked})`)
  check(await page.evaluate(() => document.querySelectorAll('ruby rt').length) === rubies && await page.evaluate(() => [...document.querySelectorAll('rt')].every(rt => rt.parentElement.tagName === 'RUBY')), `${name}: every ruby still whole`)
  if (!parts) continue
  await page.keyboard.press('Enter')
  await page.waitForTimeout(400)
  const current = await page.evaluate(() => {
    const head = document.querySelector('.windowed-scroller-viewport mark.search-match.current')
    if (!head) return null
    const continued = [...document.querySelectorAll(`.windowed-scroller-viewport mark.search-match-continued[data-search-part="${head.dataset.searchPart}"]`)]
    return { head: head.textContent, continued: continued.filter(m => m.textContent).map(m => m.textContent), currentParts: continued.every(m => m.classList.contains('current')) }
  })
  check(!!current && current.head === parts[0] && JSON.stringify(current.continued) === JSON.stringify(parts[1]) && current.currentParts,
        `${name}: the match is marked "${parts[0]}" + ${JSON.stringify(parts[1])}, all current (${JSON.stringify(current)})`)
}
check(errors.length === 0, `no console errors (${errors.join(' | ')})`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
