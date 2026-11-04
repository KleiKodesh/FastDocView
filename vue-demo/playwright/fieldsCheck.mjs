// Live check of Word fields in the APP's FastDocView page through the demo: page placeholders, computed results, EQ as
// MathML, form fields, in-document field links, and the search finding computed text but never codes. DOM + console +
// the page's DEV hook - no screenshots.
//   node playwright/fieldsCheck.mjs <file> [<file> ...]   (test-data/word-fields/fields-v1.doc / .rtf, the --fields .docx)
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const files = process.argv.slice(2)
if (files.length === 0) { console.error('usage: node playwright/fieldsCheck.mjs <file> [<file> ...]'); process.exit(2) }

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const consoleErrors = []
page.on('console', message => { if (message.type() === 'error') consoleErrors.push(message.text().slice(0, 200)) })
page.on('pageerror', error => consoleErrors.push(String(error).slice(0, 200)))
const waitFor = (predicate, argument, timeout = 60_000) => page.waitForFunction(predicate, argument, { timeout })

await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await waitFor(() => !!window.__fastDocViewDemo, null, 120_000)
for (const file of files) {
  console.log(`[${new Date().toLocaleTimeString()}] ${file}`)
  await page.evaluate(path => window.__fastDocViewDemo.open(path), file)
  await waitFor(() => window.__fastDocViewTest?.state().isIndexComplete, null, 120_000)
  await page.evaluate(() => window.__fastDocViewTest.scrollToBlock(0))
  await page.waitForTimeout(800)
  const view = await page.evaluate(() => {
    const viewport = document.querySelector('.windowed-scroller-viewport')
    return {
      text: viewport?.innerText ?? '',
      placeholders: viewport?.querySelectorAll('.field-page-placeholder').length ?? 0,
      linkedPlaceholder: !!viewport?.querySelector('a[href^="#bookmark-"] .field-page-placeholder'),
      fraction: !!viewport?.querySelector('math mfrac'),
      fractionWidth: viewport?.querySelector('math')?.getBoundingClientRect().width ?? 0,
    }
  })
  const isFixture = !file.includes('fields-lists')   // the harness's TOC / list fixture: only its link is checked
  if (isFixture) check(view.placeholders >= 2 && view.linkedPlaceholder, `page placeholders drawn (${view.placeholders}), PAGEREF \\h one inside its link`)
  if (isFixture) check(view.fraction && view.fractionWidth > 0, 'EQ \\f(1,2) laid out as MathML')
  if (isFixture) check(view.text.includes('Press here') && view.text.includes('Blueberry') && view.text.includes('Moshe Cohen'), 'computed MACROBUTTON, drop-down entry, AUTHOR shown')
  if (isFixture) check(!/MACROBUTTON|FORMCHECKBOX|hiddenentry|nestedno/.test(view.text), 'no field code, hidden field or unchosen IF branch shown')
  // a field link jumps to its bookmark in the document
  const jumped = !isFixture || await page.evaluate(async () => {
    const link = document.querySelector('.windowed-scroller-viewport a[href="#bookmark-Target"]')
    if (!link) return false
    window.__fastDocViewTest.scrollToBlock(Math.max(0, window.__fastDocViewTest.state().blockCount - 1))
    await new Promise(r => setTimeout(r, 300))
    const again = document.querySelector('.windowed-scroller-viewport a[href="#bookmark-Target"]')
    ;(again ?? link).click()
    await new Promise(r => setTimeout(r, 800))
    const target = document.querySelector('#bookmark-Target')
    return !!target && target.getBoundingClientRect().top < window.innerHeight
  })
  check(jumped, 'a REF / HYPERLINK \\l field link brings its bookmark into view')
  // search: computed text found, field codes never
  // a TOC FastDocView built (\h): "#block-N" jumps to the heading's block through the app page
  const blockLink = await page.evaluate(() => document.querySelector('.windowed-scroller-viewport a[href^="#block-"]')?.getAttribute('href') ?? null)
  if (blockLink) {
    const target = Number(blockLink.slice('#block-'.length))
    await page.evaluate(() => window.__fastDocViewTest.scrollToBlock(window.__fastDocViewTest.state().blockCount - 1))
    await page.waitForTimeout(300)
    await page.evaluate(() => window.__fastDocViewTest.scrollToBlock(0))
    await page.waitForTimeout(300)
    await page.click(`.windowed-scroller-viewport a[href="${blockLink}"]`)
    await page.waitForTimeout(800)
    const shown = await page.evaluate(index => {
      const row = document.querySelector(`.windowed-scroller-viewport [data-index="${index}"]`)
      const box = row?.getBoundingClientRect()
      return !!box && box.bottom > 0 && box.top < window.innerHeight
    }, target)
    check(shown, `TOC \h link ${blockLink} brings its heading block into view`)
  }
  if (!isFixture) continue
  await page.click('.windowed-scroller-viewport')
  await page.keyboard.press('Control+f')
  await waitFor(() => !!document.querySelector('.search-bar input'))
  for (const [query, found] of [['Press here', true], ['MACROBUTTON', false], ['hiddenentry', false]]) {
    await page.fill('.search-bar input', query)
    await page.waitForTimeout(1500)
    const s = await page.evaluate(() => window.__fastDocViewTest.state().search)
    check(found ? s.matchCount > 0 : s.matchCount === 0, `search "${query}": ${found ? 'found' : 'not found'} (${s.matchCount} matches)`)
  }
  await page.keyboard.press('Escape')
}
check(consoleErrors.length === 0, `no console errors${consoleErrors.length ? ': ' + consoleErrors.join(' | ') : ''}`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
