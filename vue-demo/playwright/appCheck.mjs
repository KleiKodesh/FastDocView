// Live check of the APP's FastDocView page (the demo page IS the real page; only the transport is the
// demo's). DOM + console + the page's DEV test hook (window.__fastDocViewTest) - no screenshots.
//   node playwright/appCheck.mjs <file> [--headed]
// Cold first open (index rebuilt), count grows while the index builds, completes, exact jumps to the
// middle and the last block, footnote links both ways (when the file has notes), reload restores the
// position, no console errors.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
const headed = process.argv.includes('--headed')
if (!filePath) { console.error('usage: node playwright/appCheck.mjs <file> [--headed]'); process.exit(2) }

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch({ headless: !headed })
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const consoleErrors = []
page.on('console', message => { if (message.type() === 'error') consoleErrors.push(message.text().slice(0, 200)) })
page.on('pageerror', error => consoleErrors.push(String(error).slice(0, 200)))
// The hook is registered on mount; a hot reload (the app's files being edited) remounts the page, so wait for it each time.
const hookReady = () => page.waitForFunction(() => !!window.__fastDocViewTest, null, { timeout: 60_000 })
const state = async () => { await hookReady(); return page.evaluate(() => window.__fastDocViewTest.state()) }
const scrollToBlock = async index => { await hookReady(); await page.evaluate(i => window.__fastDocViewTest.scrollToBlock(i), index) }
const waitFor = (predicate, argument, timeout = 60_000) => page.waitForFunction(predicate, argument, { timeout })

step(`open ${DEMO_URL}`)
await page.goto(`${DEMO_URL}`, { waitUntil: 'domcontentloaded' })
await waitFor(() => !!window.__fastDocViewDemo, null, 120_000)

step(`cold first open (index rebuilt) ${filePath}`)
const openStarted = Date.now()
await page.evaluate(path => window.__fastDocViewDemo.open(path, { rebuildIndex: true, forceSidecar: true }), filePath)
await waitFor(() => [...document.querySelectorAll('.windowed-scroller-viewport [data-index]')].some(row => (row.textContent ?? '').trim().length > 0), null, 120_000)
const textShown = Date.now() - openStarted
const opened = await state()
console.log(`  text shown after ${textShown} ms; count ${opened?.blockCount}, complete ${opened?.isIndexComplete}`)
check(!!opened, 'the page exposes its DEV test hook')
check(textShown < 2000, 'text shown in under 2 s wall')
check(opened.blockCount === opened.knownBlockCount, 'the count is the known blocks, never an estimate')

if (!opened.isIndexComplete) {
  await page.waitForTimeout(1500)
  const building = await state()
  console.log(`  after 1.5 s: count ${building.blockCount}, complete ${building.isIndexComplete}, firstVisible ${building.firstVisible?.index}`)
  check(building.blockCount >= opened.blockCount, 'the count only grows while the index builds')
  check((building.firstVisible?.index ?? 0) <= 2, 'the reader stays at the top while the index builds')
}

step('the index completes')
await waitFor(() => window.__fastDocViewTest?.state().isIndexComplete, null, 180_000)
const complete = await state()
console.log(`  complete: count ${complete.blockCount}, headings ${complete.headingCount}, notes ${complete.footnoteCount}+${complete.endnoteCount}`)
check(complete.blockCount === complete.knownBlockCount, 'count exact once complete')

for (const target of [Math.floor(complete.blockCount / 2), complete.blockCount - 1]) {
  step(`scrollToBlock(${target})`)
  await scrollToBlock(target)
  await waitFor(index => !!document.querySelector(`.windowed-scroller-viewport [data-index="${index}"]`), target, 30_000).catch(() => {})
  await page.waitForTimeout(300)
  const landed = await state()
  const onScreen = await page.$eval(`.windowed-scroller-viewport [data-index="${target}"]`, row => {
    const view = row.closest('.windowed-scroller-viewport').getBoundingClientRect(), rect = row.getBoundingClientRect()
    return rect.bottom > view.top && rect.top < view.bottom
  }).catch(() => false)
  check(Math.abs((landed.firstVisible?.index ?? -1) - target) <= 2 || onScreen, `block ${target} is on screen (first visible ${landed.firstVisible?.index})`)
}

if (complete.footnoteCount + complete.endnoteCount > 0) {
  step('footnote links both ways')
  await scrollToBlock(0)
  // find a block with a reference: scan forward window by window through the hook (windowed)
  let found = false
  for (let block = 0; block < complete.blockCount && !found; block += 40) {
    await scrollToBlock(block)
    await page.waitForTimeout(150)
    found = await page.$('.windowed-scroller-viewport .note-reference[data-note-kind]').then(Boolean)
  }
  check(found, 'a note reference is rendered in the text')
  if (found) {
    await page.click('.windowed-scroller-viewport .note-reference[data-note-kind]')
    await page.waitForTimeout(400)
    const active = await page.$('.fast-doc-view-note--active')
    check(!!active && (await state()).notesPaneOpen, 'clicking a reference opens the notes pane on that note, highlighted')
    if (active) {
      await scrollToBlock(Math.max(0, (await state()).blockCount - 1))
      await page.waitForTimeout(400)
      const noteLabel = await page.$('.fast-doc-view-note--active .note-label')
      if (noteLabel) {
        await noteLabel.click()
        await page.waitForTimeout(600)
        check(!!(await page.$('.note-reference--current')), 'clicking the note label jumps back and flashes its reference')
      }
    }
  }
}

step('reload restores the position')
const middle = Math.floor(complete.blockCount / 3)
await scrollToBlock(middle)
await page.waitForTimeout(800)                        // the page saves its position debounced
const before = await state()
await page.reload({ waitUntil: 'domcontentloaded' })
await waitFor(() => window.__fastDocViewTest?.state()?.isIndexComplete && window.__fastDocViewTest.state().firstVisible, null, 120_000)
await page.waitForTimeout(1500)
const after = await state()
console.log(`  before: block ${before.firstVisible?.index} @ ${before.firstVisible?.offsetWithinRow?.toFixed(1)} px   after: block ${after.firstVisible?.index} @ ${after.firstVisible?.offsetWithinRow?.toFixed(1)} px`)
check(after.firstVisible?.index === before.firstVisible?.index, 'same first visible block after reload')
check(Math.abs((after.firstVisible?.offsetWithinRow ?? 0) - (before.firstVisible?.offsetWithinRow ?? 0)) <= 2, 'same offset within the block (2 px)')

check(consoleErrors.length === 0, `no console errors${consoleErrors.length ? ': ' + consoleErrors.slice(0, 3).join(' | ') : ''}`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
