// Live check of the APP's footnote-reference tooltip (the book view's word-link panel, full lifecycle). DOM + console only.
//   node playwright/noteTooltipCheck.mjs <file-with-notes>
// Hover opens after the delay; the panel has the text-size and close buttons; the pointer can cross into it and it stays;
// its text is selectable; leaving closes after the grace; Esc closes; a click on the reference lands the note in the pane.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 1000 } })
const consoleErrors = []
page.on('console', m => { if (m.type() === 'error') consoleErrors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => consoleErrors.push(String(e).slice(0, 200)))
const waitFor = (fn, arg, timeout = 60_000) => page.waitForFunction(fn, arg, { timeout })

step('open')
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await waitFor(() => !!window.__fastDocViewDemo, null, 120_000)
await page.evaluate(p => window.__fastDocViewDemo.open(p), filePath)
await waitFor(() => window.__fastDocViewTest?.state().isIndexComplete, null, 180_000)
await waitFor(() => !!document.querySelector('.windowed-scroller-viewport .note-reference[data-note-kind]'))
const reference = page.locator('.windowed-scroller-viewport .note-reference[data-note-kind]').first()
await reference.scrollIntoViewIfNeeded()
await page.waitForTimeout(300)

step('hover')
await reference.hover()
await waitFor(() => !!document.querySelector('.word-link-tooltip'), null, 3000).catch(() => {})
const panel = await page.evaluate(() => {
  const t = document.querySelector('.word-link-tooltip')
  if (!t) return null
  return {
    isStatic: t.classList.contains('is-static'),
    buttons: t.querySelectorAll('.word-link-tooltip-button, .word-link-tooltip-close').length,
    pointer: getComputedStyle(t).pointerEvents,
    select: getComputedStyle(t.querySelector('.word-link-tooltip-body')).userSelect,
    resize: getComputedStyle(t).resize,
    text: t.querySelector('.word-link-tooltip-body').textContent.trim().slice(0, 40),
  }
})
console.log('  panel', JSON.stringify(panel))
check(panel && !panel.isStatic && panel.buttons === 3 && panel.pointer === 'auto' && panel.text.length > 0, 'interactive panel with zoom out / zoom in / close and the note text')
check(panel && panel.select !== 'none' && panel.resize === 'both', 'text selectable, resizable')

step('cross into the panel')
const box = await page.locator('.word-link-tooltip').boundingBox()
await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2, { steps: 8 })
await page.waitForTimeout(500)
check(await page.locator('.word-link-tooltip').count() === 1, 'stays open with the pointer inside')
const before = await page.evaluate(() => getComputedStyle(document.querySelector('.word-link-tooltip')).fontSize)
await page.locator('.word-link-tooltip-button').nth(1).click()
const after = await page.evaluate(() => getComputedStyle(document.querySelector('.word-link-tooltip')).fontSize)
check(parseFloat(after) > parseFloat(before), `zoom in button grows the text (${before} -> ${after})`)
await page.locator('.word-link-tooltip-button').nth(0).click()

step('leave')
await page.mouse.move(5, 5, { steps: 4 })
await page.waitForTimeout(600)
check(await page.locator('.word-link-tooltip').count() === 0, 'closes after leaving')

step('Esc')
await reference.hover()
await waitFor(() => !!document.querySelector('.word-link-tooltip'), null, 3000).catch(() => {})
await page.keyboard.press('Escape')
await page.waitForTimeout(100)
check(await page.locator('.word-link-tooltip').count() === 0, 'Esc closes')

step('click lands the note')
await reference.click()
await page.waitForTimeout(1500)
check(await page.locator('.word-link-tooltip').count() === 0, 'click closes the preview')
check(await page.evaluate(() => !!document.querySelector('.notes-scroller')), 'notes pane open after the click')

check(consoleErrors.length === 0, `no console errors ${JSON.stringify(consoleErrors)}`)
await browser.close()
console.log(failures ? `${failures} FAILED` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
