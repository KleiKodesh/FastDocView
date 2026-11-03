// Live check of the notes pane's default size and column (DOM measurement, no screenshots):
//   node playwright/notesLayoutCheck.mjs <file-with-notes>
// The pane opens at a quarter of the page below, right and left; each note's column is the document column
// (same max-width and gutter as the text blocks).
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
if (!filePath) { console.error('usage: node playwright/notesLayoutCheck.mjs <file>'); process.exit(2) }
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } })
step(`open ${filePath}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
await page.evaluate(path => window.__fastDocViewDemo.open(path, {}), filePath)
await page.waitForFunction(() => (window.__fastDocViewTest?.state().footnoteCount ?? 0) + (window.__fastDocViewTest?.state().endnoteCount ?? 0) > 0, null, { timeout: 180_000 })
await page.waitForFunction(() => document.querySelectorAll('.fast-doc-view-block').length > 0, null, { timeout: 60_000 })

for (const place of ['bottom', 'right', 'left']) {
  step(`notes pane ${place}`)
  await page.evaluate(where => window.__fastDocViewTest.setNotesPlace(where), place)
  await page.waitForFunction(() => document.querySelectorAll('.fast-doc-view-notes-pane .fast-doc-view-note').length > 0, null, { timeout: 60_000 })
  await page.waitForTimeout(400)
  const measured = await page.evaluate(where => {
    const split = document.querySelector('.fast-doc-view-notes-split').getBoundingClientRect()
    const pane = document.querySelector('.fast-doc-view-notes-pane').getBoundingClientRect()
    const block = document.querySelector('.fast-doc-view-block'), note = document.querySelector('.fast-doc-view-notes-pane .fast-doc-view-note')
    const blockStyle = getComputedStyle(block), noteStyle = getComputedStyle(note)
    return {
      share: where === 'bottom' ? pane.height / split.height : pane.width / split.width,
      block: [blockStyle.paddingInlineStart, blockStyle.maxWidth], note: [noteStyle.paddingInlineStart, noteStyle.maxWidth],
      fonts: [getComputedStyle(note.querySelector('.note-body')).fontSize, noteStyle.fontSize],
    }
  }, place)
  check(Math.abs(measured.share - 0.25) < 0.02, `the pane takes a quarter of the page (${(measured.share * 100).toFixed(1)}%)`)
  check(measured.fonts[0] === measured.fonts[1], `the note text keeps the pane's size (${measured.fonts.join(' / ')})`)
  check(measured.note[0] === measured.block[0] && measured.note[1] === measured.block[1],
        `notes follow the document column: gutter ${measured.note[0]} / ${measured.block[0]}, max-width ${measured.note[1]} / ${measured.block[1]}`)
}
await browser.close()
console.log(failures ? `${failures} FAILED` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
