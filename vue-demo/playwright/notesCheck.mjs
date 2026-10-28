// Live check of footnotes/endnotes in the demo (npm run dev first):
//   node playwright/notesCheck.mjs [<file> …]   (defaults: both repo samples)
// For each file: open, jump to blocks that reference notes, and verify in the DOM that every
// reference on screen is numbered, its note is listed in the notes pane with a matching label,
// clicking it highlights the note, hovering shows a tooltip, and the console stays clean.
import { chromium } from 'playwright'

const DEMO_URL = 'http://localhost:5189/?file='
const files = process.argv.slice(2).length ? process.argv.slice(2) : ['samples/dor-hamelaktim-shabbat.doc', 'samples/dor-hamelaktim-shabbat.docx']

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const consoleErrors = []
page.on('console', message => { if (message.type() === 'error') consoleErrors.push(message.text()) })
page.on('pageerror', error => consoleErrors.push(`pageerror: ${error.message}`))
await page.goto(DEMO_URL, { waitUntil: 'networkidle' })
await page.evaluate(() => fetch('/api/documents', { method: 'DELETE' }))

for (const file of files) {
  console.log(`[${new Date().toLocaleTimeString()}] ${file}`)
  await page.evaluate(path => window.__wordDemo.open(path), file)
  await page.waitForFunction(() => window.__wordDemo.state()?.hasWeights, null, { timeout: 180_000 })
  const state = await page.evaluate(() => window.__wordDemo.state())
  console.log(`  ${state.document.blockCount.toLocaleString()} blocks, ${state.document.footnoteCount} footnotes, ${state.document.endnoteCount} endnotes, ${state.headingCount} headings`)
  check(state.document.footnoteCount + state.document.endnoteCount > 0, 'document has notes')

  const references = await page.evaluate(id => fetch(`/api/documents/${id}/note-references`).then(r => r.json()), state.document.id)
  check(references.length === state.document.footnoteCount + state.document.endnoteCount, `note-references lists all ${references.length}`)
  const numbers = references.filter(r => r.kind === 'footnote').map(r => r.number).filter(n => n > 0)
  check(numbers.every((n, i) => n === i + 1), 'footnote numbers run 1, 2, 3 … in reference order')

  const samples = [references[0], references[Math.floor(references.length / 2)], references.at(-1)].filter(Boolean)
  for (const reference of samples) {
    await page.evaluate(i => window.__wordDemo.scrollToBlock(i), reference.blockIndex)
    await page.waitForFunction(i => document.querySelector(`[data-block="${i}"] .note-reference`), reference.blockIndex, { timeout: 30_000 })
    await page.waitForTimeout(300)
    const onScreen = await page.$$eval(`[data-block="${reference.blockIndex}"] .note-reference`, sups => sups.map(s => ({ kind: s.dataset.noteKind, number: Number(s.dataset.noteNumber), text: s.textContent })))
    const target = onScreen.find(r => r.kind === reference.kind && r.number === reference.number)
    check(Boolean(target), `${reference.kind} ${reference.number} reference is in block ${reference.blockIndex} (shows "${target?.text}")`)

    const paneLabels = await page.$$eval('[data-test="note"] .note-mark', marks => marks.map(m => m.textContent))
    check(paneLabels.includes(target?.text ?? '?'), `notes pane lists note "${target?.text}" (pane shows ${paneLabels.length}: ${paneLabels.slice(0, 6).join(', ')})`)

    const sup = page.locator(`[data-block="${reference.blockIndex}"] .note-reference[data-note-number="${reference.number}"]`).first()
    await sup.click()
    await page.waitForTimeout(200)
    const highlighted = await page.$eval('[data-test="note"].highlighted .note-mark', mark => mark.textContent).catch(() => null)
    check(highlighted === target?.text, `clicking the reference highlights note "${highlighted}"`)

    await sup.hover()
    await page.waitForTimeout(150)
    const tooltipText = await page.$eval('[data-test="note-tooltip"]', tip => tip.textContent.trim()).catch(() => '')
    check(tooltipText.length > 0, `hover shows the note (${tooltipText.length} chars)`)
    const noteText = await page.$eval('[data-test="note"].highlighted', note => note.textContent.trim()).catch(() => '')
    check(noteText.length > 2, `note body has text (${noteText.length} chars)`)
  }
}

check(consoleErrors.length === 0, consoleErrors.length ? `console errors:\n    ${consoleErrors.join('\n    ')}` : 'no console errors')
await browser.close()
console.log(failures ? `\n${failures} FAILURE(S)` : '\nALL PASSED')
process.exit(failures ? 1 : 0)
