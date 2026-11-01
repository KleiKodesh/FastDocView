// Zoom check of the APP's FastDocView page through the demo: zoom 100 -> 150 -> 70 must scale the
// document text (body and largest heading-sized run) and the notes, keep the size ratios, keep the
// column's line breaks (text width scales with the text) and keep the reading position. DOM only.
//   node playwright/zoomCheck.mjs <file> [<file>...]
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const files = process.argv.slice(2)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const waitFor = (predicate, argument, timeout = 60_000) => page.waitForFunction(predicate, argument, { timeout })

// Effective on-screen font size of an element: its computed size times every ancestor's zoom.
const measure = () => page.evaluate(() => {
  const effective = element => {
    let scale = 1
    for (let node = element; node; node = node.parentElement) scale *= Number(getComputedStyle(node).zoom) || 1
    return Number.parseFloat(getComputedStyle(element).fontSize) * scale
  }
  const rows = [...document.querySelectorAll('.windowed-scroller-viewport [data-index]')]
  const sizes = new Map()
  for (const row of rows) {
    const walker = document.createTreeWalker(row, NodeFilter.SHOW_TEXT)
    for (let text = walker.nextNode(); text; text = walker.nextNode()) {
      if (!text.textContent.trim()) continue
      const size = Math.round(effective(text.parentElement) * 100) / 100
      sizes.set(size, (sizes.get(size) ?? 0) + text.textContent.length)
    }
  }
  const body = [...sizes].sort((a, b) => b[1] - a[1])[0]?.[0] ?? 0
  const largest = Math.max(0, ...sizes.keys())
  const note = document.querySelector('.notes-scroller *')
  // line breaks: the line count of the first multi-line paragraph-ish row
  const lineCounts = rows.slice(0, 12).map(row => {
    const range = document.createRange(); range.selectNodeContents(row)
    return new Set([...range.getClientRects()].map(r => Math.round(r.top))).size
  })
  const first = window.__fastDocViewTest.state().firstVisible
  return { body, largest, note: note ? effective(note) : null, firstIndex: first?.index ?? -1, rowIndexes: rows.slice(0, 12).map(r => r.dataset.index).join(','), lineCounts: lineCounts.join(',') }
})

await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await waitFor(() => !!window.__fastDocViewDemo, null, 120_000)
for (const file of files) {
  step(`open ${file}`)
  await page.evaluate(path => window.__fastDocViewDemo.open(path), file)
  await waitFor(() => window.__fastDocViewTest?.state().isIndexComplete, null, 180_000)
  const count = await page.evaluate(() => window.__fastDocViewTest.state().blockCount)
  await page.evaluate(i => window.__fastDocViewTest.scrollToBlock(i), Math.min(20, count - 1))
  await page.waitForTimeout(500)
  const results = {}
  for (const zoom of [100, 150, 70]) {
    await page.evaluate(z => window.__fastDocViewTest.setZoom(z), zoom)
    await page.waitForTimeout(600)
    results[zoom] = await measure()
    console.log(`  zoom ${zoom}: ${JSON.stringify(results[zoom])}`)
  }
  for (const zoom of [150, 70]) {
    const ratio = zoom / 100, near = (a, b) => Math.abs(a - b) <= 0.02 * Math.max(1, b)
    check(near(results[zoom].body, results[100].body * ratio), `zoom ${zoom}: body ${results[zoom].body} = ${results[100].body} x ${ratio}`)
    check(near(results[zoom].largest, results[100].largest * ratio), `zoom ${zoom}: largest ${results[zoom].largest} = ${results[100].largest} x ${ratio}`)
    if (results[100].note != null && results[zoom].note != null) check(results[zoom].note === results[100].note, `zoom ${zoom}: the notes keep their own zoom (${results[zoom].note})`)
    check(results[zoom].firstIndex === results[100].firstIndex, `zoom ${zoom}: position kept (first visible ${results[zoom].firstIndex})`)
    if (results[zoom].rowIndexes === results[100].rowIndexes) check(results[zoom].lineCounts === results[100].lineCounts, `zoom ${zoom}: same line breaks (${results[zoom].lineCounts})`)
  }
  await page.evaluate(() => window.__fastDocViewTest.setZoom(100))

  // position kept in the middle of the document
  const total = await page.evaluate(() => window.__fastDocViewTest.state().knownBlockCount)
  console.log(`  blocks: ${total}`)
  if (total > 60) {
    await page.evaluate(i => window.__fastDocViewTest.scrollToBlock(i), Math.floor(total / 2))
    await page.waitForTimeout(600)
    const before = await page.evaluate(() => window.__fastDocViewTest.state().firstVisible)
    for (const zoom of [150, 70, 100]) {
      await page.evaluate(z => window.__fastDocViewTest.setZoom(z), zoom)
      await page.waitForTimeout(600)
      const after = await page.evaluate(() => window.__fastDocViewTest.state().firstVisible)
      check(after.index === before.index, `mid-document zoom ${zoom}: first visible block kept (${before.index} -> ${after.index}, offset ${Math.round(after.offsetWithinRow)})`)
    }
  }

  // the zoom model: Ctrl+wheel over a pane zooms that pane only; the toolbar steps both
  const zooms = () => page.evaluate(() => { const s = window.__fastDocViewTest.state(); return { text: s.zoom, notes: s.notesZoom } })
  const box = async selector => { const b = await page.locator(selector).first().boundingBox(); return b && { x: b.x + b.width / 2, y: b.y + b.height / 2 } }
  const textAt = await box('.fast-doc-view-scroller')
  await page.mouse.move(textAt.x, textAt.y); await page.keyboard.down('Control'); await page.mouse.wheel(0, -20); await page.keyboard.up('Control')
  await page.waitForTimeout(200)
  let z = await zooms()
  check(z.text === 120 && z.notes === 100, `Ctrl+wheel over the text: text only (${JSON.stringify(z)})`)
  if (await page.locator('.notes-scroller').count() === 0 && await page.evaluate(() => { const s = window.__fastDocViewTest.state(); return s.footnoteCount + s.endnoteCount > 0 })) {
    // the notes step button: the first one after the section-nav (chevron) buttons
    await page.evaluate(() => { const buttons = [...document.querySelectorAll('.fast-doc-view-toolbar button')]; const last = buttons.findLastIndex(b => b.querySelector('.dir-icon')); buttons[last + 1].click() })
    await page.waitForTimeout(800)
  }
  const notesAt = await page.locator('.notes-scroller').count() ? await box('.notes-scroller') : null
  if (notesAt) {
    await page.mouse.move(notesAt.x, notesAt.y); await page.keyboard.down('Control'); await page.mouse.wheel(0, -30); await page.keyboard.up('Control')
    await page.waitForTimeout(200)
    z = await zooms()
    check(z.text === 120 && z.notes === 130, `Ctrl+wheel over the notes: notes only (${JSON.stringify(z)})`)
    const noteSize = () => page.evaluate(() => Number.parseFloat(getComputedStyle(document.querySelector('.notes-scroller')).fontSize))
    console.log(`  notes font px at notes zoom 130: ${await noteSize()}`)
  } else console.log('  (no notes pane open: notes wheel not checked)')
  const before = await zooms()
  const zoomInButton = page.locator('.fast-doc-view-toolbar button[title*="(Ctrl+)"]').first()
  await zoomInButton.dispatchEvent('pointerdown'); await zoomInButton.dispatchEvent('pointerup')
  await page.waitForTimeout(200)
  z = await zooms()
  check(z.text === before.text + 2 && z.notes === before.notes + 2, `toolbar zoom in: both +2 (${JSON.stringify(before)} -> ${JSON.stringify(z)})`)
  await page.locator('.fast-doc-view-toolbar').first().focus().catch(() => {})
  await page.evaluate(() => { document.activeElement?.blur(); document.querySelector('.fast-doc-view-toolbar button')?.focus() })
  await page.keyboard.press('Control+Digit0')
  await page.waitForTimeout(200)
  z = await zooms()
  check(z.text === 100 && z.notes === 100, `Ctrl+0 with focus outside both panes: both reset (${JSON.stringify(z)})`)
  await page.locator('.fast-doc-view-scroller').first().focus()
  await page.keyboard.press('Control+Equal')
  await page.waitForTimeout(200)
  z = await zooms()
  check(z.text === 102 && z.notes === 100, `Ctrl+= with focus in the text: text only (${JSON.stringify(z)})`)
  await page.evaluate(() => { window.__fastDocViewTest.setZoom(100); window.__fastDocViewTest.setNotesZoom(100) })
}
await browser.close()
console.log(failures ? `${failures} FAILED` : 'ALL PASS')
process.exit(failures ? 1 : 0)
