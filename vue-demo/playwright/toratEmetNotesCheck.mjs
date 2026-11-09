// Live check of Torat Emet note lines: the header line carries the notes (after the source, joined by a middle dot), no plain note line is in the text, a
// title-only book keeps every line (light + dark).
//   node playwright/toratEmetNotesCheck.mjs <ramban file> <title-only file with a note-looking line> [<real title-only file>]      (ASCII output only)
import { chromium } from 'playwright'
const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [ramban, titleOnly, realTitleOnly] = process.argv.slice(2)
const NOTE = '\u05D4\u05E1\u05E4\u05E8 \u05D0\u05D9\u05E0\u05D5 \u05E9\u05DC\u05DD \u05D5\u05D0\u05D9\u05E0\u05D5 \u05DE\u05D5\u05D2\u05D4'
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const APP_SOURCE = (process.env.FASTDOCVIEW_APP_FRONTEND ?? '../../vue-frontend').split(String.fromCharCode(92)).join('/') + '/src'
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, 120_000)
await page.evaluate(source => { window.__appFrontendSource = source }, APP_SOURCE)
const read = async file => {
  await page.evaluate(path => window.__fastDocViewDemo.open(path, { rebuildIndex: true }), file)
  await page.waitForFunction(() => document.querySelectorAll('.word-document .te-meta').length >= 1, null, 20_000)
  await page.waitForTimeout(300)
  return page.evaluate(() => {
    const meta = document.querySelector('.word-document .te-meta')
    const lines = [...document.querySelectorAll('.windowed-scroller-viewport [data-index] .text-line:not(.te-meta)')].map(e => e.textContent.replace(/\u200b/g, '').trim())
    const css = getComputedStyle(meta)
    return { header: meta.textContent.trim(), lines, headingCount: document.querySelectorAll('.word-document .te-heading').length, fontSize: css.fontSize, color: css.color, rows: document.querySelectorAll('.windowed-scroller-viewport [data-index]').length,
             metaCount: document.querySelectorAll('.word-document .te-meta').length }
  })
}
const hex = s => [...s].map(c => c.codePointAt(0).toString(16)).join(' ')
for (const scheme of ['light', 'dark']) {
  await page.evaluate(async preset => (await import('/@fs/' + window.__appFrontendSource + '/theme/themes.ts')).applyTheme(preset), `default-${scheme}`)
  console.log(`[${scheme}]`)
  const a = await read(ramban)
  const parts = a.header.split(' \u00B7 ')
  check(a.metaCount === 1, `ramban: ONE header line (${a.metaCount})`)
  check(parts.length >= 2 && parts[parts.length - 1] === NOTE, `ramban: the header ends with the note, after the source (${parts.length} parts, last part code points ${hex(parts[parts.length - 1]).slice(0, 80)})`)
  check(!a.lines.some(l => l.includes(NOTE)), 'ramban: no plain note line in the text')
  check(a.headingCount >= 2 && a.lines.length >= 1, `ramban: the title and the first section heading follow the header (${a.headingCount} headings in view)`)
  check(parseFloat(a.fontSize) < 16, `ramban: the header stays small and muted (font ${a.fontSize}, color ${a.color})`)
  const b = await read(titleOnly)
  check(b.header.indexOf(NOTE) < 0 && b.lines.some(l => l === NOTE) && b.lines.includes('body a') && b.lines.includes('body b'), `title-only: the note-looking line stays text, header has no note (${b.header})`)
  if (realTitleOnly) { const c = await read(realTitleOnly); check(c.lines.length >= 2 && c.metaCount === 1, `real title-only book: ${c.lines.length} text lines in view, one header line`) }
}
check(errors.length === 0, `no console errors ${errors.slice(0, 2).join(' | ')}`)
await browser.close()
process.exit(failures ? 1 : 0)
