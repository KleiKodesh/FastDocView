// Live check: a web archive opened for the first time, its index built on demand. Blocks drawn before their parent's end is
// recorded lack their later-sibling style (provisional); once the index completes the page fetches them again: the style
// appears and the reading position does not move. DOM only, no screenshots. ASCII output.
//   node playwright/provisionalStylesCheck.mjs <output folder> [paragraphs]
// Writes css-provisional-live-v1.mht there: one long div; every paragraph but the last has :nth-last-of-type(n+2) (green, and
// a bottom padding so the refreshed blocks change height), the first one :nth-last-of-type(N) (underlined).
import { chromium } from 'playwright'
import fs from 'node:fs'
import path from 'node:path'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const folder = process.argv[2], paragraphs = Number(process.argv[3] ?? 400000)
const plain = process.argv.includes('--plain')   // diagnostic: the same archive without the sheet (nothing provisional)
const file = path.resolve(folder, 'css-provisional-live-v1.mht')
const boundary = '----=_NextPart_01DB0000.00000003'
const lines = [`MIME-Version: 1.0`, `Content-Type: multipart/related; boundary="${boundary}"`, ``, `--${boundary}`,
  `Content-Location: file:///C:/css-provisional-live.htm`, `Content-Transfer-Encoding: 8bit`, `Content-Type: text/html; charset="us-ascii"`, ``,
  `<html>`, `<head>`, `<style>`, `<!--`,
  ...(plain ? ['p { margin:0 }'] : [`div.long > p:nth-last-of-type(n+2) { color:#2E7D32; padding-bottom:6pt; }`,
  `div.long > p:nth-last-of-type(${paragraphs}) { text-decoration:underline; }`]),
  `-->`, `</style>`, `</head>`, `<body>`, `<div class=long>`]
for (let i = 1; i <= paragraphs; i++) lines.push(`<p>Paragraph ${i} of the long div</p>`)
lines.push(`</div>`, `</body>`, `</html>`, ``, `--${boundary}--`, ``)
fs.writeFileSync(file, lines.join('\r\n'))

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 180_000 })
const blockFetches = []
page.on('request', r => { const m = /\/blocks\?first=(\d+)/.exec(r.url()); if (m) blockFetches.push({ first: Number(m[1]), at: Date.now() }) })

// the paragraph's style and its top on screen
const probe = n => page.evaluate(t => {
  for (const root of document.querySelectorAll('.word-document')) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
    for (let node; (node = walker.nextNode());) if (node.data === t) {
      const cs = getComputedStyle(node.parentElement)
      return { color: cs.color, deco: cs.textDecorationLine, top: node.parentElement.getBoundingClientRect().top }
    }
  }
  return null
}, `Paragraph ${n} of the long div`)

const openedAt = Date.now()
await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true, forceSidecar: true }), file)
await page.waitForFunction(() => [...document.querySelectorAll('.word-document')].some(e => e.textContent.includes('Paragraph 1 of the long div')), null, { timeout: 60_000 })
const firstPaintMs = Date.now() - openedAt
const completeAtPaint = await page.evaluate(() => window.__fastDocViewTest?.state()?.isIndexComplete === true)
// read somewhere past the top: the blocks there are drawn from the partial index too
await page.evaluate(() => window.__fastDocViewTest.scrollToBlock(200))
await page.waitForFunction(() => [...document.querySelectorAll('.word-document')].some(e => /Paragraph 2\d\d of/.test(e.textContent)), null, { timeout: 30_000 })
await page.waitForFunction(() => document.fonts.status === 'loaded')
// the reading position: the paragraph at the top edge of the scroller's viewport and its offset there
const readingPosition = () => page.evaluate(() => {
  const viewport = document.querySelector('.windowed-scroller-viewport')
  const top = viewport.getBoundingClientRect().top
  for (const p of viewport.querySelectorAll('.word-document p')) {
    const r = p.getBoundingClientRect(), m = /^Paragraph (\d+) of/.exec(p.textContent)
    if (m && r.bottom > top) return { paragraph: Number(m[1]), offset: r.top - top }
  }
  return null
})
const readingBefore = await readingPosition()
const anchor = readingBefore?.paragraph
const before = await probe(anchor), firstVisibleBefore = await page.evaluate(() => window.__fastDocViewTest.state().firstVisible)
console.log(`  first paint ${firstPaintMs} ms; index complete at first paint: ${completeAtPaint}; anchor paragraph ${anchor} ${JSON.stringify(before)} firstVisible ${JSON.stringify(firstVisibleBefore)}`)
check(!completeAtPaint, 'the index is still being built when the first blocks show')
check(before && before.color !== 'rgb(46, 125, 50)', 'drawn before the div ended: no later-sibling style yet')
await page.waitForFunction(() => window.__fastDocViewTest?.state()?.isIndexComplete === true, null, { timeout: 300_000, polling: 'raf' })
const completedAt = Date.now()
// the refetch of the provisional blocks on screen brings the style (no timer: wait on the DOM)
await page.waitForFunction(n => {
  for (const p of document.querySelectorAll('.word-document p')) if (p.textContent.startsWith(`Paragraph ${n} of`)) return getComputedStyle(p).color === 'rgb(46, 125, 50)'
  return false
}, anchor, { timeout: 30_000 }).catch(() => {})
const readingAfter = await readingPosition()
const after = await probe(anchor), firstVisibleAfter = await page.evaluate(() => window.__fastDocViewTest.state().firstVisible)
console.log(`  after ${JSON.stringify(after)} firstVisible ${JSON.stringify(firstVisibleAfter)}; block fetches ${JSON.stringify(blockFetches.map(f => f.first))}`)
check(after?.color === 'rgb(46, 125, 50)', `after the index completes: paragraph ${anchor} has :nth-last-of-type(n+2)`)
check(blockFetches.some(f => f.at >= completedAt - 50 && Math.abs(f.first - (firstVisibleBefore?.index ?? 0)) < 200), 'the provisional blocks on screen were fetched again once the index completed')
check(readingBefore && readingAfter && readingAfter.paragraph === readingBefore.paragraph && Math.abs(readingAfter.offset - readingBefore.offset) <= 2,   // scrollTop is whole pixels
  `reading position kept although the blocks grew: ${JSON.stringify(readingBefore)} -> ${JSON.stringify(readingAfter)}`)
await page.evaluate(() => window.__fastDocViewTest.scrollToBlock(0))
await page.waitForFunction(() => [...document.querySelectorAll('.word-document p')].some(p => p.textContent.startsWith('Paragraph 1 of') && getComputedStyle(p).textDecorationLine === 'underline'), null, { timeout: 30_000 }).catch(() => {})
const first = await probe(1)
check(first?.deco === 'underline', `paragraph 1 has :nth-last-of-type(${paragraphs}) (${JSON.stringify(first)})`)
check(errors.length === 0, 'no console errors ' + JSON.stringify(errors).replace(/[^ -~]/g, '?'))
await browser.close()
fs.rmSync(file, { force: true })
console.log(failures ? `FAILED ${failures}` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
