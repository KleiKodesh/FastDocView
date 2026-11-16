// Live check: a Word web archive's class styles resolved inline, and the ODT paragraph-wide script split (ASCII output).
//   node playwright/webArchiveSheetCheck.mjs <word-saved .mht> <odt-paragraph-script .fodt> [<table-margin .mht>]
import { chromium } from 'playwright'
const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [mht, fodt, tableMht] = process.argv.slice(2)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
page.on('response', r => { if (r.status() >= 400) console.log('  http ' + r.status() + ' ' + r.url().slice(0, 160) + ' ' + JSON.stringify(r.request().postData() ?? '').replace(/[^ -~]/g, '?').slice(0, 200)) })
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, 120_000)
const styleOf = text => page.evaluate(t => {
  for (const root of document.querySelectorAll('.word-document')) {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  for (let node; (node = walker.nextNode());) if (node.data.includes(t)) {
    const cs = getComputedStyle(node.parentElement)
    return { family: cs.fontFamily, size: cs.fontSize, weight: cs.fontWeight, style: cs.fontStyle, color: cs.color, align: getComputedStyle(node.parentElement.closest('p,h1,h2,div')).textAlign }
  } }
  return null
}, text)
await page.evaluate(path => window.__fastDocViewDemo.open(path, { rebuildIndex: true }), mht)
await page.waitForFunction(() => [...document.querySelectorAll('.word-document')].some(e => e.textContent.includes('Before')), null, 30_000)
const fancy = await styleOf('Fancy styled'), heading = await styleOf('Heading one'), green = await styleOf('green')
console.log('  ' + JSON.stringify({ fancy, heading, green }))
check(fancy && /Courier New/.test(fancy.family) && fancy.size === '26.6667px' && Number(fancy.weight) >= 700 && fancy.color === 'rgb(255, 0, 0)' && fancy.align === 'center',
  'mht: Word paragraph style (Courier New 20pt bold red centred) as Word prints it')
check(heading && /Calibri Light/.test(heading.family) && heading.size === '26.6667px' && heading.color === 'rgb(47, 84, 150)', 'mht: Heading 1 (Calibri Light 20pt #2F5496)')
check(green && green.style === 'italic' && green.color === 'rgb(0, 128, 0)' && green.size === '21.3333px', 'mht: character style (green italic 16pt)')
const sheets = await page.evaluate(() => [...document.querySelectorAll('style')].map(s => s.textContent).join('\n'))
check(!sheets.includes('MsoNormal') && !sheets.includes('FancyPara'), 'mht: the archive sheet is not on the page')
await page.evaluate(path => window.__fastDocViewDemo.open(path, { rebuildIndex: true }), fodt)
await page.waitForFunction(() => [...document.querySelectorAll('.word-document')].some(e => e.textContent.includes('2024')), null, 30_000)
const digits = await styleOf(' 2024'), english = await styleOf('English')
console.log('  ' + JSON.stringify({ digits, english }))
check(digits && english && digits.size === english.size && digits.size === '14.6667px', 'fodt: digits after an English word keep the Latin size across a span (LibreOffice)')
if (tableMht) {
  // Word's PDF (test-data/word-mhtml/table-margin-v1.pdf): body text at 90.05 pt, cell text at 90.29 pt, table border at 84.7 pt
  await page.evaluate(path => window.__fastDocViewDemo.open(path, { rebuildIndex: true }), tableMht)
  await page.waitForFunction(() => [...document.querySelectorAll('.word-document')].some(e => e.textContent.includes('CellA')), null, 30_000)
  const x = await page.evaluate(() => {
    const left = t => { for (const root of document.querySelectorAll('.word-document')) { const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      for (let n; (n = w.nextNode());) if (n.data.includes(t)) { const r = document.createRange(); r.setStart(n, n.data.indexOf(t)); r.setEnd(n, n.data.indexOf(t) + 1); return r.getBoundingClientRect().left } } return NaN }
    return { body: left('Body line'), cell: left('CellA'), border: document.querySelector('.word-document table td')?.getBoundingClientRect().left }
  })
  console.log('  ' + JSON.stringify(x))
  check(Math.abs(x.cell - x.body) <= 1.5 && Math.abs((x.body - x.border) - 7.2) <= 1.5, 'table mht: cell text lines up with the body text, border 5.4 pt to the left (Word PDF)')
}
check(errors.length === 0, `no console errors ${JSON.stringify(errors)}`)
await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILED`)
process.exit(failures === 0 ? 0 : 1)
