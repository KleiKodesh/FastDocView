// Live check of the Torat Emet start of a book: one quiet header line, no rules around the title, a useful tab title.
//   node playwright/toratEmetHeaderCheck.mjs <file> [<file> ...]      (ASCII output only: Hebrew is reported as counts / code points)
import { chromium } from 'playwright'
const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
import { readFileSync } from 'node:fs'
const files = process.argv[2] === '--list' ? readFileSync(process.argv[3], 'utf8').split(/\r?\n/).filter(Boolean) : process.argv.slice(2)
const ascii = text => [...(text ?? '')].map(c => (c.codePointAt(0) < 128 ? c : 'U+' + c.codePointAt(0).toString(16))).join('').slice(0, 90)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const APP_SOURCE = (process.env.FASTDOCVIEW_APP_FRONTEND ?? '../../vue-frontend').split('\\').join('/') + '/src'
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, 120_000)
await page.evaluate(source => { window.__appFrontendSource = source }, APP_SOURCE)
for (const file of files) {
  console.log(`== ${ascii(file)}`)
  const title = await page.evaluate(async path => {
    const response = await fetch('/api/documents', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ path }) })
    const body = await response.json()
    return response.ok ? (body.title ?? null) : 'HTTP ' + response.status
  }, file)
  await page.evaluate(path => window.__fastDocViewDemo.open(path, { rebuildIndex: true }), file)
  await page.waitForFunction(() => [...document.querySelectorAll('.windowed-scroller-viewport [data-index]')].some(r => (r.textContent ?? '').trim().length > 0), null, 60_000)
  for (const scheme of ['light', 'dark']) {
    await page.evaluate(async preset => (await import('/@fs/' + window.__appFrontendSource + '/theme/themes.ts')).applyTheme(preset), `default-${scheme}`)
    await page.waitForTimeout(250)
    const info = await page.evaluate(() => {
      const css = (el, p) => getComputedStyle(el)[p]
      const rows = [...document.querySelectorAll('.windowed-scroller-viewport [data-index]')].slice(0, 6)
      const meta = document.querySelector('.word-document .te-meta')
      const headings = [...document.querySelectorAll('.word-document .te-heading')].slice(0, 3).map(block => {
        const h = block.firstElementChild
        return { tag: h.tagName, text: h.textContent, borderTop: css(h, 'borderTopWidth'), borderBottom: css(h, 'borderBottomWidth'),
          blockBorderTop: css(block, 'borderTopWidth'), blockBorderBottom: css(block, 'borderBottomWidth'), margin: css(h, 'marginTop') + '/' + css(h, 'marginBottom') }
      })
      return {
        metaText: meta?.textContent ?? null, metaDisplay: meta ? css(meta, 'display') : null, metaColor: meta ? css(meta, 'color') : null, metaSize: meta ? css(meta, 'fontSize') : null,
        metaBorder: meta ? css(meta, 'borderBottomWidth') : null, bodyColor: css(document.querySelector('.word-document .text-line:not(.te-meta)') ?? document.body, 'color'),
        bodySize: css(document.querySelector('.word-document .text-line:not(.te-meta)') ?? document.body, 'fontSize'),
        hrCount: document.querySelectorAll('.word-document hr').length, headings,
        rowHeights: rows.map(r => Math.round(r.getBoundingClientRect().height)),
      }
    })
    console.log(`  [${scheme}] header ${info.metaDisplay === 'none' ? 'hidden (empty)' : JSON.stringify(ascii(info.metaText))} size ${info.metaSize} (body ${info.bodySize}) color ${info.metaColor} (body ${info.bodyColor}) border ${info.metaBorder}`)
    console.log(`  [${scheme}] rows ${info.rowHeights.join(',')} hr ${info.hrCount}; headings ${info.headings.map(h => `${h.tag} b=${h.borderTop}/${h.borderBottom} blockB=${h.blockBorderTop}/${h.blockBorderBottom} m=${h.margin}`).join(' | ')}`)
    check(info.hrCount === 0 && info.headings.every(h => h.borderTop === '0px' && h.borderBottom === '0px' && h.blockBorderTop === '0px' && h.blockBorderBottom === '0px'), `${scheme}: no rules around headings`)
    check(info.metaDisplay === 'none' || (info.metaBorder === '0px' && info.metaColor !== info.bodyColor && parseFloat(info.metaSize) < parseFloat(info.bodySize)), `${scheme}: header is muted, smaller, no rule (or hidden when empty)`)
  }
  console.log(`  tab title (summary title): ${title === null ? 'null (file name stays)' : JSON.stringify(ascii(title))}`)
}
check(errors.length === 0, `no console errors ${errors.slice(0, 2).join(' | ')}`)
await browser.close()
process.exit(failures ? 1 : 0)
