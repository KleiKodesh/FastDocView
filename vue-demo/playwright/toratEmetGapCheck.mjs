// Live check of the gap after Torat Emet headings: distance from the bottom of a heading to the top of the next block's text, with every contributor (light + dark).
//   node playwright/toratEmetGapCheck.mjs [--verbose] [--bound <px>] --list <file>      (ASCII output only)
import { chromium } from 'playwright'
import { readFileSync } from 'node:fs'
const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const args = process.argv.slice(2)
const verbose = args.includes('--verbose')
const boundAt = args.indexOf('--bound'); const bound = boundAt >= 0 ? Number(args[boundAt + 1]) : 6      // heading -> text line
const boundHeading = 16                                           // heading -> heading: the next heading's own top margin is in it
const listAt = args.indexOf('--list')
const files = listAt >= 0 ? readFileSync(args[listAt + 1], 'utf8').split(/\r?\n/).filter(Boolean) : []
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
for (const [n, file] of files.entries()) {
  console.log(`== file ${n}`)
  await page.evaluate(path => window.__fastDocViewDemo.open(path, { rebuildIndex: true }), file)
  await page.waitForFunction(() => document.querySelectorAll('.word-document .te-heading').length >= 1, null, 20_000)
  for (const scheme of ['light', 'dark']) {
    await page.evaluate(async preset => (await import('/@fs/' + window.__appFrontendSource + '/theme/themes.ts')).applyTheme(preset), `default-${scheme}`)
    await page.waitForTimeout(250)
    const rows = await page.evaluate(() => {
      const px = v => Math.round(parseFloat(v) * 100) / 100
      const out = []
      const all = [...document.querySelectorAll('.windowed-scroller-viewport [data-index]')]
      for (const row of all) {
        const wrapper = row.querySelector('.te-heading')
        if (!wrapper) continue
        const next = row.nextElementSibling
        if (!next || !next.hasAttribute('data-index')) continue
        const h = wrapper.firstElementChild
        const nextLine = next.querySelector('.text-line') ?? next
        const nextIsHeading = !!next.querySelector('.te-heading')
        const nextBox = nextIsHeading ? next.querySelector('.te-heading > *') : nextLine
        const range = document.createRange(); range.selectNodeContents(nextBox)
        const glyph = [...range.getClientRects()][0]
        const cs = el => getComputedStyle(el)
        const hr = h.getBoundingClientRect(), nb = nextBox.getBoundingClientRect()
        out.push({
          tag: h.tagName, nextIsHeading, nextTag: nextBox.tagName,
          gapBox: px(nb.top - hr.bottom),                          // bottom of the heading element -> top of the next block's box
          gapText: glyph ? px(glyph.top - hr.bottom) : null,       // -> top of the next block's first text rect
          hMarginBottom: px(cs(h).marginBottom), hPaddingBottom: px(cs(h).paddingBottom), hLineHeight: cs(h).lineHeight, hFont: cs(h).fontSize,
          wrapMargin: cs(wrapper).marginTop + '/' + cs(wrapper).marginBottom, wrapPadding: cs(wrapper).paddingTop + '/' + cs(wrapper).paddingBottom,
          wrapBelowH: px(wrapper.getBoundingClientRect().bottom - hr.bottom),
          rowPadding: cs(row).paddingBottom, rowMarginBottom: cs(row).marginBottom, rowGap: cs(row.parentElement).rowGap,
          nextMargin: cs(nextLine).marginTop + '/' + cs(nextLine).marginBottom, nextPaddingTop: cs(nextLine).paddingTop, nextMinHeight: cs(nextLine).minHeight, nextLineHeight: cs(nextLine).lineHeight,
          nextRowPaddingTop: cs(next).paddingTop, nextRowMarginTop: cs(next).marginTop,
          whiteSpace: cs(wrapper).whiteSpace, hasBr: !!h.querySelector('br'), hTrailingWs: /\s$/.test(h.textContent ?? ''),
          hTrailingNode: h.lastChild ? h.lastChild.nodeName : '-',
        })
      }
      return out
    })
    const byKind = {}
    for (const r of rows) { const k = r.tag + (r.nextIsHeading ? '>head' : '>text'); (byKind[k] ??= []).push(r) }
    console.log(`  [${scheme}] pairs=${rows.length}`)
    for (const [k, list] of Object.entries(byKind).sort()) {
      const g = list.map(r => r.gapBox)
      console.log(`    ${k} n=${list.length} gapBox px min=${Math.min(...g)} max=${Math.max(...g)}`)
      if (verbose) console.log('      ' + JSON.stringify(list[0]))
    }
    check(rows.length > 0, `${scheme}: heading/next pairs found`)
    const toText = rows.filter(r => !r.nextIsHeading), toHeading = rows.filter(r => r.nextIsHeading)
    check(toText.every(r => r.gapBox <= bound), `${scheme}: heading -> text gap <= ${bound}px (max ${toText.length ? Math.max(...toText.map(r => r.gapBox)) : 'n/a'})`)
    check(toHeading.every(r => r.gapBox <= boundHeading), `${scheme}: heading -> heading gap <= ${boundHeading}px (max ${toHeading.length ? Math.max(...toHeading.map(r => r.gapBox)) : 'n/a'})`)
    check(['H1', 'H2', 'H3', 'H4'].some(t => rows.some(r => r.tag === t)), `${scheme}: heading levels measured (${[...new Set(rows.map(r => r.tag))].sort().join(',')})`)
  }
}
check(errors.length === 0, `no console errors ${errors.slice(0, 2).join(' | ')}`)
await browser.close()
process.exit(failures ? 1 : 0)
