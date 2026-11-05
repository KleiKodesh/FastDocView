// Live check of Torat Emet headings: bold, size steps, small margins, no empty block right after a heading (light + dark).
//   node playwright/toratEmetHeadingCheck.mjs --list <file>      (ASCII output only)
import { chromium } from 'playwright'
import { readFileSync } from 'node:fs'
const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const files = process.argv[2] === '--list' ? readFileSync(process.argv[3], 'utf8').split(/\r?\n/).filter(Boolean) : process.argv.slice(2)
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
    const info = await page.evaluate(() => {
      const css = (el, p) => getComputedStyle(el)[p]
      const rows = [...document.querySelectorAll('.windowed-scroller-viewport [data-index]')]
      const levels = {}
      let blankAfter = 0, headings = 0
      for (const row of rows) {
        const block = row.querySelector('.te-heading')
        if (!block) continue
        headings++
        const h = block.firstElementChild
        levels[h.tagName] = { weight: css(h, 'fontWeight'), size: css(h, 'fontSize'), margin: css(h, 'marginTop') + '/' + css(h, 'marginBottom'), color: css(h, 'color') }
        const next = row.nextElementSibling
        if (next && next.hasAttribute('data-index') && (next.textContent ?? '').replace(/[\u200b\s]/g, '') === '' && !next.querySelector('.te-meta')) blankAfter++
      }
      return { headings, levels, blankAfter, text: css(document.querySelector('.word-document .text-line:not(.te-heading):not(.te-meta)') ?? document.body, 'color') }
    })
    console.log(`  [${scheme}] headings=${info.headings} blankAfter=${info.blankAfter} ` + Object.entries(info.levels).sort().map(([t, v]) => `${t} w=${v.weight} s=${v.size} m=${v.margin}`).join(' | '))
    check(info.headings > 0 && info.blankAfter === 0, `${scheme}: no empty block directly after a heading`)
    check(Object.values(info.levels).every(v => v.weight === '700' || Number(v.weight) >= 700), `${scheme}: every heading level bold`)
    const sizes = ['H1', 'H2', 'H3', 'H4'].filter(t => info.levels[t]).map(t => parseFloat(info.levels[t].size))
    check(sizes.every((s, i) => i === 0 || s < sizes[i - 1]), `${scheme}: sizes step down (${sizes.join('>')})`)
  }
}
check(errors.length === 0, `no console errors ${errors.slice(0, 2).join(' | ')}`)
await browser.close()
process.exit(failures ? 1 : 0)
