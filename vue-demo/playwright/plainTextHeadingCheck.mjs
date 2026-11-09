// Live check of plain .txt: no empty blocks, h1-h6 bold in size steps, gap after a heading (light + dark).
//   node playwright/plainTextHeadingCheck.mjs --list <file>      (ASCII output only)
import { chromium } from 'playwright'
import { readFileSync } from 'node:fs'
const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const args = process.argv.slice(2)
const files = readFileSync(args[args.indexOf('--list') + 1], 'utf8').split(/\r?\n/).filter(Boolean)
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
for (const file of files) {
  await page.evaluate(path => window.__fastDocViewDemo.open(path, { rebuildIndex: true }), file)
  await page.waitForFunction(() => document.querySelectorAll('.word-document .text-line > h1, .word-document .text-line > h2, .word-document .text-line > h3').length >= 1, null, 30_000)
  for (const scheme of ['light', 'dark']) {
    await page.evaluate(async preset => (await import('/@fs/' + window.__appFrontendSource + '/theme/themes.ts')).applyTheme(preset), `default-${scheme}`)
    await page.waitForTimeout(250)
    const r = await page.evaluate(() => {
      const lines = [...document.querySelectorAll('.windowed-scroller-viewport .text-line')]
      const empty = lines.filter(l => l.children.length === 1 && l.firstElementChild.tagName === 'BR' && !l.textContent.trim()).length
      const heads = [...document.querySelectorAll('.text-line > h1, .text-line > h2, .text-line > h3, .text-line > h4, .text-line > h5, .text-line > h6')]
      const sizes = {}, weights = new Set(); let maxGap = 0
      for (const h of heads) {
        const cs = getComputedStyle(h); sizes[h.tagName] = cs.fontSize; weights.add(cs.fontWeight)
        const next = h.parentElement.closest('[data-index]')?.nextElementSibling
        if (next && next.hasAttribute('data-index')) maxGap = Math.max(maxGap, next.getBoundingClientRect().top - h.getBoundingClientRect().bottom)
      }
      return { lines: lines.length, empty, heads: heads.length, sizes, weights: [...weights], maxGap: Math.round(maxGap * 10) / 10, after: heads[0] ? getComputedStyle(heads[0].parentElement, '::after').content : '' }
    })
    console.log(`  [${scheme}] ${JSON.stringify(r)}`)
    check(r.lines > 0 && r.empty === 0, `${scheme}: ${r.lines} lines, no empty blocks`)
    check(r.heads > 0 && r.weights.every(w => Number(w) >= 700), `${scheme}: headings bold`)
    const px = t => parseFloat(r.sizes[t] ?? 'NaN'); const levels = ['H1', 'H2', 'H3'].filter(t => r.sizes[t])
    check(levels.every((t, i) => i === 0 || px(levels[i - 1]) >= px(t)), `${scheme}: sizes step down (${levels.map(t => t + '=' + r.sizes[t]).join(' ')})`)
    check(r.maxGap <= 16 && (r.after === 'none' || r.after === 'normal'), `${scheme}: gap after heading <= 16px (max ${r.maxGap}), ::after ${r.after}`)
  }
}
check(errors.length === 0, `no console errors ${errors.slice(0, 2).join(' | ')}`)
await browser.close()
process.exit(failures ? 1 : 0)
