// Live check of picture bullets in the app's FastDocView page, DOM measurements only:
//   node playwright/listBulletsCheck.mjs <test-data/word-lists/list-bullets-v1.docx> [.odt]
// Every picture bullet loads (naturalWidth > 0 through the host's picture endpoint), is drawn about 0.9 em of its label's
// font high at a 2:1 ratio, and no console errors appear.
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })

for (const path of process.argv.slice(2)) {
  step(`open ${path}`)
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { rebuildIndex: true }), path)
  await page.waitForFunction(() => document.querySelectorAll('img.word-picture-bullet').length >= 2, null, { timeout: 60_000 })
  await page.waitForFunction(() => [...document.querySelectorAll('img.word-picture-bullet')].every(i => i.complete), null, { timeout: 30_000 })
  const bullets = await page.evaluate(() => [...document.querySelectorAll('img.word-picture-bullet')].map(i => {
    const box = i.getBoundingClientRect(), font = parseFloat(getComputedStyle(i.parentElement).fontSize)
    const zoom = box.height / parseFloat(getComputedStyle(i).height)
    return { loaded: i.naturalWidth, ratio: box.width / box.height, emHeight: box.height / zoom / font }
  }))
  for (const [n, b] of bullets.entries()) {
    check(b.loaded > 0, `${path}: bullet ${n} picture loaded (${b.loaded} px wide)`)
    check(Math.abs(b.ratio - 2) < 0.1, `${path}: bullet ${n} ratio 2:1 (${b.ratio.toFixed(2)})`)
    check(Math.abs(b.emHeight - 0.9) < 0.05, `${path}: bullet ${n} 0.9 em high (${b.emHeight.toFixed(2)})`)
  }
}
check(errors.length === 0, `no console errors (${errors.join(' | ')})`)
await browser.close()
console.log(failures ? `${failures} FAILURE(S)` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
