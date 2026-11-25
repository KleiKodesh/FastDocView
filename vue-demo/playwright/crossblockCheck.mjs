// Live check: web archive later-sibling pseudo-classes across blocks (css-crossblock-v1.mht vs Edge print). ASCII output.
//   node playwright/crossblockCheck.mjs <css-crossblock-v1.mht>
import { chromium } from 'playwright'
const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [mht] = process.argv.slice(2)
let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const errors = []
page.on('console', m => { if (m.type() === 'error') errors.push(m.text().slice(0, 200)) })
page.on('pageerror', e => errors.push(String(e).slice(0, 200)))
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 180_000 })
const styleOf = text => page.evaluate(t => {
  for (const root of document.querySelectorAll('.word-document')) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
    for (let node; (node = walker.nextNode());) if (node.data.includes(t)) {
      const cs = getComputedStyle(node.parentElement)
      return { weight: cs.fontWeight, style: cs.fontStyle, color: cs.color, deco: cs.textDecorationLine }
    }
  }
  return null
}, text)
for (const mode of ['rebuild', 'reopen', 'sidecar']) {
  await page.evaluate(([path, rebuild]) => window.__fastDocViewDemo.open(path, { rebuildIndex: rebuild !== 'reopen', forceSidecar: rebuild === 'sidecar' }), [mht, mode])
  await page.waitForFunction(() => [...document.querySelectorAll('.word-document')].some(e => e.textContent.includes('Last paragraph')), null, { timeout: 60_000 })
  await page.waitForTimeout(1500)
  const s = {}
  for (const t of ['Item A', 'Item B', 'Item C', 'Item D', 'Lone item', 'Only para', 'First of two', 'Second of two', 'Inner one', 'Inner2 one', 'After inner2', 'Last paragraph']) s[t] = await styleOf(t)
  console.log('  ' + mode + ' ' + JSON.stringify(s))
  const bold = x => Number(x?.weight) >= 700
  check(s['Item A'] && s['Item A'].color === 'rgb(31, 31, 31)' && !bold(s['Item A']), `${mode}: Item A plain`)
  check(s['Item B']?.deco === 'underline', `${mode}: Item B :nth-last-of-type(3) underline`)
  check(s['Item C']?.color === 'rgb(17, 34, 51)', `${mode}: Item C :nth-last-child(2) #112233`)
  check(bold(s['Item D']), `${mode}: Item D :last-of-type bold`)
  check(s['Lone item']?.color === 'rgb(0, 176, 80)', `${mode}: Lone item :only-of-type green`)
  check(s['Only para']?.color === 'rgb(68, 85, 102)' && s['Only para'].style === 'italic', `${mode}: Only para`)
  check(s['First of two']?.style === 'normal' && s['First of two'].color === 'rgb(31, 31, 31)', `${mode}: First of two plain`)
  check(s['Second of two']?.style === 'italic', `${mode}: Second of two italic`)
  check(s['Inner one']?.color === 'rgb(170, 0, 0)', `${mode}: Inner one (div.inner:last-child) #AA0000`)
  check(s['Inner2 one']?.color === 'rgb(31, 31, 31)', `${mode}: Inner2 one not :last-child`)
  check(s['After inner2']?.color === 'rgb(68, 85, 102)', `${mode}: After inner2 :only-of-type`)
  check(s['Last paragraph']?.color === 'rgb(68, 85, 102)', `${mode}: Last paragraph :only-of-type`)
}
check(errors.length === 0, 'no console errors ' + JSON.stringify(errors).replace(/[^ -~]/g, '?'))
await browser.close()
console.log(failures ? `FAILED ${failures}` : 'ALL PASSED')
process.exit(failures ? 1 : 0)
