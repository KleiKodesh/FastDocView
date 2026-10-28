// Live check of the two sample buttons: each opens its file and renders rows.
import { chromium } from 'playwright'
const browser = await chromium.launch()
const page = await browser.newPage()
const errors = []; page.on('console', m => m.type() === 'error' && errors.push(m.text()))
await page.goto('http://localhost:5189/?file=')
for (let i = 0; i < 2; i++) {
  const label = await page.locator('[data-test="sample-button"]').nth(i).textContent()
  const t = Date.now()
  await page.locator('[data-test="sample-button"]').nth(i).click()
  await page.waitForFunction(() => window.__wordDemo.state()?.document && document.querySelector('[data-block]'), null, { timeout: 120000 })
  const s = await page.evaluate(() => window.__wordDemo.state())
  console.log(`${label}: ${Date.now() - t} ms wall, ${s.document.blockCount} blocks, server ${s.document.openMilliseconds.toFixed(0)} ms`)
}
console.log('console errors:', errors.length)
await browser.close()
