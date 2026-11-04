// Live check of FastDocView's error banners on the APP's page (the demo transport, no screenshots):
//   node playwright/errorsCheck.mjs <copy-of-a-large-docx> <app-vue-frontend-folder>
// 1. an encrypted and a corrupt file show the translated message (he or en), the raw detail as the secondary line;
// 2. the open file is renamed away mid-read: jumping to an unread block shows the read banner with "try again";
//    after the file is back, the retry hides the banner and the text shows again.
import { chromium } from 'playwright'
import fs from 'node:fs'
import path from 'node:path'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [movablePath, appFrontend] = process.argv.slice(2)
if (!movablePath || !appFrontend) { console.error('usage: node playwright/errorsCheck.mjs <docx copy> <app vue-frontend>'); process.exit(2) }
const testData = path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')), '../../test-data')
const strings = lang => JSON.parse(fs.readFileSync(path.join(appFrontend, `src/i18n/locales/${lang}/fastDocView.json`), 'utf8'))
const he = strings('he'), en = strings('en')
const isText = (actual, key, code) => [he, en].some(s => (code ? s[key][code] : s[key]) === actual)

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const waitFor = (predicate, argument, timeout = 60_000) => page.waitForFunction(predicate, argument, { timeout })
const banner = () => page.evaluate(() => [...document.querySelectorAll('.message-banner--error')].map(element => ({
  message: element.querySelector('.message-banner-message')?.textContent ?? '',
  details: element.querySelector('.message-banner-details')?.textContent ?? '',
  retryOpen: !!element.querySelector('[data-fast-doc-view-retry-open]'),
  retryRead: !!element.querySelector('[data-fast-doc-view-retry-read]'),
})))

try {
  step(`open ${DEMO_URL}`)
  await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
  await waitFor(() => !!window.__fastDocViewDemo, null, 120_000)

  for (const [file, code] of [
    ['libreoffice/sw/qa/extras/ooxmlexport/data/Encrypted_MSO2007_abc.docx', 'encrypted'],
    ['apache-tika/testODTEncrypted.odt', 'encrypted'],
    ['libreoffice/sw/qa/core/data/ww8/pass/crash-2.doc', 'corrupt'],
  ]) {
    step(`open ${file}`)
    await page.evaluate(p => window.__fastDocViewDemo.open(p), path.join(testData, file))
    await waitFor(() => !!document.querySelector('.message-banner--error'), null, 30_000)
    const [shown] = await banner()
    console.log(`  message: ${shown?.message} | detail: ${shown?.details?.slice(0, 90)}`)
    check(isText(shown?.message, 'openFailure', code), `translated "${code}" message`)
    check((shown?.details ?? '').length > 0 && !shown.details.includes('FastDocViewOpenFailed'), 'raw detail as the secondary line, marker stripped')
    check(shown?.retryOpen === true, 'a "try again" button')
  }

  step(`open ${movablePath} (index mode), then rename it away mid-read`)
  await page.evaluate(p => window.__fastDocViewDemo.open(p, { forceSidecar: true }), movablePath)
  await waitFor(() => [...document.querySelectorAll('.windowed-scroller-viewport [data-index]')].some(row => (row.textContent ?? '').trim().length > 0), null, 120_000)
  await waitFor(() => window.__fastDocViewTest?.state().isIndexComplete, null, 180_000)
  const { blockCount } = await page.evaluate(() => window.__fastDocViewTest.state())
  const awayPath = movablePath + '.away'
  fs.renameSync(movablePath, awayPath)
  const target = Math.floor(blockCount * 0.7)
  await page.evaluate(i => window.__fastDocViewTest.scrollToBlock(i), target)
  await waitFor(() => !!document.querySelector('[data-fast-doc-view-retry-read]'), null, 30_000)
  const [readBanner] = await banner()
  console.log(`  message: ${readBanner?.message} | detail: ${readBanner?.details?.slice(0, 90)}`)
  check(isText(readBanner?.message, 'openFailure', 'not-found'), 'translated "not-found" message on the read banner')
  check(!!(await page.$('.windowed-scroller')), 'the document stays on screen under the banner')

  step('retry while the file is still missing: the banner stays')
  await page.click('[data-fast-doc-view-retry-read]')
  await page.waitForTimeout(1500)
  check(!!(await page.$('[data-fast-doc-view-retry-read]')), 'banner still shown')

  step('file back, retry: banner gone, text shown at the target')
  fs.renameSync(awayPath, movablePath)
  await page.click('[data-fast-doc-view-retry-read]')
  await waitFor(() => !document.querySelector('[data-fast-doc-view-retry-read]'), null, 30_000)
  await page.evaluate(i => window.__fastDocViewTest.scrollToBlock(i), target)
  await waitFor(t => [...document.querySelectorAll('.windowed-scroller-viewport [data-index]')]
    .some(row => Math.abs(Number(row.getAttribute('data-index')) - t) < 5 && (row.textContent ?? '').trim().length > 0), target, 60_000)
  check(true, 'text near the target block shown after the retry')
} catch (error) {
  console.log(`  FAIL  ${String(error).slice(0, 300)}`)
  failures++
} finally {
  if (fs.existsSync(movablePath + '.away')) fs.renameSync(movablePath + '.away', movablePath)
  await browser.close()
}
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
