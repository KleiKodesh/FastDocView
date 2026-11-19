// Live check: a library-app link (Otzaria, Zayit, KitveiHakodesh) clicked in a FastDocView document opens inside
// the app and never goes to the system browser. The demo has no library: a link by id opens a book tab (the app's
// tab store gets a /book-view tab); a link by title cannot be found without the catalogue, so the page says the
// book is not in the library. A web link still opens outside (window.open in the browser rig).
//   node playwright/appLinksCheck.mjs <file with the links of FastDocView.Tests/LibraryAppLinksTest> [--headed]
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const filePath = process.argv[2]
if (!filePath) { console.error('usage: node playwright/appLinksCheck.mjs <file> [--headed]'); process.exit(2) }

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }

const browser = await chromium.launch({ headless: !process.argv.includes('--headed') })
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
await page.addInitScript(() => {
  window.__opened = []
  window.open = (url) => { window.__opened.push(String(url)); return null }
})
// ?file= opens the document right away (not the big default sample first)
await page.goto(`${DEMO_URL}?file=${encodeURIComponent(filePath)}`, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
await page.waitForFunction(() => document.querySelectorAll('.windowed-scroller-viewport a[data-app-link]').length >= 9, null, { timeout: 120_000 })

const tabs = () => page.evaluate(() => {
  const pinia = document.querySelector('#app').__vue_app__.config.globalProperties.$pinia
  return pinia._s.get('tabs').tabs.map(tab => ({ route: tab.route, bookId: tab.bookId ?? null, line: tab.openTocLineIndex ?? null }))
})
// the demo has no toast banner: read the app's one toast state (the same module instance Vite serves the page)
const TOAST_MODULE = `/@fs/${(process.env.FASTDOCVIEW_APP_FRONTEND ?? '').split(String.fromCharCode(92)).join('/')}/src/composables/useToast.ts`
const toastText = () => page.evaluate(async url => { const m = await import(url); const text = m.useToast().message.value ?? ''; m.dismissToast(); return text.length }, TOAST_MODULE)

// [link label, expected: 'tab' (book id, line) | 'toast' (title, no library in the demo)]
const expectations = [
  ['applink0', { bookId: 12, line: null }],          // otzaria://open/book/12 (no index: the top)
  ['applink1', { bookId: 4, line: null }],           // zayit://book/4/line/99 (the line id needs the library: the top)
  ['applink2', { bookId: 12, line: 3 }],             // OTZARIA://open/book/12?index=3&mark
  ['applink3', { bookId: 120, line: null }],         // otzaria://open/pdf/120?index=17 (a page is not a line)
  ['applink4', 'toast'],                             // seforim://line/5 (only the library knows the line's book)
  ['applink5', { bookId: 7, line: null }],           // kitveihakodeshapp://book/7
  ['applink6', { bookId: 7, line: 2 }],              // seforimapp://book/7/?index=2
  ['applink7', 'toast'],                             // otzaria://inline-link?path=... (by title)
  ['applink8', 'toast'],                             // book://<title>#<heading> (by title)
]
for (const [label, expected] of expectations) {
  const before = JSON.stringify(await tabs())
  await toastText()
  const link = page.locator('.windowed-scroller-viewport a[data-app-link]', { hasText: new RegExp(`^${label}$`) }).first()
  const href = await link.getAttribute('href')
  await link.click()
  await page.waitForTimeout(400)
  const after = await tabs()
  const opened = await page.evaluate(() => window.__opened.length)
  if (expected === 'toast') {
    check(JSON.stringify(after) === before && opened === 0 && (await toastText()) > 0, `${label} (${href.replace(/[^\x20-\x7e]/g, '~')}): not placed without a library - the page says so, no browser`)
  } else {
    const tab = after.at(-1)
    check(JSON.stringify(after) !== before && tab.route === '/book-view' && tab.bookId === expected.bookId && tab.line === expected.line && opened === 0,
          `${label} (${href}): opens book ${expected.bookId} line ${expected.line} inside the app (${JSON.stringify(tab)}), no browser`)
  }
  // back to a neutral tab for the next click (the demo keeps one tab: openTab replaces it)
  await page.evaluate(() => {
    const store = document.querySelector('#app').__vue_app__.config.globalProperties.$pinia._s.get('tabs')
    store.updateActiveTab({ bookId: undefined, openTocLineIndex: undefined, route: '/fast-doc-view' })
  })
  await page.waitForTimeout(300)
}

const web = page.locator('.windowed-scroller-viewport a', { hasText: /^web$/ }).first()
await web.click()
await page.waitForTimeout(300)
check((await page.evaluate(() => window.__opened)).includes('https://example.org/x'), 'a web link still opens in the system browser')
check((await page.evaluate(() => window.__opened.length)) === 1, 'no app link ever reached the system browser')

await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
