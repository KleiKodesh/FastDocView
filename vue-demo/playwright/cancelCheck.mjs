// Live check: an open still running in the service is cancelled when the page goes away (the app's tab closed) or
// another file replaces it - the server counts the cancelled opens and keeps no document open for them.
//   node playwright/cancelCheck.mjs <slow-to-open file> <other file> [--headed]
// The slow file must take a while to open cold (a big .doc: its index pass runs inside the open).
import { chromium } from 'playwright'

const DEMO_URL = process.env.FASTDOCVIEW_DEMO_URL ?? `http://localhost:${process.env.FASTDOCVIEW_DEMO_PORT ?? 5189}/`
const [slowFile, otherFile] = process.argv.slice(2).filter(argument => !argument.startsWith('--'))
if (!slowFile || !otherFile) { console.error('usage: node playwright/cancelCheck.mjs <slow-to-open file> <other file> [--headed]'); process.exit(2) }

let failures = 0
const check = (ok, what) => { console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${what}`); if (!ok) failures++ }
const step = text => console.log(`[${new Date().toLocaleTimeString()}] ${text}`)

const browser = await chromium.launch({ headless: !process.argv.includes('--headed') })
const page = await browser.newPage({ viewport: { width: 1300, height: 900 } })
const consoleErrors = []
page.on('console', message => { if (message.type() === 'error') consoleErrors.push(message.text().slice(0, 200)) })
page.on('pageerror', error => consoleErrors.push(String(error).slice(0, 200)))
const server = async path => page.evaluate(async url => (await fetch(url)).json(), path)
const cancelled = async () => (await server('/api/open-cancels')).cancelledOpens
const openDocuments = async () => (await server('/api/memory?collect=false')).openDocuments
const waitUntil = async (predicate, timeout = 10_000) => {
  const started = Date.now()
  while (Date.now() - started < timeout) { if (await predicate()) return true; await page.waitForTimeout(100) }
  return false
}

step(`open ${DEMO_URL}`)
await page.goto(DEMO_URL, { waitUntil: 'domcontentloaded' })
await page.waitForFunction(() => !!window.__fastDocViewDemo, null, { timeout: 120_000 })
await page.waitForFunction(() => window.__fastDocViewTest?.state().isIndexComplete, null, { timeout: 120_000 }).catch(() => {})
// the page's own (first) document is open and shown before the counts start
await page.waitForFunction(() => [...document.querySelectorAll('.windowed-scroller-viewport [data-index]')].some(row => (row.textContent ?? '').trim().length > 0), null, { timeout: 120_000 })
await page.waitForTimeout(500)

// documents other runs left open on this dev server (idle-closed later) are not this check's: counts are relative
const shownPageDocuments = await openDocuments()     // includes the page's own document
console.log(`  open documents at start: ${shownPageDocuments}`)
step('tab closed while a cold open runs')
let before = await cancelled()
await page.evaluate(path => window.__fastDocViewDemo.open(path, { rebuildIndex: true }), slowFile)
await page.waitForTimeout(150)
await page.evaluate(() => window.__fastDocViewDemo.close())
check(await waitUntil(async () => (await cancelled()) > before), 'the service cancelled the open')
check(await waitUntil(async () => (await openDocuments()) === shownPageDocuments - 1), "no document stays open (the page's own closed, the cancelled one never opened)")

console.log(`  open documents after close: ${await openDocuments()}`)
step('another file opened while a cold open runs')
before = await cancelled()
await page.evaluate(path => window.__fastDocViewDemo.open(path, { rebuildIndex: true }), slowFile)
await page.waitForTimeout(150)
await page.evaluate(path => window.__fastDocViewDemo.open(path), otherFile)
check(await waitUntil(async () => (await cancelled()) > before), 'the service cancelled the replaced open')
await page.waitForFunction(() => [...document.querySelectorAll('.windowed-scroller-viewport [data-index]')].some(row => (row.textContent ?? '').trim().length > 0), null, { timeout: 60_000 })
check(true, 'the other file shows')
check(await waitUntil(async () => (await openDocuments()) === shownPageDocuments), 'only the shown document is open')
check(consoleErrors.length === 0, `no console errors${consoleErrors.length ? ': ' + consoleErrors.join(' | ') : ''}`)

await browser.close()
console.log(failures === 0 ? 'ALL PASSED' : `${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
