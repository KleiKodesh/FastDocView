// Stands in for the app's `@/webview-host/serviceClient` in the demo (vite.config.ts aliases it here). The demo has
// no KitveiHakodesh service: the FastDocView page talks to this repo's dev server through fastDocViewDemoApi.ts, and
// anything else the app's modules ask the service for (the library database, book links) fails at once with a clear
// message instead of the real client's 80 s endpoint discovery and console errors.

const NOT_AVAILABLE = 'No KitveiHakodesh service in the FastDocView demo'

export async function serviceCall<T = unknown>(op: string, _args: object = {}, _signal?: AbortSignal): Promise<T> {
  throw new Error(`${NOT_AVAILABLE} (op ${op})`)
}

export function serviceCallVoid(_op: string, _args: object = {}): void { /* nothing to call */ }

export async function* serviceStream<T = unknown>(op: string, _args: object = {}, _signal?: AbortSignal): AsyncGenerator<T, void, void> {
  throw new Error(`${NOT_AVAILABLE} (stream ${op})`)
}
