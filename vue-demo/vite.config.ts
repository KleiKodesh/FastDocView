import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { fileURLToPath, URL } from 'node:url'
import { spawn, type ChildProcess } from 'node:child_process'
import path from 'node:path'
import http from 'node:http'
import { createRequire } from 'node:module'
import { pathToFileURL } from 'node:url'

// The demo runs the APP's own components: `@/` resolves to KitveiHakodesh's vue-frontend source and
// the shared packages come from the app's node_modules (one Vue / Pinia / vue-i18n instance, the
// app's versions). Only the service transport is swapped (src/appPage/fastDocViewDemoApi.ts).
const APP_FRONTEND = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../../KitveiHakodesh/vue-frontend')
const APP_SOURCE = path.join(APP_FRONTEND, 'src')
const APP_SOURCE_POSIX = APP_SOURCE.split(path.sep).join('/')
const appRequire = createRequire(path.join(APP_FRONTEND, 'package.json'))
const appPackage = (name: string) => path.join(APP_FRONTEND, 'node_modules', name)
// the app compiles its locale JSON with this plugin (runtimeOnly) - the same version, from the app
const { default: VueI18nPlugin } = await import(pathToFileURL(appRequire.resolve('@intlify/unplugin-vue-i18n/vite')).href)

// Like the app's dev server, this spawns its own C# host (the dev server project) and proxies
// /api to it, so `npm run dev` is the whole rig. The host is killed with vite.
// FASTDOCVIEW_DEV_URL / FASTDOCVIEW_DEMO_PORT let a second rig run beside the one in use (the C# host reads the same variable).
const DEV_SERVER_URL = process.env.FASTDOCVIEW_DEV_URL ?? 'http://127.0.0.1:5188'
const DEV_SERVER_PROJECT = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '../FastDocView.DevServer',
)

let devServerProcess: ChildProcess | null = null

function isDevServerUp(): Promise<boolean> {
  return new Promise(resolve => {
    const request = http.get(`${DEV_SERVER_URL}/api/memory`, response => { response.resume(); resolve(response.statusCode === 200) })
    request.on('error', () => resolve(false))
    request.setTimeout(500, () => { request.destroy(); resolve(false) })
  })
}

async function ensureDevServer() {
  if (await isDevServerUp()) { console.log('[word-demo] C# dev server already running'); return }
  console.log('[word-demo] starting C# dev server…')
  devServerProcess = spawn('dotnet', ['run', '-c', 'Release', '--project', DEV_SERVER_PROJECT], { stdio: 'inherit', shell: true })
  const stop = () => { devServerProcess?.kill(); devServerProcess = null }
  process.on('exit', stop)
  process.on('SIGINT', () => { stop(); process.exit(0) })
  process.on('SIGTERM', () => { stop(); process.exit(0) })
  for (let attempt = 0; attempt < 120; attempt++) {
    if (await isDevServerUp()) { console.log('[word-demo] C# dev server ready'); return }
    await new Promise(resolve => setTimeout(resolve, 500))
  }
  console.warn('[word-demo] C# dev server did not answer within 60 s')
}

export default defineConfig({
  plugins: [
    vue(),
    VueI18nPlugin({ include: [path.join(APP_SOURCE, 'i18n/locales/**')], runtimeOnly: true, compositionOnly: true, fullInstall: false }),
    {
      name: 'word-demo-dev-server',
      apply: 'serve',
      async configureServer() { await ensureDevServer() },
    },
  ],
  resolve: {
    alias: [
      // the one swap: the app's service transport -> this repo's dev server
      { find: /^@\/webview-host\/fastDocViewApi$/, replacement: fileURLToPath(new URL('./src/appPage/fastDocViewDemoApi.ts', import.meta.url)) },
      // no KitveiHakodesh service in the demo: the app's service client fails fast instead of discovering an endpoint
      { find: /^@\/webview-host\/serviceClient$/, replacement: fileURLToPath(new URL('./src/appPage/serviceClientDemoStub.ts', import.meta.url)) },
      { find: /^\.\/serviceClient$/, replacement: fileURLToPath(new URL('./src/appPage/serviceClientDemoStub.ts', import.meta.url)) },
      { find: /^@app-webview-host\/(.*)$/, replacement: `${APP_SOURCE_POSIX}/webview-host/$1` },
      { find: /^@\//, replacement: `${APP_SOURCE_POSIX}/` },
      ...['vue', 'pinia', 'vue-i18n', '@vueuse/core', '@iconify-prerendered/vue-fluent', '@iconify-prerendered/vue-fluent-color']
        .map(name => ({ find: new RegExp(`^${name.replace(/[/.]/g, ch => '\\' + ch)}$`), replacement: appPackage(name) })),
    ],
  },
  optimizeDeps: { include: ['@iconify-prerendered/vue-fluent', '@iconify-prerendered/vue-fluent-color'] },
  server: {
    port: Number(process.env.FASTDOCVIEW_DEMO_PORT ?? 5189),
    strictPort: true,
    proxy: { '/api': { target: DEV_SERVER_URL, changeOrigin: true } },
    fs: { allow: [path.dirname(fileURLToPath(import.meta.url)), APP_FRONTEND] },
  },
})
