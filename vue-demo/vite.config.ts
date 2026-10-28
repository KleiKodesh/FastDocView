import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { fileURLToPath, URL } from 'node:url'
import { spawn, type ChildProcess } from 'node:child_process'
import path from 'node:path'
import http from 'node:http'

// Like the app's dev server, this spawns its own C# host (the dev server project) and proxies
// /api to it, so `npm run dev` is the whole rig. The host is killed with vite.
const DEV_SERVER_URL = 'http://127.0.0.1:5188'
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
    {
      name: 'word-demo-dev-server',
      apply: 'serve',
      async configureServer() { await ensureDevServer() },
    },
  ],
  resolve: { alias: { '@': fileURLToPath(new URL('./src', import.meta.url)) } },
  server: {
    port: 5189,
    strictPort: true,
    proxy: { '/api': { target: DEV_SERVER_URL, changeOrigin: true } },
  },
})
