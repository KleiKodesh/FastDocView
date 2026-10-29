// The app's own FastDocView page, mounted outside the app: same components, same stores, same CSS,
// same i18n as KitveiHakodesh (vite.config.ts resolves `@/` to the app's source). Only the service
// transport is swapped (fastDocViewDemoApi.ts), so a behaviour seen here is the app's behaviour.
import { createApp } from 'vue'
import { createPinia } from 'pinia'
import { i18n, uiLanguage } from '@/i18n'
import '@/assets/styles/main.css'
import '@/assets/styles/readingColumn.css'
import { useSettingsStore } from '@/stores/settingsStore'
import { useThemeStore } from '@/theme/themeStore'
import { useTabStore } from '@/stores/tabStore'
import AppPageHost from './AppPageHost.vue'

document.documentElement.lang = uiLanguage
document.documentElement.dir = uiLanguage === 'en' ? 'ltr' : 'rtl'

const app = createApp(AppPageHost).use(createPinia()).use(i18n)
// as the app's main.ts: synchronous reads from localStorage
useSettingsStore().init()
useThemeStore().init()
useTabStore().init()
app.mount('#app')
