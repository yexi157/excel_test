import { createApp } from 'vue'
import { createPinia } from 'pinia'
import App from './App.vue'
import './style.css'

async function bootstrap() {
  const useMock = import.meta.env.VITE_USE_MOCK !== 'false'
  if (useMock) {
    const { startMock } = await import('./mocks/browser')
    await startMock()
  }
  const app = createApp(App)
  app.use(createPinia())
  app.mount('#app')
}

bootstrap().catch(err => {
  console.error('[bootstrap] failed:', err)
  document.getElementById('app')!.textContent = 'Bootstrap failed: ' + err
})
