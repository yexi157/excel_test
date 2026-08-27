import { setupWorker } from 'msw/browser'
import { handlers } from './handlers'
import { seedIfEmpty, resetDB as resetDbInternal } from './seed'

const worker = setupWorker(...handlers)

export async function startMock() {
  // ?reset=1 时清空 IndexedDB（spec §7.3 失败注入）
  if (new URLSearchParams(location.search).has('reset')) {
    const { resetDB } = await import('./db')
    await resetDB()
  }
  await worker.start({
    onUnhandledRequest: 'bypass',
    serviceWorker: { url: '/mockServiceWorker.js' },
  })
  await seedIfEmpty()
}

// 防止 TS unused
void resetDbInternal
