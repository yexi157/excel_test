// Vitest 全局 setup
// vitest 默认 jsdom 环境，IndexedDB 不可用 → 用 fake-indexeddb 兜底
import 'fake-indexeddb/auto'

// handlers.ts 中的 maybeFail() 会读取 localStorage.getItem('mockFailRate')。
// jsdom 自带 localStorage，但 node 环境的 spec 没有，故在 globalThis 上加最小 stub。
if (typeof globalThis.localStorage === 'undefined') {
  ;(globalThis as any).localStorage = {
    getItem: () => null,
    setItem: () => {},
    removeItem: () => {},
    clear: () => {},
    key: () => null,
    length: 0,
  }
}

// MSW 解析 handler 中的相对路径（如 '/api/files/tree'）时，会用 location.href 作为 base。
// node 环境下没有 location，需要补一个最小 stub 让 MSW 能 new URL('/api/x', location.href)。
// 用 'http://localhost' 作为默认 origin，与测试中 axios baseURL 对齐。
if (typeof (globalThis as any).location === 'undefined') {
  ;(globalThis as any).location = new URL('http://localhost/')
}
