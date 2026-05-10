import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import path from 'node:path'

// 全本地化约束（spec §5.5）：
// - 不引入任何 CDN 注入插件（vite-plugin-cdn-import 等）
// - 不在 build.rollupOptions.external 把依赖踢给 CDN
// - 所有 worker 走 new URL('./worker.ts', import.meta.url) 形式
export default defineConfig({
  plugins: [vue()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, 'src'),
    },
  },
  server: {
    port: 5173,
  },
  build: {
    outDir: 'dist',
    sourcemap: true,
  },
})
