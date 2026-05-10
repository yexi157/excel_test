# 在线 Excel 系统 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 实现一个浏览器内的在线 Excel 系统：左侧文件树 + 右侧 Excel 编辑器，支持 .xlsx 双向，前端完整功能 + Mock 后端。

**Architecture:** Vue 3 + Univer OSS（Excel 渲染）+ Luckyexcel（.xlsx 转换）+ Pinia（状态）+ MSW（mock）+ IndexedDB（持久化）。后端永不感知 Excel 内部结构——只交换 .xlsx 二进制和元数据 JSON。`xlsxConverter` 隔离 Luckyexcel 风险。

**Tech Stack:** Vue 3, Vite, TypeScript, Univer OSS (`@univerjs/*` + `@univerjs/ui-adapter-vue3`), Luckyexcel, element-plus, Pinia, axios, MSW, idb, Vitest

**Spec:** `docs/superpowers/specs/2026-05-10-online-excel-design.md`

**Critical gates:**
- 阶段 1 完成后必须通过 PoC 验收（spec §9）。不通过则停下来，按 R1 fallback 决策（替换 xlsxConverter 实现）后再继续阶段 2。

---

## 文件清单（plan 全部新建）

```
frontend/
├── src/
│   ├── main.ts                            # 入口 + mock 启动
│   ├── App.vue                            # 顶层布局
│   ├── style.css                          # 全局样式
│   ├── env.d.ts                           # Vite 类型声明
│   ├── components/
│   │   ├── FileTreePanel.vue
│   │   ├── EditorPanel.vue
│   │   ├── ToolbarBar.vue
│   │   └── UniverHost.vue
│   ├── stores/
│   │   └── fileStore.ts
│   ├── api/
│   │   ├── fileApi.ts
│   │   └── types.ts
│   ├── utils/
│   │   ├── xlsxConverter.ts               # 抽象接口 + 默认实现导出
│   │   ├── xlsxConverter.luckyexcel.ts    # Luckyexcel 实现
│   │   └── guardDirty.ts                  # 切换前确认弹窗
│   └── mocks/
│       ├── handlers.ts                    # MSW 7 个端点
│       ├── browser.ts                     # MSW worker 引导
│       ├── db.ts                          # IndexedDB schema
│       └── seed.ts                        # 初始数据注入
├── public/
│   ├── mockServiceWorker.js               # MSW CLI 生成
│   └── fixtures/
│       ├── simple-100x10.xlsx
│       ├── merged-cells.xlsx
│       └── formulas-sum-vlookup.xlsx
├── tests/
│   ├── setup.ts                           # Vitest 全局 setup
│   ├── xlsxConverter.spec.ts
│   ├── fileStore.spec.ts
│   ├── fileApi.spec.ts                    # 集成测试 (含 MSW)
│   └── fixtures/                          # 测试用 .xlsx（与 public/fixtures 共享）
├── index.html
├── vite.config.ts
├── vitest.config.ts
├── tsconfig.json
├── tsconfig.node.json
├── package.json
└── .gitignore
```

---

## 阶段 0：脚手架与依赖（3 任务）

### Task 0.1: 初始化 Vite + Vue 3 + TypeScript 项目

**Files:**
- Create: `frontend/package.json`
- Create: `frontend/vite.config.ts`
- Create: `frontend/tsconfig.json`
- Create: `frontend/tsconfig.node.json`
- Create: `frontend/index.html`
- Create: `frontend/src/main.ts`
- Create: `frontend/src/App.vue`
- Create: `frontend/src/env.d.ts`
- Create: `frontend/src/style.css`
- Create: `frontend/.gitignore`

- [ ] **Step 1: 创建 frontend 目录**

```bash
mkdir -p /Users/churuikai/Desktop/online_excel/frontend
cd /Users/churuikai/Desktop/online_excel/frontend
```

- [ ] **Step 2: 写 package.json（核心 scripts + 依赖占位）**

`frontend/package.json`:
```json
{
  "name": "online-excel-frontend",
  "private": true,
  "version": "0.1.0",
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "vue-tsc -b && vite build",
    "preview": "vite preview --port 4173",
    "test": "vitest run",
    "test:watch": "vitest"
  },
  "dependencies": {
    "vue": "^3.4.0",
    "pinia": "^2.1.0"
  },
  "devDependencies": {
    "@vitejs/plugin-vue": "^5.0.0",
    "typescript": "^5.4.0",
    "vue-tsc": "^2.0.0",
    "vite": "^5.2.0"
  }
}
```

- [ ] **Step 3: 写 vite.config.ts（明确禁止 CDN 注入）**

`frontend/vite.config.ts`:
```ts
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
```

- [ ] **Step 4: 写 tsconfig.json + tsconfig.node.json**

`frontend/tsconfig.json`:
```json
{
  "compilerOptions": {
    "target": "ES2022",
    "useDefineForClassFields": true,
    "module": "ESNext",
    "lib": ["ES2022", "DOM", "DOM.Iterable"],
    "skipLibCheck": true,
    "moduleResolution": "bundler",
    "allowImportingTsExtensions": false,
    "resolveJsonModule": true,
    "isolatedModules": true,
    "noEmit": true,
    "jsx": "preserve",
    "strict": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "noFallthroughCasesInSwitch": true,
    "baseUrl": ".",
    "paths": {
      "@/*": ["src/*"]
    },
    "types": ["vite/client"]
  },
  "include": ["src/**/*", "tests/**/*"],
  "references": [{ "path": "./tsconfig.node.json" }]
}
```

`frontend/tsconfig.node.json`:
```json
{
  "compilerOptions": {
    "composite": true,
    "skipLibCheck": true,
    "module": "ESNext",
    "moduleResolution": "bundler",
    "allowSyntheticDefaultImports": true,
    "strict": true
  },
  "include": ["vite.config.ts", "vitest.config.ts"]
}
```

- [ ] **Step 5: 写 index.html + main.ts + App.vue 骨架**

`frontend/index.html`:
```html
<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>在线 Excel</title>
  </head>
  <body>
    <div id="app"></div>
    <script type="module" src="/src/main.ts"></script>
  </body>
</html>
```

`frontend/src/main.ts`:
```ts
import { createApp } from 'vue'
import { createPinia } from 'pinia'
import App from './App.vue'
import './style.css'

const app = createApp(App)
app.use(createPinia())
app.mount('#app')
```

`frontend/src/App.vue`:
```vue
<script setup lang="ts">
// 顶层组件骨架，后续 Task 4.1 补充布局
</script>

<template>
  <div class="app">在线 Excel — 脚手架就绪</div>
</template>

<style scoped>
.app {
  height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
}
</style>
```

`frontend/src/env.d.ts`:
```ts
/// <reference types="vite/client" />

declare module '*.vue' {
  import type { DefineComponent } from 'vue'
  const component: DefineComponent<{}, {}, any>
  export default component
}

interface ImportMetaEnv {
  readonly VITE_USE_MOCK?: string
}

interface ImportMeta {
  readonly env: ImportMetaEnv
}
```

`frontend/src/style.css`:
```css
* { box-sizing: border-box; }
html, body, #app { margin: 0; height: 100%; font-family: system-ui, -apple-system, sans-serif; }
```

`frontend/.gitignore`:
```
node_modules/
dist/
*.log
.DS_Store
.vite/
coverage/
```

- [ ] **Step 6: 安装基础依赖并验证 dev server 启动**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm install
npm run dev
```

Expected: 浏览器打开 http://localhost:5173 显示 "在线 Excel — 脚手架就绪"。Ctrl+C 终止。

- [ ] **Step 7: 在项目根 init git + commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git init
git add docs/ frontend/
git commit -m "chore: scaffold Vue3 + Vite + TS frontend"
```

---

### Task 0.2: 安装项目核心依赖

**Files:**
- Modify: `frontend/package.json`（dependencies / devDependencies 增补）

- [ ] **Step 1: 安装 Univer OSS 全套**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm install \
  @univerjs/core \
  @univerjs/design \
  @univerjs/engine-formula \
  @univerjs/engine-render \
  @univerjs/sheets \
  @univerjs/sheets-ui \
  @univerjs/sheets-formula \
  @univerjs/sheets-formula-ui \
  @univerjs/sheets-numfmt \
  @univerjs/sheets-numfmt-ui \
  @univerjs/themes \
  @univerjs/ui \
  @univerjs/ui-adapter-vue3
```

- [ ] **Step 2: 安装 Luckyexcel + axios + element-plus + idb**

```bash
npm install luckyexcel axios element-plus @element-plus/icons-vue idb
```

- [ ] **Step 3: 安装测试与 mock 工具**

```bash
npm install -D vitest @vitest/ui jsdom @vue/test-utils msw
```

- [ ] **Step 4: 初始化 MSW service worker 文件到 public/**

```bash
npx msw init public/ --save
```

Expected: 生成 `frontend/public/mockServiceWorker.js`，并在 `package.json` 加 `"msw"` 字段记录路径。

- [ ] **Step 5: 验证安装**

```bash
ls node_modules/@univerjs/core node_modules/luckyexcel node_modules/element-plus node_modules/msw
ls public/mockServiceWorker.js
```

Expected: 所有路径都存在。

- [ ] **Step 6: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/package.json frontend/package-lock.json frontend/public/mockServiceWorker.js
git commit -m "chore: install Univer / Luckyexcel / element-plus / MSW deps"
```

---

### Task 0.3: 配置 vitest + 准备 fixtures 占位

**Files:**
- Create: `frontend/vitest.config.ts`
- Create: `frontend/tests/setup.ts`
- Create: `frontend/tests/fixtures/.gitkeep`
- Create: `frontend/public/fixtures/.gitkeep`

- [ ] **Step 1: 写 vitest.config.ts**

`frontend/vitest.config.ts`:
```ts
import { defineConfig } from 'vitest/config'
import vue from '@vitejs/plugin-vue'
import path from 'node:path'

export default defineConfig({
  plugins: [vue()],
  resolve: {
    alias: { '@': path.resolve(__dirname, 'src') },
  },
  test: {
    environment: 'jsdom',
    globals: true,
    setupFiles: ['./tests/setup.ts'],
    include: ['tests/**/*.spec.ts'],
  },
})
```

- [ ] **Step 2: 写 tests/setup.ts**

`frontend/tests/setup.ts`:
```ts
// Vitest 全局 setup
// 这里可以注册 jsdom polyfill 等。当前留空，后续按需补。
```

- [ ] **Step 3: 创建 fixtures 目录占位**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
mkdir -p tests/fixtures public/fixtures
touch tests/fixtures/.gitkeep public/fixtures/.gitkeep
```

- [ ] **Step 4: 跑一个空 vitest 验证配置**

```bash
npm test
```

Expected: 输出 "No test files found"（这是预期，配置可用）。

- [ ] **Step 5: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/vitest.config.ts frontend/tests/ frontend/public/fixtures/
git commit -m "chore: configure vitest and fixture placeholders"
```

---

## 阶段 1：PoC 验收 gate（4 任务）

### Task 1.1: 在 main.ts 集成最小 Univer + Vue3 Adapter

**Files:**
- Modify: `frontend/src/main.ts`
- Modify: `frontend/src/App.vue`
- Modify: `frontend/src/style.css`

- [ ] **Step 1: 修改 App.vue 提供全屏挂载点**

`frontend/src/App.vue`:
```vue
<script setup lang="ts">
import { onMounted, onUnmounted, ref } from 'vue'
import { LocaleType, LogLevel, Univer, UniverInstanceType } from '@univerjs/core'
import { defaultTheme } from '@univerjs/themes'
import { UniverRenderEnginePlugin } from '@univerjs/engine-render'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverUIPlugin } from '@univerjs/ui'
import { UniverVue3AdapterPlugin } from '@univerjs/ui-adapter-vue3'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import { UniverSheetsFormulaUIPlugin } from '@univerjs/sheets-formula-ui'
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt'
import { UniverSheetsNumfmtUIPlugin } from '@univerjs/sheets-numfmt-ui'
import { FUniver } from '@univerjs/core/facade'

import '@univerjs/design/lib/index.css'
import '@univerjs/ui/lib/index.css'
import '@univerjs/sheets-ui/lib/index.css'
import '@univerjs/sheets-formula-ui/lib/index.css'
import '@univerjs/sheets-numfmt-ui/lib/index.css'

// Facade side-effects
import '@univerjs/core/facade'
import '@univerjs/sheets/facade'
import '@univerjs/sheets-ui/facade'
import '@univerjs/engine-formula/facade'

const container = ref<HTMLDivElement>()
let univer: Univer | null = null

onMounted(() => {
  univer = new Univer({
    theme: defaultTheme,
    locale: LocaleType.ZH_CN,
    logLevel: LogLevel.WARN,
  })
  univer.registerPlugins([
    [UniverRenderEnginePlugin],
    [UniverFormulaEnginePlugin],
    [UniverUIPlugin, { container: container.value! }],
    [UniverVue3AdapterPlugin],
    [UniverSheetsPlugin],
    [UniverSheetsUIPlugin],
    [UniverSheetsFormulaPlugin],
    [UniverSheetsFormulaUIPlugin],
    [UniverSheetsNumfmtPlugin],
    [UniverSheetsNumfmtUIPlugin],
  ])
  univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id: 'poc-workbook',
    name: 'PoC',
    sheetOrder: ['s1'],
    sheets: { s1: { id: 's1', name: 'Sheet1', rowCount: 100, columnCount: 26 } },
  })
  ;(window as any).univerAPI = FUniver.newAPI(univer)
})

onUnmounted(() => {
  univer?.dispose()
})
</script>

<template>
  <div ref="container" class="univer-host" />
</template>

<style scoped>
.univer-host {
  width: 100vw;
  height: 100vh;
}
</style>
```

- [ ] **Step 2: 启动 dev server 验证**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm run dev
```

Expected: http://localhost:5173 显示一个空 Excel 表格，可在 cell 输入内容。Ctrl+C 终止。

- [ ] **Step 3: 浏览器 Console 验证 Facade API**

在 DevTools Console:
```js
window.univerAPI.getActiveWorkbook().getActiveSheet().getRange('A1').setValue('hello')
```

Expected: A1 显示 "hello"。

- [ ] **Step 4: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/App.vue
git commit -m "feat(poc): integrate Univer OSS with Vue 3 adapter"
```

---

### Task 1.2: 准备 PoC 用的 fixture .xlsx

**Files:**
- Create: `frontend/public/fixtures/simple-100x10.xlsx`（手动用桌面 Excel 准备）
- Create: `frontend/public/fixtures/README.md`

- [ ] **Step 1: 用桌面 Excel / WPS 创建 PoC 样本文件**

手动操作：
1. 打开桌面 Excel / WPS
2. 在 Sheet1 的 A1:J100 范围填入测试数据：A 列文本 "row-1"..."row-100"，B-J 列填随机数字
3. 在 K 列加合并单元格 + 边框
4. 保存为 `simple-100x10.xlsx` 到 `frontend/public/fixtures/`

- [ ] **Step 2: 写 fixtures README 说明**

`frontend/public/fixtures/README.md`:
```markdown
# Fixtures

测试与开发用 .xlsx 样本文件。

| 文件 | 用途 |
|---|---|
| `simple-100x10.xlsx` | 阶段 1 PoC 验收 |
| `merged-cells.xlsx` | 合并单元格 round-trip 测试 |
| `formulas-sum-vlookup.xlsx` | 公式 round-trip 测试 |

每个文件大小应 < 100KB。
```

- [ ] **Step 3: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/public/fixtures/
git commit -m "test: add PoC fixture xlsx"
```

---

### Task 1.3: PoC — 跑通 fetch → toUniver → 显示 → toXlsx → 下载

**Files:**
- Modify: `frontend/src/App.vue`

- [ ] **Step 1: 在 App.vue 增加加载 fixture + 下载按钮**

`frontend/src/App.vue`（替换 template + 增补 script）:
```vue
<script setup lang="ts">
import { onMounted, onUnmounted, ref } from 'vue'
import { LocaleType, LogLevel, Univer, UniverInstanceType } from '@univerjs/core'
import { defaultTheme } from '@univerjs/themes'
import { UniverRenderEnginePlugin } from '@univerjs/engine-render'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverUIPlugin } from '@univerjs/ui'
import { UniverVue3AdapterPlugin } from '@univerjs/ui-adapter-vue3'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import { UniverSheetsFormulaUIPlugin } from '@univerjs/sheets-formula-ui'
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt'
import { UniverSheetsNumfmtUIPlugin } from '@univerjs/sheets-numfmt-ui'
import { FUniver } from '@univerjs/core/facade'
import LuckyExcel from 'luckyexcel'

import '@univerjs/design/lib/index.css'
import '@univerjs/ui/lib/index.css'
import '@univerjs/sheets-ui/lib/index.css'
import '@univerjs/sheets-formula-ui/lib/index.css'
import '@univerjs/sheets-numfmt-ui/lib/index.css'
import '@univerjs/core/facade'
import '@univerjs/sheets/facade'
import '@univerjs/sheets-ui/facade'
import '@univerjs/engine-formula/facade'

const container = ref<HTMLDivElement>()
let univer: Univer | null = null
let univerAPI: ReturnType<typeof FUniver.newAPI> | null = null
let currentUnitId = 'poc-workbook'

function createUniver() {
  const u = new Univer({
    theme: defaultTheme,
    locale: LocaleType.ZH_CN,
    logLevel: LogLevel.WARN,
  })
  u.registerPlugins([
    [UniverRenderEnginePlugin],
    [UniverFormulaEnginePlugin],
    [UniverUIPlugin, { container: container.value! }],
    [UniverVue3AdapterPlugin],
    [UniverSheetsPlugin],
    [UniverSheetsUIPlugin],
    [UniverSheetsFormulaPlugin],
    [UniverSheetsFormulaUIPlugin],
    [UniverSheetsNumfmtPlugin],
    [UniverSheetsNumfmtUIPlugin],
  ])
  return u
}

async function loadFixture() {
  const response = await fetch('/fixtures/simple-100x10.xlsx')
  const blob = await response.blob()
  const file = new File([blob], 'simple-100x10.xlsx')
  // Luckyexcel 异步回调风格
  await new Promise<void>((resolve, reject) => {
    LuckyExcel.transformExcelToUniver(
      file,
      (workbookData: any /* IWorkbookData */, _luckysheetfile: unknown) => {
        univer?.disposeUnit(currentUnitId)
        currentUnitId = workbookData.id || 'fixture-workbook'
        univer?.createUnit(UniverInstanceType.UNIVER_SHEET, {
          ...workbookData,
          id: currentUnitId,
        })
        resolve()
      },
      (err: Error) => reject(err),
    )
  })
}

async function downloadCurrent() {
  const snapshot = univerAPI!.getActiveWorkbook()!.save()
  await new Promise<void>((resolve, reject) => {
    LuckyExcel.transformUniverToExcel(
      { snapshot, fileName: 'poc-export.xlsx' },
      () => resolve(),
      (err: Error) => reject(err),
    )
  })
}

onMounted(() => {
  univer = createUniver()
  univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id: currentUnitId,
    name: 'PoC',
    sheetOrder: ['s1'],
    sheets: { s1: { id: 's1', name: 'Sheet1', rowCount: 100, columnCount: 26 } },
  })
  univerAPI = FUniver.newAPI(univer)
  ;(window as any).univerAPI = univerAPI
})

onUnmounted(() => {
  univer?.dispose()
})
</script>

<template>
  <div class="poc-toolbar">
    <button @click="loadFixture">加载 fixture</button>
    <button @click="downloadCurrent">下载为 xlsx</button>
  </div>
  <div ref="container" class="univer-host" />
</template>

<style scoped>
.poc-toolbar {
  height: 40px;
  display: flex;
  gap: 8px;
  padding: 4px 12px;
  align-items: center;
  border-bottom: 1px solid #e0e0e0;
}
.univer-host {
  width: 100vw;
  height: calc(100vh - 40px);
}
</style>
```

- [ ] **Step 2: 启动 dev server 跑通流程**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm run dev
```

操作步骤：
1. 浏览器打开 http://localhost:5173
2. 点击「加载 fixture」按钮
3. Expected: 表格显示 simple-100x10.xlsx 内容（100 行数据 + 合并单元格 + 边框）
4. 在某个 cell 编辑：例如 K2 输入 "edited"
5. 点击「下载为 xlsx」
6. Expected: 浏览器下载 poc-export.xlsx

- [ ] **Step 3: 用桌面 Excel / WPS 打开下载的 poc-export.xlsx 验证**

Expected:
- 100 行数据完整
- 合并单元格保留
- 边框保留
- K2 显示 "edited"

- [ ] **Step 4: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/App.vue
git commit -m "feat(poc): integrate Luckyexcel for xlsx round-trip"
```

---

### Task 1.4: PoC 验收 gate

**Files:** （无文件改动，纯验收）

- [ ] **Step 1: 准备额外 2 个真实业务 .xlsx 样本**

放置到 `frontend/public/fixtures/`：
- `merged-cells.xlsx`（重度使用合并单元格）
- `formulas-sum-vlookup.xlsx`（含 SUM/VLOOKUP 公式）

```bash
ls /Users/churuikai/Desktop/online_excel/frontend/public/fixtures/
```

Expected: 至少 3 个 .xlsx 文件。

- [ ] **Step 2: 修改 App.vue 的 loadFixture，让按钮可加载任一 fixture**

把 `loadFixture` 改为接受文件名参数（替换原函数）：
```ts
async function loadFixture(name: string) {
  const response = await fetch(`/fixtures/${name}`)
  const blob = await response.blob()
  const file = new File([blob], name)
  await new Promise<void>((resolve, reject) => {
    LuckyExcel.transformExcelToUniver(
      file,
      (workbookData: any) => {
        univer?.disposeUnit(currentUnitId)
        currentUnitId = workbookData.id || 'fixture-workbook'
        univer?.createUnit(UniverInstanceType.UNIVER_SHEET, {
          ...workbookData,
          id: currentUnitId,
        })
        resolve()
      },
      (err: Error) => reject(err),
    )
  })
}
```

template 改成 3 个按钮：
```vue
<button @click="loadFixture('simple-100x10.xlsx')">加载 simple</button>
<button @click="loadFixture('merged-cells.xlsx')">加载 merged</button>
<button @click="loadFixture('formulas-sum-vlookup.xlsx')">加载 formulas</button>
<button @click="downloadCurrent">下载</button>
```

- [ ] **Step 3: 对每个 fixture 跑完整流程**

每个文件依次：
1. 加载 → 视觉验证内容正确
2. 在表格里做 1-2 处编辑
3. 下载
4. 桌面 Excel 打开下载文件 → 视觉一致 + 编辑保留

记录任何丢失/异常的特性。

- [ ] **Step 4: 验收判定**

**通过条件**：3 个 fixture 全部 round-trip 视觉一致。
- ✅ 通过：进入阶段 2
- ❌ 不通过：**停止 plan**，更新 spec 风险登记 R1，按 fallback 决策（评估 SheetJS 自写 / 后端 Python openpyxl），确认替代方案后 reset task 1.3

- [ ] **Step 5: Commit 验收结果**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/App.vue frontend/public/fixtures/
git commit -m "test(poc): pass acceptance gate with 3 fixtures"
```

---

## 阶段 2：核心抽象层 + 单测（4 任务）

### Task 2.1: 创建 api/types.ts 类型定义

**Files:**
- Create: `frontend/src/api/types.ts`

- [ ] **Step 1: 写 types.ts**

`frontend/src/api/types.ts`:
```ts
// File system 类型（spec §7.1）

export type NodeType = 'folder' | 'file'

export interface FileTreeNode {
  id: string                   // uuid v4
  name: string
  type: NodeType
  parentId: string | null
  children?: FileTreeNode[]    // 仅 folder
}

export interface FileMetadata {
  id: string
  name: string
  type: NodeType
  parentId: string | null
  size: number                 // bytes，folder 为 0
  mtime: string                // ISO 8601
  lockOwner: string | null     // 预留：未来文件锁
  lockExpires: string | null   // 预留
}

export interface ErrorBody {
  error: string                // 机器可读，如 "FILE_NOT_FOUND"
  message: string
}

// 业务异常分类
export type ApiErrorCode =
  | 'FILE_NOT_FOUND'
  | 'FILE_LOCKED'
  | 'FILE_TOO_LARGE'
  | 'NAME_CONFLICT'
  | 'INVALID_NAME'
  | 'NETWORK'
  | 'UNKNOWN'

export class ApiError extends Error {
  constructor(public code: ApiErrorCode, public httpStatus: number, message: string) {
    super(message)
    this.name = 'ApiError'
  }
}
```

- [ ] **Step 2: 类型自检**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npx vue-tsc --noEmit
```

Expected: 无错误输出（exit 0）。

- [ ] **Step 3: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/api/types.ts
git commit -m "feat(api): add type definitions for file tree and metadata"
```

---

### Task 2.2: xlsxConverter 抽象 + Luckyexcel 实现 + round-trip 单测

**Files:**
- Create: `frontend/src/utils/xlsxConverter.ts`
- Create: `frontend/src/utils/xlsxConverter.luckyexcel.ts`
- Create: `frontend/tests/xlsxConverter.spec.ts`
- Create: `frontend/tests/fixtures/`（与 public/fixtures 共享，通过 symlink）

- [ ] **Step 1: 把 public/fixtures 软链到 tests/fixtures**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend/tests
rm -rf fixtures
ln -s ../public/fixtures fixtures
ls -la fixtures
```

Expected: `fixtures -> ../public/fixtures` 软链显示。

- [ ] **Step 2: 写 xlsxConverter.ts 抽象接口**

`frontend/src/utils/xlsxConverter.ts`:
```ts
// 抽象接口：让 Luckyexcel 可替换（spec §5.3）。
// 默认导出 Luckyexcel 实现；如踩坑可换 SheetJS 自写或后端 Python，调用方零改动。

import type { IWorkbookData } from '@univerjs/core'
import { createLuckyexcelConverter } from './xlsxConverter.luckyexcel'

export interface XlsxConverter {
  /** .xlsx Blob → Univer IWorkbookData */
  toUniver(blob: Blob, fileName: string): Promise<IWorkbookData>
  /** Univer IWorkbookData → .xlsx Blob */
  toXlsx(snapshot: IWorkbookData, fileName: string): Promise<Blob>
}

export const xlsxConverter: XlsxConverter = createLuckyexcelConverter()
```

- [ ] **Step 3: 写 Luckyexcel 实现**

`frontend/src/utils/xlsxConverter.luckyexcel.ts`:
```ts
import type { IWorkbookData } from '@univerjs/core'
import LuckyExcel from 'luckyexcel'
import type { XlsxConverter } from './xlsxConverter'

export function createLuckyexcelConverter(): XlsxConverter {
  return {
    async toUniver(blob: Blob, fileName: string): Promise<IWorkbookData> {
      const file = new File([blob], fileName)
      return new Promise<IWorkbookData>((resolve, reject) => {
        LuckyExcel.transformExcelToUniver(
          file,
          (workbookData: IWorkbookData) => resolve(workbookData),
          (err: Error) => reject(err),
        )
      })
    },

    async toXlsx(snapshot: IWorkbookData, fileName: string): Promise<Blob> {
      return new Promise<Blob>((resolve, reject) => {
        LuckyExcel.transformUniverToExcel(
          {
            snapshot,
            fileName,
            getBuffer: true,                  // 让回调收到 ArrayBuffer 而非触发下载
          } as any,
          (buffer: ArrayBuffer) => {
            resolve(new Blob([buffer], {
              type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            }))
          },
          (err: Error) => reject(err),
        )
      })
    },
  }
}
```

> **注**：Luckyexcel 的 `transformUniverToExcel` 默认会触发浏览器下载。`getBuffer: true` 选项让它回调返回 ArrayBuffer，便于我们自行包成 Blob。如果运行时发现该选项不生效，task 5.5（下载）可改用 Luckyexcel 直接下载，代价是失去对 Blob 的控制——届时调整。

- [ ] **Step 4: 写 round-trip 单测**

`frontend/tests/xlsxConverter.spec.ts`:
```ts
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { xlsxConverter } from '@/utils/xlsxConverter'
import type { IWorkbookData } from '@univerjs/core'

function loadFixtureBlob(name: string): Blob {
  const buf = readFileSync(path.resolve(__dirname, 'fixtures', name))
  // jsdom 环境下 Blob 接受 BufferSource
  return new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
}

/**
 * 等价比较：忽略 mtime / Univer 内部 ID / 空 cellData 字段差异。
 * 比较关键内容：sheetOrder、sheets 名、每个 sheet 的非空 cellData。
 */
function workbookEquivalent(a: IWorkbookData, b: IWorkbookData): boolean {
  if (a.sheetOrder.length !== b.sheetOrder.length) return false
  for (let i = 0; i < a.sheetOrder.length; i++) {
    const sheetA = a.sheets[a.sheetOrder[i]]
    const sheetB = b.sheets[b.sheetOrder[i]]
    if (!sheetA || !sheetB) return false
    if (sheetA.name !== sheetB.name) return false
    if (JSON.stringify(sheetA.cellData ?? {}) !== JSON.stringify(sheetB.cellData ?? {})) return false
  }
  return true
}

describe('xlsxConverter round-trip', () => {
  const fixtures = ['simple-100x10.xlsx', 'merged-cells.xlsx', 'formulas-sum-vlookup.xlsx']

  it.each(fixtures)('%s: blob → IWorkbookData → blob 后内容等价', async (name) => {
    const original = loadFixtureBlob(name)
    const data = await xlsxConverter.toUniver(original, name)
    const exported = await xlsxConverter.toXlsx(data, name)
    const reparsed = await xlsxConverter.toUniver(exported, name)
    expect(workbookEquivalent(data, reparsed)).toBe(true)
  })
})
```

- [ ] **Step 5: 运行测试验证（先期望失败）**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm test -- tests/xlsxConverter.spec.ts
```

Expected: 测试运行；如果 Luckyexcel API 与代码假设不符或 jsdom 环境缺 polyfill，会报错——记录错误内容。

> **如果失败**：可能需要调整：
> 1. `getBuffer: true` 选项不存在 → 改用文件名约定 + 拦截下载
> 2. jsdom 缺 `URL.createObjectURL` → 在 setup.ts 加 polyfill
> 3. 调用约定与文档不符 → 看 luckyexcel npm readme 调整

修复后重新跑，直到 PASS。

- [ ] **Step 6: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/utils/ frontend/tests/xlsxConverter.spec.ts frontend/tests/fixtures
git commit -m "feat(utils): xlsxConverter abstraction with Luckyexcel impl + round-trip tests"
```

---

### Task 2.3: api/fileApi.ts 接口 + axios 实现

**Files:**
- Create: `frontend/src/api/fileApi.ts`

- [ ] **Step 1: 写 fileApi.ts**

`frontend/src/api/fileApi.ts`:
```ts
import axios, { AxiosError, type AxiosInstance } from 'axios'
import { ApiError, type FileMetadata, type FileTreeNode, type ErrorBody } from './types'

export interface FileApi {
  // 元数据 ops（JSON）
  listTree(): Promise<FileTreeNode[]>
  renameFile(id: string, newName: string): Promise<FileMetadata>
  deleteFile(id: string): Promise<void>
  createFolder(parentId: string | null, name: string): Promise<FileMetadata>

  // 内容 ops（.xlsx 二进制）
  getFile(id: string): Promise<Blob>
  saveFile(id: string, blob: Blob): Promise<void>
  uploadFile(parentId: string | null, name: string, blob: Blob): Promise<FileMetadata>
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

function wrap(err: unknown): ApiError {
  if (axios.isAxiosError(err)) {
    const ax = err as AxiosError<ErrorBody>
    const status = ax.response?.status ?? 0
    const body = ax.response?.data
    if (body?.error) {
      return new ApiError(body.error as any, status, body.message ?? ax.message)
    }
    if (status === 0) return new ApiError('NETWORK', 0, 'Network error')
    return new ApiError('UNKNOWN', status, ax.message)
  }
  return new ApiError('UNKNOWN', 0, String(err))
}

export function createFileApi(client: AxiosInstance = axios.create({ baseURL: '/api' })): FileApi {
  return {
    async listTree() {
      try {
        const res = await client.get<FileTreeNode[]>('/files/tree')
        return res.data
      } catch (e) { throw wrap(e) }
    },

    async renameFile(id, newName) {
      try {
        const res = await client.patch<FileMetadata>(`/files/${id}`, { name: newName })
        return res.data
      } catch (e) { throw wrap(e) }
    },

    async deleteFile(id) {
      try {
        await client.delete(`/files/${id}`)
      } catch (e) { throw wrap(e) }
    },

    async createFolder(parentId, name) {
      try {
        const res = await client.post<FileMetadata>('/folders', { parentId, name })
        return res.data
      } catch (e) { throw wrap(e) }
    },

    async getFile(id) {
      try {
        const res = await client.get<Blob>(`/files/${id}`, { responseType: 'blob' })
        return res.data
      } catch (e) { throw wrap(e) }
    },

    async saveFile(id, blob) {
      try {
        await client.put(`/files/${id}`, blob, {
          headers: { 'Content-Type': XLSX_MIME },
        })
      } catch (e) { throw wrap(e) }
    },

    async uploadFile(parentId, name, blob) {
      try {
        const form = new FormData()
        if (parentId) form.append('parentId', parentId)
        form.append('file', new File([blob], name, { type: XLSX_MIME }))
        const res = await client.post<FileMetadata>('/files', form, {
          headers: { 'Content-Type': 'multipart/form-data' },
        })
        return res.data
      } catch (e) { throw wrap(e) }
    },
  }
}

export const fileApi = createFileApi()
```

- [ ] **Step 2: 类型自检**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npx vue-tsc --noEmit
```

Expected: 无错误。

- [ ] **Step 3: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/api/fileApi.ts
git commit -m "feat(api): fileApi interface and axios implementation"
```

---

### Task 2.4: stores/fileStore.ts 状态机 + 行为骨架 + 单测

**Files:**
- Create: `frontend/src/stores/fileStore.ts`
- Create: `frontend/tests/fileStore.spec.ts`

- [ ] **Step 1: 写 fileStore.ts 骨架（先放占位 actions，单测覆盖状态行为）**

`frontend/src/stores/fileStore.ts`:
```ts
import { defineStore } from 'pinia'
import type { Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/facade'
import type { FileTreeNode } from '@/api/types'
import { fileApi } from '@/api/fileApi'
import { xlsxConverter } from '@/utils/xlsxConverter'

export type SaveState = 'idle' | 'saving' | 'saved' | 'error'

interface State {
  treeNodes: FileTreeNode[]
  currentFileId: string | null
  currentFileName: string | null
  dirty: boolean
  saveState: SaveState
  univerInstance: Univer | null
  univerAPI: FUniver | null
}

export const useFileStore = defineStore('file', {
  state: (): State => ({
    treeNodes: [],
    currentFileId: null,
    currentFileName: null,
    dirty: false,
    saveState: 'idle',
    univerInstance: null,
    univerAPI: null,
  }),

  actions: {
    /** 由 UniverHost 在 onMounted 调用 */
    bindUniver(instance: Univer, api: FUniver) {
      this.univerInstance = instance
      this.univerAPI = api
    },

    /** 由 UniverHost 在 commandService 监听到 mutation 时调用 */
    markDirty() {
      this.dirty = true
    },

    async refreshTree() {
      this.treeNodes = await fileApi.listTree()
    },

    async save() {
      if (!this.currentFileId || !this.univerAPI) return
      this.saveState = 'saving'
      try {
        const snapshot = this.univerAPI.getActiveWorkbook()!.save()
        const blob = await xlsxConverter.toXlsx(snapshot, this.currentFileName ?? 'untitled.xlsx')
        await fileApi.saveFile(this.currentFileId, blob)
        this.dirty = false
        this.saveState = 'saved'
        setTimeout(() => { if (this.saveState === 'saved') this.saveState = 'idle' }, 3000)
      } catch (e) {
        this.saveState = 'error'
        throw e
      }
    },

    // 后续补充：openFile (task 5.1), upload (5.4), download (5.5), createNewSheet (5.6), rename/delete/createFolder (5.7)
  },
})
```

- [ ] **Step 2: 写 fileStore 单测（覆盖状态机）**

`frontend/tests/fileStore.spec.ts`:
```ts
import { describe, it, expect, beforeEach, vi } from 'vitest'
import { setActivePinia, createPinia } from 'pinia'
import { useFileStore } from '@/stores/fileStore'

vi.mock('@/api/fileApi', () => ({
  fileApi: {
    listTree: vi.fn().mockResolvedValue([
      { id: 'root', name: 'root', type: 'folder', parentId: null, children: [] },
    ]),
    saveFile: vi.fn().mockResolvedValue(undefined),
  },
}))

vi.mock('@/utils/xlsxConverter', () => ({
  xlsxConverter: {
    toUniver: vi.fn(),
    toXlsx: vi.fn().mockResolvedValue(new Blob(['fake-xlsx'])),
  },
}))

describe('fileStore', () => {
  beforeEach(() => {
    setActivePinia(createPinia())
  })

  it('初始状态正确', () => {
    const store = useFileStore()
    expect(store.dirty).toBe(false)
    expect(store.saveState).toBe('idle')
    expect(store.currentFileId).toBeNull()
  })

  it('markDirty 把 dirty 置为 true', () => {
    const store = useFileStore()
    store.markDirty()
    expect(store.dirty).toBe(true)
  })

  it('refreshTree 拉取并存入 treeNodes', async () => {
    const store = useFileStore()
    await store.refreshTree()
    expect(store.treeNodes).toHaveLength(1)
    expect(store.treeNodes[0].name).toBe('root')
  })

  it('save 在无 currentFileId 时直接返回不抛错', async () => {
    const store = useFileStore()
    await expect(store.save()).resolves.toBeUndefined()
  })

  it('save 流程改变 saveState：idle → saving → saved', async () => {
    const store = useFileStore()
    store.currentFileId = 'f1'
    store.currentFileName = 'a.xlsx'
    store.univerAPI = {
      getActiveWorkbook: () => ({ save: () => ({ id: 'f1', sheetOrder: [], sheets: {} }) }),
    } as any

    const promise = store.save()
    expect(store.saveState).toBe('saving')
    await promise
    expect(store.saveState).toBe('saved')
    expect(store.dirty).toBe(false)
  })
})
```

- [ ] **Step 3: 运行测试**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm test -- tests/fileStore.spec.ts
```

Expected: 5 个测试 PASS。

- [ ] **Step 4: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/stores/ frontend/tests/fileStore.spec.ts
git commit -m "feat(store): fileStore skeleton with state machine + tests"
```

---

## 阶段 3：Mock 后端（5 任务）

### Task 3.1: mocks/db.ts — IndexedDB schema

**Files:**
- Create: `frontend/src/mocks/db.ts`

- [ ] **Step 1: 写 db.ts**

`frontend/src/mocks/db.ts`:
```ts
import { openDB, type DBSchema, type IDBPDatabase } from 'idb'
import type { FileMetadata } from '@/api/types'

interface OnlineExcelDB extends DBSchema {
  metadata: {
    key: string                    // file/folder id
    value: FileMetadata
    indexes: { 'by-parentId': string }
  }
  blobs: {
    key: string                    // file id（仅 type=file）
    value: Blob                    // .xlsx 二进制
  }
}

let dbPromise: Promise<IDBPDatabase<OnlineExcelDB>> | null = null

export function getDB(): Promise<IDBPDatabase<OnlineExcelDB>> {
  if (!dbPromise) {
    dbPromise = openDB<OnlineExcelDB>('online-excel-mock', 1, {
      upgrade(db) {
        const meta = db.createObjectStore('metadata', { keyPath: 'id' })
        meta.createIndex('by-parentId', 'parentId')
        db.createObjectStore('blobs')
      },
    })
  }
  return dbPromise
}

export async function resetDB() {
  const db = await getDB()
  await db.clear('metadata')
  await db.clear('blobs')
}
```

- [ ] **Step 2: 类型自检**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npx vue-tsc --noEmit
```

Expected: 无错误。

- [ ] **Step 3: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/mocks/db.ts
git commit -m "feat(mocks): IndexedDB schema for mock backend"
```

---

### Task 3.2: mocks/seed.ts — fixtures 注入

**Files:**
- Create: `frontend/src/mocks/seed.ts`

- [ ] **Step 1: 写 seed.ts**

`frontend/src/mocks/seed.ts`:
```ts
import { getDB } from './db'
import type { FileMetadata } from '@/api/types'

const SEED_FIXTURES = [
  'simple-100x10.xlsx',
  'merged-cells.xlsx',
  'formulas-sum-vlookup.xlsx',
]

function uuid(): string {
  return crypto.randomUUID()
}

function nowIso(): string {
  return new Date().toISOString()
}

export async function seedIfEmpty(): Promise<void> {
  const db = await getDB()
  const count = await db.count('metadata')
  if (count > 0) return

  // 创建一个根文件夹
  const rootId = uuid()
  const root: FileMetadata = {
    id: rootId,
    name: '示例文件夹',
    type: 'folder',
    parentId: null,
    size: 0,
    mtime: nowIso(),
    lockOwner: null,
    lockExpires: null,
  }
  await db.put('metadata', root)

  // 注入示例文件
  for (const name of SEED_FIXTURES) {
    const response = await fetch(`/fixtures/${name}`)
    if (!response.ok) {
      console.warn(`[seed] fixture ${name} fetch failed: ${response.status}`)
      continue
    }
    const blob = await response.blob()
    const id = uuid()
    const meta: FileMetadata = {
      id,
      name,
      type: 'file',
      parentId: rootId,
      size: blob.size,
      mtime: nowIso(),
      lockOwner: null,
      lockExpires: null,
    }
    await db.put('metadata', meta)
    await db.put('blobs', blob, id)
  }
}
```

- [ ] **Step 2: 类型自检**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npx vue-tsc --noEmit
```

Expected: 无错误。

- [ ] **Step 3: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/mocks/seed.ts
git commit -m "feat(mocks): seed fixtures into IndexedDB on first load"
```

---

### Task 3.3: mocks/handlers.ts — 7 个端点实现

**Files:**
- Create: `frontend/src/mocks/handlers.ts`

- [ ] **Step 1: 写 handlers.ts**

`frontend/src/mocks/handlers.ts`:
```ts
import { http, HttpResponse, delay } from 'msw'
import { getDB } from './db'
import type { FileMetadata, FileTreeNode } from '@/api/types'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const MAX_SIZE = 100 * 1024 * 1024  // 100MB

function uuid() { return crypto.randomUUID() }
function nowIso() { return new Date().toISOString() }

async function maybeFail(): Promise<Response | null> {
  await delay(50 + Math.random() * 150)
  const rate = Number(localStorage.getItem('mockFailRate') ?? '0')
  if (rate > 0 && Math.random() < rate) {
    return HttpResponse.json(
      { error: 'INJECTED_FAILURE', message: 'Random failure injection' },
      { status: 500 }
    )
  }
  return null
}

// 把扁平 metadata 列表组装成树
async function buildTree(): Promise<FileTreeNode[]> {
  const db = await getDB()
  const all = await db.getAll('metadata')
  const byId = new Map<string, FileTreeNode>()
  for (const m of all) {
    byId.set(m.id, { id: m.id, name: m.name, type: m.type, parentId: m.parentId, ...(m.type === 'folder' ? { children: [] } : {}) })
  }
  const roots: FileTreeNode[] = []
  for (const node of byId.values()) {
    if (node.parentId === null) {
      roots.push(node)
    } else {
      const parent = byId.get(node.parentId)
      if (parent && parent.children) parent.children.push(node)
    }
  }
  return roots
}

async function nameConflict(parentId: string | null, name: string, excludeId?: string): Promise<boolean> {
  const db = await getDB()
  const all = await db.getAll('metadata')
  return all.some(m => m.parentId === parentId && m.name === name && m.id !== excludeId)
}

export const handlers = [
  // 1. GET /api/files/tree
  http.get('/api/files/tree', async () => {
    const fail = await maybeFail()
    if (fail) return fail
    const tree = await buildTree()
    return HttpResponse.json(tree)
  }),

  // 2. GET /api/files/:id
  http.get('/api/files/:id', async ({ params }) => {
    const fail = await maybeFail()
    if (fail) return fail
    const db = await getDB()
    const meta = await db.get('metadata', params.id as string)
    if (!meta || meta.type !== 'file') {
      return HttpResponse.json({ error: 'FILE_NOT_FOUND', message: 'File not found' }, { status: 404 })
    }
    const blob = await db.get('blobs', params.id as string)
    if (!blob) {
      return HttpResponse.json({ error: 'FILE_NOT_FOUND', message: 'Blob missing' }, { status: 404 })
    }
    return new HttpResponse(blob, {
      status: 200,
      headers: {
        'Content-Type': XLSX_MIME,
        'Content-Disposition': `attachment; filename="${encodeURIComponent(meta.name)}"`,
      },
    })
  }),

  // 3. PUT /api/files/:id
  http.put('/api/files/:id', async ({ request, params }) => {
    const fail = await maybeFail()
    if (fail) return fail
    const db = await getDB()
    const meta = await db.get('metadata', params.id as string)
    if (!meta || meta.type !== 'file') {
      return HttpResponse.json({ error: 'FILE_NOT_FOUND', message: 'File not found' }, { status: 404 })
    }
    const blob = await request.blob()
    if (blob.size > MAX_SIZE) {
      return HttpResponse.json({ error: 'FILE_TOO_LARGE', message: 'Max 100MB' }, { status: 413 })
    }
    await db.put('blobs', blob, meta.id)
    meta.size = blob.size
    meta.mtime = nowIso()
    await db.put('metadata', meta)
    return new HttpResponse(null, { status: 204 })
  }),

  // 4. POST /api/files (multipart upload)
  http.post('/api/files', async ({ request }) => {
    const fail = await maybeFail()
    if (fail) return fail
    const form = await request.formData()
    const parentIdRaw = form.get('parentId')
    const parentId = (typeof parentIdRaw === 'string' && parentIdRaw) ? parentIdRaw : null
    const file = form.get('file')
    if (!(file instanceof File)) {
      return HttpResponse.json({ error: 'INVALID_NAME', message: 'No file in request' }, { status: 400 })
    }
    if (file.size > MAX_SIZE) {
      return HttpResponse.json({ error: 'FILE_TOO_LARGE', message: 'Max 100MB' }, { status: 413 })
    }
    if (await nameConflict(parentId, file.name)) {
      return HttpResponse.json({ error: 'NAME_CONFLICT', message: 'Name already exists' }, { status: 409 })
    }
    const id = uuid()
    const meta: FileMetadata = {
      id, name: file.name, type: 'file', parentId, size: file.size,
      mtime: nowIso(), lockOwner: null, lockExpires: null,
    }
    const db = await getDB()
    await db.put('metadata', meta)
    await db.put('blobs', file, id)
    return HttpResponse.json(meta, { status: 201 })
  }),

  // 5. POST /api/folders
  http.post('/api/folders', async ({ request }) => {
    const fail = await maybeFail()
    if (fail) return fail
    const body = await request.json() as { parentId: string | null; name: string }
    if (!body.name || /[\\/:*?"<>|]/.test(body.name)) {
      return HttpResponse.json({ error: 'INVALID_NAME', message: 'Invalid folder name' }, { status: 400 })
    }
    if (await nameConflict(body.parentId, body.name)) {
      return HttpResponse.json({ error: 'NAME_CONFLICT', message: 'Name already exists' }, { status: 409 })
    }
    const id = uuid()
    const meta: FileMetadata = {
      id, name: body.name, type: 'folder', parentId: body.parentId,
      size: 0, mtime: nowIso(), lockOwner: null, lockExpires: null,
    }
    const db = await getDB()
    await db.put('metadata', meta)
    return HttpResponse.json(meta, { status: 201 })
  }),

  // 6. PATCH /api/files/:id
  http.patch('/api/files/:id', async ({ request, params }) => {
    const fail = await maybeFail()
    if (fail) return fail
    const body = await request.json() as { name: string }
    if (!body.name || /[\\/:*?"<>|]/.test(body.name)) {
      return HttpResponse.json({ error: 'INVALID_NAME', message: 'Invalid name' }, { status: 400 })
    }
    const db = await getDB()
    const meta = await db.get('metadata', params.id as string)
    if (!meta) {
      return HttpResponse.json({ error: 'FILE_NOT_FOUND', message: 'Not found' }, { status: 404 })
    }
    if (await nameConflict(meta.parentId, body.name, meta.id)) {
      return HttpResponse.json({ error: 'NAME_CONFLICT', message: 'Name already exists' }, { status: 409 })
    }
    meta.name = body.name
    meta.mtime = nowIso()
    await db.put('metadata', meta)
    return HttpResponse.json(meta)
  }),

  // 7. DELETE /api/files/:id (递归删除 folder)
  http.delete('/api/files/:id', async ({ params }) => {
    const fail = await maybeFail()
    if (fail) return fail
    const db = await getDB()
    const root = await db.get('metadata', params.id as string)
    if (!root) {
      return HttpResponse.json({ error: 'FILE_NOT_FOUND', message: 'Not found' }, { status: 404 })
    }
    // BFS 收集所有需要删除的 id
    const toDelete: string[] = [root.id]
    const queue = [root.id]
    while (queue.length) {
      const parentId = queue.shift()!
      const all = await db.getAll('metadata')
      for (const m of all) {
        if (m.parentId === parentId) {
          toDelete.push(m.id)
          if (m.type === 'folder') queue.push(m.id)
        }
      }
    }
    for (const id of toDelete) {
      await db.delete('metadata', id)
      await db.delete('blobs', id)
    }
    return new HttpResponse(null, { status: 204 })
  }),
]
```

- [ ] **Step 2: 类型自检**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npx vue-tsc --noEmit
```

Expected: 无错误。

- [ ] **Step 3: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/mocks/handlers.ts
git commit -m "feat(mocks): implement 7 REST endpoints in MSW handlers"
```

---

### Task 3.4: mocks/browser.ts + main.ts 启动 mock

**Files:**
- Create: `frontend/src/mocks/browser.ts`
- Modify: `frontend/src/main.ts`

- [ ] **Step 1: 写 browser.ts**

`frontend/src/mocks/browser.ts`:
```ts
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
```

修复 seed.ts 没导出 resetDB：

`frontend/src/mocks/seed.ts` 末尾增补：
```ts
export { resetDB } from './db'
```

- [ ] **Step 2: 修改 main.ts 在挂载前启动 mock**

`frontend/src/main.ts`:
```ts
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
```

- [ ] **Step 3: 启动 dev server 验证 mock 注册**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm run dev
```

浏览器打开 http://localhost:5173 → DevTools Console:
```
[MSW] Mocking enabled.
```

DevTools Application → Service Workers 应能看到 mockServiceWorker.js 已注册。

DevTools Application → IndexedDB → online-excel-mock 应能看到 metadata 和 blobs store，metadata 含 1 个 folder + 3 个 file。

- [ ] **Step 4: 浏览器 Console 测试 fetch**

```js
const res = await fetch('/api/files/tree')
console.log(await res.json())
```

Expected: 返回包含 1 个根 folder（含 3 个子 file）的数组。

- [ ] **Step 5: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/mocks/browser.ts frontend/src/mocks/seed.ts frontend/src/main.ts
git commit -m "feat(mocks): bootstrap MSW worker and seed on app start"
```

---

### Task 3.5: 集成测试 fileApi × handlers

**Files:**
- Create: `frontend/tests/fileApi.spec.ts`

- [ ] **Step 1: 写 fileApi.spec.ts**

`frontend/tests/fileApi.spec.ts`:
```ts
import { describe, it, expect, beforeAll, afterAll, beforeEach } from 'vitest'
import { setupServer } from 'msw/node'
import { handlers } from '@/mocks/handlers'
import { createFileApi } from '@/api/fileApi'
import { resetDB } from '@/mocks/db'
import axios from 'axios'

// node 端 axios 需要绝对 URL
const api = createFileApi(axios.create({ baseURL: 'http://localhost/api' }))

const server = setupServer(...handlers)

beforeAll(() => server.listen({ onUnhandledRequest: 'error' }))
afterAll(() => server.close())
beforeEach(async () => {
  await resetDB()
  server.resetHandlers()
})

describe('fileApi × MSW handlers', () => {
  it('createFolder + listTree', async () => {
    const folder = await api.createFolder(null, 'docs')
    expect(folder.type).toBe('folder')
    const tree = await api.listTree()
    expect(tree).toHaveLength(1)
    expect(tree[0].name).toBe('docs')
  })

  it('uploadFile + getFile round-trip', async () => {
    const blob = new Blob(['hello world'], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const meta = await api.uploadFile(null, 'a.xlsx', blob)
    expect(meta.size).toBe(11)

    const fetched = await api.getFile(meta.id)
    expect(await fetched.text()).toBe('hello world')
  })

  it('saveFile 更新内容', async () => {
    const initial = new Blob(['v1'])
    const meta = await api.uploadFile(null, 'a.xlsx', initial)
    await api.saveFile(meta.id, new Blob(['v2-larger']))
    const fetched = await api.getFile(meta.id)
    expect(await fetched.text()).toBe('v2-larger')
  })

  it('renameFile', async () => {
    const meta = await api.uploadFile(null, 'old.xlsx', new Blob(['x']))
    const renamed = await api.renameFile(meta.id, 'new.xlsx')
    expect(renamed.name).toBe('new.xlsx')
  })

  it('deleteFile 级联删除', async () => {
    const folder = await api.createFolder(null, 'parent')
    await api.uploadFile(folder.id, 'child.xlsx', new Blob(['x']))
    await api.deleteFile(folder.id)
    const tree = await api.listTree()
    expect(tree).toHaveLength(0)
  })

  it('uploadFile 同名冲突返回 NAME_CONFLICT', async () => {
    await api.uploadFile(null, 'dup.xlsx', new Blob(['a']))
    await expect(api.uploadFile(null, 'dup.xlsx', new Blob(['b']))).rejects.toMatchObject({
      code: 'NAME_CONFLICT',
      httpStatus: 409,
    })
  })

  it('getFile 不存在返回 FILE_NOT_FOUND', async () => {
    await expect(api.getFile('nonexistent')).rejects.toMatchObject({
      code: 'FILE_NOT_FOUND',
      httpStatus: 404,
    })
  })
})
```

- [ ] **Step 2: 在 setup.ts 给 jsdom 加 IndexedDB polyfill（如缺）**

修改 `frontend/tests/setup.ts`:
```ts
// vitest 用 jsdom 环境，IndexedDB 默认不可用
// 用 fake-indexeddb 兜底
import 'fake-indexeddb/auto'
```

安装 polyfill:
```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm install -D fake-indexeddb
```

- [ ] **Step 3: 运行测试**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm test -- tests/fileApi.spec.ts
```

Expected: 7 个测试 PASS。

> **如果失败**：MSW node 模式与 axios 集成可能有问题。检查 axios baseURL 是绝对 URL；检查 MSW 版本（v2 API）。

- [ ] **Step 4: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/tests/fileApi.spec.ts frontend/tests/setup.ts frontend/package.json frontend/package-lock.json
git commit -m "test: integration tests for fileApi against MSW handlers"
```

---

## 阶段 4：UI 组件（5 任务）

### Task 4.1: App.vue 顶层布局（element-plus container）

**Files:**
- Modify: `frontend/src/App.vue`
- Modify: `frontend/src/main.ts`

- [ ] **Step 1: 在 main.ts 注册 element-plus**

`frontend/src/main.ts`（替换为）:
```ts
import { createApp } from 'vue'
import { createPinia } from 'pinia'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
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
  app.use(ElementPlus)
  app.mount('#app')
}

bootstrap().catch(err => {
  console.error('[bootstrap] failed:', err)
  document.getElementById('app')!.textContent = 'Bootstrap failed: ' + err
})
```

- [ ] **Step 2: 重写 App.vue 为左右分栏布局**

`frontend/src/App.vue`（完全替换 PoC 版本）:
```vue
<script setup lang="ts">
import { onMounted } from 'vue'
import FileTreePanel from '@/components/FileTreePanel.vue'
import EditorPanel from '@/components/EditorPanel.vue'
import { useFileStore } from '@/stores/fileStore'

const store = useFileStore()
onMounted(async () => {
  await store.refreshTree()
})
</script>

<template>
  <el-container class="app-root">
    <el-aside width="280px" class="aside">
      <FileTreePanel />
    </el-aside>
    <el-main class="main">
      <EditorPanel />
    </el-main>
  </el-container>
</template>

<style scoped>
.app-root { height: 100vh; }
.aside { border-right: 1px solid #e4e7ed; padding: 0; }
.main { padding: 0; }
</style>
```

- [ ] **Step 3: 创建空的 FileTreePanel 与 EditorPanel 占位**

`frontend/src/components/FileTreePanel.vue`:
```vue
<script setup lang="ts">
// Task 4.4 实现
</script>
<template>
  <div class="file-tree-placeholder">文件树（待实现）</div>
</template>
<style scoped>
.file-tree-placeholder { padding: 16px; color: #909399; }
</style>
```

`frontend/src/components/EditorPanel.vue`:
```vue
<script setup lang="ts">
// Task 4.5 实现
</script>
<template>
  <div class="editor-placeholder">编辑器（待实现）</div>
</template>
<style scoped>
.editor-placeholder { padding: 16px; color: #909399; }
</style>
```

- [ ] **Step 4: 启动 dev server 验证布局**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm run dev
```

Expected: 浏览器显示左侧 280px 宽的灰色 "文件树（待实现）"，右侧 "编辑器（待实现）"。

- [ ] **Step 5: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/main.ts frontend/src/App.vue frontend/src/components/
git commit -m "feat(ui): app layout with element-plus container + panel placeholders"
```

---

### Task 4.2: components/UniverHost.vue — 生命周期 + dirty 监听

**Files:**
- Create: `frontend/src/components/UniverHost.vue`
- Modify: `frontend/src/stores/fileStore.ts`（增加 ignoreInitial 控制）

- [ ] **Step 1: 在 fileStore 增加 setIgnoreInitial 控制**

修改 `frontend/src/stores/fileStore.ts` actions（在 `markDirty` 上方增加）:
```ts
    /** 切换文件时由 store 设为 true，加载完后 nextTick 设回 false */
    setIgnoreInitial(value: boolean) {
      this._ignoreInitial = value
    },

    markDirty() {
      if (this._ignoreInitial) return
      this.dirty = true
    },
```

state 增加 `_ignoreInitial`:
```ts
state: (): State => ({
  ...
  _ignoreInitial: true,    // 初始 createUnit 触发的 mutation 不算 dirty
}),
```

State interface 增加:
```ts
interface State {
  ...
  _ignoreInitial: boolean
}
```

- [ ] **Step 2: 写 UniverHost.vue**

`frontend/src/components/UniverHost.vue`:
```vue
<script setup lang="ts">
import { onMounted, onUnmounted, ref } from 'vue'
import { CommandType, ICommandService, LocaleType, LogLevel, Univer } from '@univerjs/core'
import { defaultTheme } from '@univerjs/themes'
import { UniverRenderEnginePlugin } from '@univerjs/engine-render'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverUIPlugin } from '@univerjs/ui'
import { UniverVue3AdapterPlugin } from '@univerjs/ui-adapter-vue3'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import { UniverSheetsFormulaUIPlugin } from '@univerjs/sheets-formula-ui'
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt'
import { UniverSheetsNumfmtUIPlugin } from '@univerjs/sheets-numfmt-ui'
import { FUniver } from '@univerjs/core/facade'

import '@univerjs/design/lib/index.css'
import '@univerjs/ui/lib/index.css'
import '@univerjs/sheets-ui/lib/index.css'
import '@univerjs/sheets-formula-ui/lib/index.css'
import '@univerjs/sheets-numfmt-ui/lib/index.css'

import '@univerjs/core/facade'
import '@univerjs/sheets/facade'
import '@univerjs/sheets-ui/facade'
import '@univerjs/engine-formula/facade'

import { useFileStore } from '@/stores/fileStore'

const store = useFileStore()
const container = ref<HTMLDivElement>()
let univer: Univer | null = null
let mutationDisposable: { dispose: () => void } | null = null

onMounted(() => {
  univer = new Univer({
    theme: defaultTheme,
    locale: LocaleType.ZH_CN,
    logLevel: LogLevel.WARN,
  })
  univer.registerPlugins([
    [UniverRenderEnginePlugin],
    [UniverFormulaEnginePlugin],
    [UniverUIPlugin, { container: container.value! }],
    [UniverVue3AdapterPlugin],
    [UniverSheetsPlugin],
    [UniverSheetsUIPlugin],
    [UniverSheetsFormulaPlugin],
    [UniverSheetsFormulaUIPlugin],
    [UniverSheetsNumfmtPlugin],
    [UniverSheetsNumfmtUIPlugin],
  ])

  const api = FUniver.newAPI(univer)
  store.bindUniver(univer, api)

  // 监听 mutation 设 dirty（spec §6.1）
  const commandService = (univer as any).__getInjector().get(ICommandService)
  mutationDisposable = commandService.onCommandExecuted((info: { type: number }) => {
    if (info.type === CommandType.MUTATION) {
      store.markDirty()
    }
  })
})

onUnmounted(() => {
  mutationDisposable?.dispose()
  univer?.dispose()
})
</script>

<template>
  <div ref="container" class="univer-host" />
</template>

<style scoped>
.univer-host {
  width: 100%;
  height: 100%;
}
</style>
```

- [ ] **Step 3: 类型自检**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npx vue-tsc --noEmit
```

Expected: 无错误。

- [ ] **Step 4: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/components/UniverHost.vue frontend/src/stores/fileStore.ts
git commit -m "feat(ui): UniverHost component with lifecycle and dirty detection"
```

---

### Task 4.3: components/ToolbarBar.vue

**Files:**
- Create: `frontend/src/components/ToolbarBar.vue`

- [ ] **Step 1: 写 ToolbarBar.vue**

`frontend/src/components/ToolbarBar.vue`:
```vue
<script setup lang="ts">
import { computed } from 'vue'
import { useFileStore } from '@/stores/fileStore'

const store = useFileStore()

const stateLabel = computed(() => {
  if (store.saveState === 'saving') return '保存中...'
  if (store.saveState === 'saved') return '已保存'
  if (store.saveState === 'error') return '保存失败'
  if (store.dirty) return '未保存'
  return store.currentFileId ? '已保存' : ''
})

const stateColor = computed(() => {
  if (store.saveState === 'error') return '#f56c6c'
  if (store.dirty || store.saveState === 'saving') return '#e6a23c'
  if (store.currentFileId) return '#67c23a'
  return '#909399'
})

// 占位 handlers，task 5.3 / 5.4 / 5.5 接通真实逻辑
function onSave() { /* task 5.3 */ }
function onUpload() { /* task 5.4 */ }
function onDownload() { /* task 5.5 */ }
</script>

<template>
  <div class="toolbar">
    <div class="toolbar-left">
      <span class="filename">{{ store.currentFileName || '未打开文件' }}</span>
      <span class="dot" :style="{ background: stateColor }" />
      <span class="state">{{ stateLabel }}</span>
    </div>
    <div class="toolbar-right">
      <el-button size="small" :disabled="!store.currentFileId" @click="onSave">保存</el-button>
      <el-button size="small" @click="onUpload">上传 .xlsx</el-button>
      <el-button size="small" :disabled="!store.currentFileId" @click="onDownload">下载</el-button>
    </div>
  </div>
</template>

<style scoped>
.toolbar {
  height: 48px;
  padding: 0 16px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  border-bottom: 1px solid #e4e7ed;
}
.toolbar-left { display: flex; align-items: center; gap: 8px; }
.filename { font-weight: 600; }
.dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }
.state { font-size: 12px; color: #606266; }
.toolbar-right { display: flex; gap: 8px; }
</style>
```

- [ ] **Step 2: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/components/ToolbarBar.vue
git commit -m "feat(ui): ToolbarBar with save/upload/download buttons (handlers stubbed)"
```

---

### Task 4.4: components/FileTreePanel.vue

**Files:**
- Modify: `frontend/src/components/FileTreePanel.vue`

- [ ] **Step 1: 实现 FileTreePanel.vue**

`frontend/src/components/FileTreePanel.vue`（完全替换占位）:
```vue
<script setup lang="ts">
import { computed } from 'vue'
import type { FileTreeNode } from '@/api/types'
import { useFileStore } from '@/stores/fileStore'

const store = useFileStore()
const treeData = computed(() => store.treeNodes)

const treeProps = {
  children: 'children',
  label: 'name',
}

function onNodeClick(node: FileTreeNode) {
  if (node.type !== 'file') return
  // task 5.1 实现 openFile
  console.log('[FileTreePanel] click', node.id, node.name)
}

// task 5.7 实现右键菜单（重命名/删除/新建文件夹/新建文件）
</script>

<template>
  <div class="tree-panel">
    <div class="tree-panel-header">
      <span>文件</span>
    </div>
    <el-tree
      :data="treeData"
      :props="treeProps"
      node-key="id"
      highlight-current
      @node-click="onNodeClick"
    >
      <template #default="{ node, data }">
        <span class="tree-node">
          <span :class="data.type === 'folder' ? 'icon-folder' : 'icon-file'" />
          {{ node.label }}
        </span>
      </template>
    </el-tree>
  </div>
</template>

<style scoped>
.tree-panel { height: 100%; display: flex; flex-direction: column; }
.tree-panel-header {
  height: 48px;
  padding: 0 16px;
  display: flex;
  align-items: center;
  border-bottom: 1px solid #e4e7ed;
  font-weight: 600;
}
.tree-node { display: inline-flex; align-items: center; gap: 6px; }
.icon-folder::before { content: '📁'; }
.icon-file::before { content: '📄'; }
</style>
```

- [ ] **Step 2: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/components/FileTreePanel.vue
git commit -m "feat(ui): FileTreePanel with el-tree (click handler stubbed)"
```

---

### Task 4.5: EditorPanel 整合 ToolbarBar + UniverHost

**Files:**
- Modify: `frontend/src/components/EditorPanel.vue`

- [ ] **Step 1: 实现 EditorPanel.vue**

`frontend/src/components/EditorPanel.vue`（完全替换占位）:
```vue
<script setup lang="ts">
import ToolbarBar from './ToolbarBar.vue'
import UniverHost from './UniverHost.vue'
</script>

<template>
  <div class="editor-panel">
    <ToolbarBar />
    <div class="editor-host">
      <UniverHost />
    </div>
  </div>
</template>

<style scoped>
.editor-panel { height: 100%; display: flex; flex-direction: column; }
.editor-host { flex: 1; min-height: 0; }
</style>
```

- [ ] **Step 2: 启动 dev server 验证整体布局**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm run dev
```

Expected:
- 左侧：文件树显示 "示例文件夹" 节点（含 3 个 .xlsx 子项）
- 右侧顶部：toolbar（"未打开文件" + 灰色圆点 + 三个按钮）
- 右侧主体：空 Univer Excel 渲染（默认空 workbook）

- [ ] **Step 3: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/components/EditorPanel.vue
git commit -m "feat(ui): EditorPanel integrates ToolbarBar and UniverHost"
```

---

## 阶段 5：场景实现（8 任务）

### Task 5.1: 场景 1 — 打开文件 + guardDirty

**Files:**
- Create: `frontend/src/utils/guardDirty.ts`
- Modify: `frontend/src/stores/fileStore.ts`
- Modify: `frontend/src/components/FileTreePanel.vue`

- [ ] **Step 1: 写 guardDirty.ts**

`frontend/src/utils/guardDirty.ts`:
```ts
import { ElMessageBox } from 'element-plus'
import { useFileStore } from '@/stores/fileStore'

export type GuardResult = 'continue' | 'cancel'

/**
 * 切换/关闭前的 dirty 守卫（spec §6.2）
 * 返回 'continue' 表示用户已处理（保存或放弃），可以继续
 * 返回 'cancel' 表示用户取消，应中止当前操作
 */
export async function guardDirty(): Promise<GuardResult> {
  const store = useFileStore()
  if (!store.dirty) return 'continue'

  try {
    const action = await ElMessageBox.confirm(
      `当前文件「${store.currentFileName ?? ''}」有未保存的改动`,
      '继续操作？',
      {
        distinguishCancelAndClose: true,
        confirmButtonText: '保存',
        cancelButtonText: '放弃',
        type: 'warning',
      },
    )
    if (action === 'confirm') {
      await store.save()
      return 'continue'
    }
    return 'continue'
  } catch (action) {
    if (action === 'cancel') return 'continue'   // 放弃改动
    return 'cancel'                              // close / esc
  }
}
```

> ElMessageBox 的"确认"resolve 为 `'confirm'`，"取消"reject 为 `'cancel'`，"关闭"reject 为 `'close'`。

- [ ] **Step 2: 在 fileStore 增加 openFile action**

修改 `frontend/src/stores/fileStore.ts` actions 增加（位置在 `save` 上方）:
```ts
    async openFile(id: string) {
      const { guardDirty } = await import('@/utils/guardDirty')
      if ((await guardDirty()) === 'cancel') return
      if (!this.univerInstance) throw new Error('Univer not bound')

      const blob = await fileApi.getFile(id)
      // 找文件名
      const found = findNodeById(this.treeNodes, id)
      const fileName = found?.name ?? 'untitled.xlsx'
      const workbookData = await xlsxConverter.toUniver(blob, fileName)

      if (this.currentFileId) {
        this.univerInstance.disposeUnit(this.currentFileId)
      }
      this.setIgnoreInitial(true)

      const { UniverInstanceType } = await import('@univerjs/core')
      this.univerInstance.createUnit(UniverInstanceType.UNIVER_SHEET, {
        ...workbookData,
        id,
      })

      this.currentFileId = id
      this.currentFileName = fileName
      this.dirty = false
      this.saveState = 'idle'

      // nextTick 后允许后续 mutation 触发 dirty
      await new Promise<void>(resolve => setTimeout(resolve, 0))
      this.setIgnoreInitial(false)
    },
```

在文件顶部 import 增加：
```ts
import type { FileTreeNode } from '@/api/types'
```

文件末尾增加辅助函数（在 store 之外）:
```ts
function findNodeById(nodes: FileTreeNode[], id: string): FileTreeNode | undefined {
  for (const n of nodes) {
    if (n.id === id) return n
    if (n.children) {
      const hit = findNodeById(n.children, id)
      if (hit) return hit
    }
  }
  return undefined
}
```

- [ ] **Step 3: 在 FileTreePanel 调 openFile**

修改 `frontend/src/components/FileTreePanel.vue` 的 `onNodeClick`:
```ts
async function onNodeClick(node: FileTreeNode) {
  if (node.type !== 'file') return
  try {
    await store.openFile(node.id)
  } catch (e) {
    console.error('[openFile] failed:', e)
  }
}
```

- [ ] **Step 4: 增加 fileStore 单测覆盖 openFile**

修改 `frontend/tests/fileStore.spec.ts` 在末尾增加：
```ts
import { UniverInstanceType } from '@univerjs/core'

  it('openFile 完成后状态正确', async () => {
    const store = useFileStore()
    store.treeNodes = [
      { id: 'f1', name: 'a.xlsx', type: 'file', parentId: null },
    ]
    const fakeUniver = {
      disposeUnit: vi.fn(),
      createUnit: vi.fn(),
    } as any
    const fakeAPI = {} as any
    store.bindUniver(fakeUniver, fakeAPI)

    // mock fileApi.getFile + xlsxConverter.toUniver
    const { fileApi } = await import('@/api/fileApi')
    const { xlsxConverter } = await import('@/utils/xlsxConverter')
    vi.mocked(fileApi.getFile = vi.fn().mockResolvedValue(new Blob(['x'])))
    vi.mocked(xlsxConverter.toUniver = vi.fn().mockResolvedValue({ id: 'f1', sheetOrder: ['s1'], sheets: { s1: {} } }))

    await store.openFile('f1')
    expect(store.currentFileId).toBe('f1')
    expect(store.currentFileName).toBe('a.xlsx')
    expect(store.dirty).toBe(false)
    expect(fakeUniver.createUnit).toHaveBeenCalledWith(UniverInstanceType.UNIVER_SHEET, expect.objectContaining({ id: 'f1' }))
  })
```

- [ ] **Step 5: 运行测试**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm test -- tests/fileStore.spec.ts
```

Expected: 全部 PASS。

- [ ] **Step 6: 浏览器验证**

```bash
npm run dev
```

操作：
1. 浏览器打开
2. 点击左侧 "示例文件夹" 展开
3. 点击 simple-100x10.xlsx
4. Expected: 右侧编辑器加载并显示文件内容；toolbar 显示文件名 + 绿色圆点

- [ ] **Step 7: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/utils/guardDirty.ts frontend/src/stores/fileStore.ts frontend/src/components/FileTreePanel.vue frontend/tests/fileStore.spec.ts
git commit -m "feat(scenarios): scenario 1 - open file with dirty guard"
```

---

### Task 5.2: 场景 2 — 编辑触发 dirty（验收）

**Files:** （无文件改动，纯手动验证 + 自动测试加固）

- [ ] **Step 1: 浏览器手动验证**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm run dev
```

操作：
1. 打开任一 .xlsx
2. Toolbar 应显示绿色 "已保存"
3. 在某 cell 输入新值（按回车确认）
4. Expected: toolbar 圆点变橙色，状态显示 "未保存"

- [ ] **Step 2: 切换文件验证 guardDirty**

操作（继续上面状态）：
1. 不保存，点击树里另一个 .xlsx
2. Expected: 弹出确认框「当前文件 ... 有未保存的改动」三选项 [保存][放弃][取消]
3. 点击 "取消" → 不切换
4. 再次点击另一文件 → 这次点 "放弃" → 切换成功，新文件加载，dirty=false

- [ ] **Step 3: 给 fileStore 测试加 dirty 行为覆盖**

修改 `frontend/tests/fileStore.spec.ts` 末尾增加：
```ts
  it('在 ignoreInitial=true 时调 markDirty 不改变 dirty', () => {
    const store = useFileStore()
    store.setIgnoreInitial(true)
    store.markDirty()
    expect(store.dirty).toBe(false)
  })

  it('ignoreInitial=false 后调 markDirty 把 dirty 置为 true', () => {
    const store = useFileStore()
    store.setIgnoreInitial(false)
    store.markDirty()
    expect(store.dirty).toBe(true)
  })
```

- [ ] **Step 4: 运行测试 + commit**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm test
```

Expected: 全部 PASS。

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/tests/fileStore.spec.ts
git commit -m "test: cover ignoreInitial behavior in fileStore"
```

---

### Task 5.3: 场景 3 — 保存（手动 + Ctrl+S）

**Files:**
- Modify: `frontend/src/components/ToolbarBar.vue`
- Modify: `frontend/src/components/UniverHost.vue`

- [ ] **Step 1: 在 ToolbarBar 接通 onSave + ElMessage 反馈**

修改 `frontend/src/components/ToolbarBar.vue` 的 `<script setup>`:
```ts
import { ElMessage } from 'element-plus'

async function onSave() {
  try {
    await store.save()
    ElMessage.success('保存成功')
  } catch (e: any) {
    ElMessage.error(`保存失败：${e.message ?? e}`)
  }
}
```

- [ ] **Step 2: 在 UniverHost 注册 Ctrl+S 全局快捷键**

修改 `frontend/src/components/UniverHost.vue` 的 `onMounted` 末尾增加 + `onUnmounted` 清理:
```ts
function onKeydown(e: KeyboardEvent) {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
    e.preventDefault()
    if (store.currentFileId) {
      store.save().catch(err => console.error('[Ctrl+S save] failed:', err))
    }
  }
}

// onMounted 末尾：
window.addEventListener('keydown', onKeydown)

// onUnmounted：
window.removeEventListener('keydown', onKeydown)
```

> 注：在 UniverHost 而非 ToolbarBar 注册是因为 UniverHost 必然挂载；Ctrl+S 调 store.save() 直接走 store 流程。

- [ ] **Step 3: 浏览器验证**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm run dev
```

操作：
1. 打开文件 → 编辑 cell（dirty=true）
2. 点击保存按钮 → 看到 toast "保存成功"，toolbar 状态变绿
3. 再编辑 → 按 Ctrl+S → 同样保存
4. 刷新页面 → 重新打开同文件 → 上次编辑的内容仍在

- [ ] **Step 4: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/components/ToolbarBar.vue frontend/src/components/UniverHost.vue
git commit -m "feat(scenarios): scenario 3 - manual save + Ctrl+S shortcut"
```

---

### Task 5.4: 场景 4 — 上传 .xlsx

**Files:**
- Modify: `frontend/src/stores/fileStore.ts`
- Modify: `frontend/src/components/ToolbarBar.vue`

- [ ] **Step 1: 在 fileStore 增加 upload action**

修改 `frontend/src/stores/fileStore.ts` actions 增加（位置在 `openFile` 上方）:
```ts
    async upload(parentId: string | null, file: File) {
      if (!file.name.toLowerCase().endsWith('.xlsx')) {
        throw new Error('只支持 .xlsx 文件')
      }
      if (file.size > 100 * 1024 * 1024) {
        throw new Error('文件超过 100MB')
      }
      const meta = await fileApi.uploadFile(parentId, file.name, file)
      await this.refreshTree()
      await this.openFile(meta.id)
    },
```

- [ ] **Step 2: 在 ToolbarBar 实现 onUpload + 隐藏 input**

修改 `frontend/src/components/ToolbarBar.vue` 完整版（增加上传逻辑）:
```vue
<script setup lang="ts">
import { computed, ref } from 'vue'
import { ElMessage } from 'element-plus'
import { useFileStore } from '@/stores/fileStore'

const store = useFileStore()
const fileInput = ref<HTMLInputElement>()

const stateLabel = computed(() => {
  if (store.saveState === 'saving') return '保存中...'
  if (store.saveState === 'saved') return '已保存'
  if (store.saveState === 'error') return '保存失败'
  if (store.dirty) return '未保存'
  return store.currentFileId ? '已保存' : ''
})

const stateColor = computed(() => {
  if (store.saveState === 'error') return '#f56c6c'
  if (store.dirty || store.saveState === 'saving') return '#e6a23c'
  if (store.currentFileId) return '#67c23a'
  return '#909399'
})

async function onSave() {
  try {
    await store.save()
    ElMessage.success('保存成功')
  } catch (e: any) {
    ElMessage.error(`保存失败：${e.message ?? e}`)
  }
}

function onUpload() {
  fileInput.value?.click()
}

async function onFileChosen(e: Event) {
  const target = e.target as HTMLInputElement
  const file = target.files?.[0]
  if (!file) return
  try {
    // 上传到根目录（task 5.7 增加上传到指定文件夹）
    await store.upload(null, file)
    ElMessage.success('上传成功')
  } catch (err: any) {
    ElMessage.error(`上传失败：${err.message ?? err}`)
  } finally {
    target.value = ''   // reset 让同一文件可重复选
  }
}

function onDownload() { /* task 5.5 */ }
</script>

<template>
  <div class="toolbar">
    <div class="toolbar-left">
      <span class="filename">{{ store.currentFileName || '未打开文件' }}</span>
      <span class="dot" :style="{ background: stateColor }" />
      <span class="state">{{ stateLabel }}</span>
    </div>
    <div class="toolbar-right">
      <el-button size="small" :disabled="!store.currentFileId" @click="onSave">保存</el-button>
      <el-button size="small" @click="onUpload">上传 .xlsx</el-button>
      <el-button size="small" :disabled="!store.currentFileId" @click="onDownload">下载</el-button>
      <input
        ref="fileInput"
        type="file"
        accept=".xlsx"
        style="display: none"
        @change="onFileChosen"
      />
    </div>
  </div>
</template>

<style scoped>
.toolbar { height: 48px; padding: 0 16px; display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid #e4e7ed; }
.toolbar-left { display: flex; align-items: center; gap: 8px; }
.filename { font-weight: 600; }
.dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }
.state { font-size: 12px; color: #606266; }
.toolbar-right { display: flex; gap: 8px; }
</style>
```

- [ ] **Step 3: 浏览器验证**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm run dev
```

操作：
1. 准备一个 .xlsx 文件在桌面
2. 点击「上传 .xlsx」按钮 → 选文件
3. Expected: toast "上传成功"；左侧树新增节点；右侧自动打开新文件
4. 再上传同名文件 → toast "上传失败：Name already exists"

- [ ] **Step 4: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/stores/fileStore.ts frontend/src/components/ToolbarBar.vue
git commit -m "feat(scenarios): scenario 4 - upload xlsx"
```

---

### Task 5.5: 场景 5 — 下载 .xlsx

**Files:**
- Modify: `frontend/src/stores/fileStore.ts`
- Modify: `frontend/src/components/ToolbarBar.vue`

- [ ] **Step 1: 在 fileStore 增加 download action**

修改 `frontend/src/stores/fileStore.ts` actions 增加（位置在 `upload` 上方）:
```ts
    async download(): Promise<'done' | 'cancel'> {
      if (!this.currentFileId) return 'done'

      if (this.dirty) {
        const { ElMessageBox } = await import('element-plus')
        try {
          await ElMessageBox.confirm(
            '当前有未保存改动。下载的将是后端最新已保存版本。',
            '继续下载？',
            {
              distinguishCancelAndClose: true,
              confirmButtonText: '先保存再下载',
              cancelButtonText: '直接下载已保存版本',
              type: 'warning',
            },
          )
          // 用户选先保存
          await this.save()
        } catch (action) {
          if (action === 'close') return 'cancel'   // X 关闭 = 取消
          // 'cancel' 走"直接下载"路径
        }
      }

      const blob = await fileApi.getFile(this.currentFileId)
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = this.currentFileName ?? 'download.xlsx'
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
      URL.revokeObjectURL(url)
      return 'done'
    },
```

- [ ] **Step 2: 在 ToolbarBar 接通 onDownload**

修改 `frontend/src/components/ToolbarBar.vue` 的 `onDownload`:
```ts
async function onDownload() {
  try {
    await store.download()
  } catch (e: any) {
    ElMessage.error(`下载失败：${e.message ?? e}`)
  }
}
```

- [ ] **Step 3: 浏览器验证**

操作：
1. 打开文件，不编辑 → 点下载 → Expected: 浏览器下载文件，与原文件一致
2. 编辑后不保存 → 点下载 → 弹"未保存改动"对话框
3. 选"先保存再下载" → 保存成功 + 下载 → 下载文件包含编辑
4. 编辑后 → 下载 → 选"直接下载已保存版本" → 下载的不含本次编辑

- [ ] **Step 4: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/stores/fileStore.ts frontend/src/components/ToolbarBar.vue
git commit -m "feat(scenarios): scenario 5 - download xlsx with dirty prompt"
```

---

### Task 5.6: 场景 6 — 新建空文件

**Files:**
- Modify: `frontend/src/stores/fileStore.ts`
- Modify: `frontend/src/components/FileTreePanel.vue`

- [ ] **Step 1: 在 fileStore 增加 createNewSheet**

修改 `frontend/src/stores/fileStore.ts` actions 增加（位置在 `upload` 上方）:
```ts
    async createNewSheet(parentId: string | null, name: string) {
      const finalName = name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`
      const empty = {
        id: 'pending',           // 后端会用 uploadFile 生成新 id
        sheetOrder: ['s1'],
        sheets: {
          s1: { id: 's1', name: 'Sheet1', rowCount: 100, columnCount: 26 },
        },
        styles: {},
        appVersion: '0.21.0',
        locale: 'zhCN',
      }
      const blob = await xlsxConverter.toXlsx(empty as any, finalName)
      const meta = await fileApi.uploadFile(parentId, finalName, blob)
      await this.refreshTree()
      await this.openFile(meta.id)
    },
```

- [ ] **Step 2: 在 FileTreePanel 增加右键菜单 - "新建文件"**

修改 `frontend/src/components/FileTreePanel.vue`（完全替换）:
```vue
<script setup lang="ts">
import { computed, ref } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import type { FileTreeNode } from '@/api/types'
import { useFileStore } from '@/stores/fileStore'

const store = useFileStore()
const treeData = computed(() => store.treeNodes)
const treeProps = { children: 'children', label: 'name' }

async function onNodeClick(node: FileTreeNode) {
  if (node.type !== 'file') return
  try {
    await store.openFile(node.id)
  } catch (e) {
    console.error('[openFile] failed:', e)
  }
}

// 新建文件（在选中文件夹下；如果选中是文件，则在其父文件夹下）
async function newFile(parentNode: FileTreeNode | null) {
  const parentId = parentNode?.type === 'folder' ? parentNode.id : (parentNode?.parentId ?? null)
  try {
    const { value } = await ElMessageBox.prompt('文件名（自动加 .xlsx 后缀）', '新建文件', {
      inputPattern: /^[^\\/:*?"<>|]+$/,
      inputErrorMessage: '名称含非法字符',
    })
    await store.createNewSheet(parentId, value)
    ElMessage.success('已创建')
  } catch (e: any) {
    if (e === 'cancel' || e === 'close') return
    ElMessage.error(`创建失败：${e.message ?? e}`)
  }
}

const contextNode = ref<FileTreeNode | null>(null)
const menuVisible = ref(false)
const menuStyle = ref<{ top: string; left: string }>({ top: '0px', left: '0px' })

function onContext(e: MouseEvent, _data: FileTreeNode, node: any) {
  e.preventDefault()
  contextNode.value = node.data
  menuStyle.value = { top: `${e.clientY}px`, left: `${e.clientX}px` }
  menuVisible.value = true
}

function closeMenu() { menuVisible.value = false }

function newFileFromContext() {
  closeMenu()
  newFile(contextNode.value)
}
</script>

<template>
  <div class="tree-panel" @click="closeMenu">
    <div class="tree-panel-header">
      <span>文件</span>
      <el-button size="small" text @click="newFile(null)">+ 新建</el-button>
    </div>
    <el-tree
      :data="treeData"
      :props="treeProps"
      node-key="id"
      highlight-current
      @node-click="onNodeClick"
      @node-contextmenu="onContext"
    >
      <template #default="{ node, data }">
        <span class="tree-node">
          <span :class="data.type === 'folder' ? 'icon-folder' : 'icon-file'" />
          {{ node.label }}
        </span>
      </template>
    </el-tree>

    <div v-if="menuVisible" class="ctx-menu" :style="menuStyle" @click.stop>
      <div class="ctx-item" @click="newFileFromContext">新建文件</div>
      <!-- task 5.7 增加：重命名 / 删除 / 新建文件夹 -->
    </div>
  </div>
</template>

<style scoped>
.tree-panel { height: 100%; display: flex; flex-direction: column; position: relative; }
.tree-panel-header {
  height: 48px; padding: 0 16px;
  display: flex; align-items: center; justify-content: space-between;
  border-bottom: 1px solid #e4e7ed; font-weight: 600;
}
.tree-node { display: inline-flex; align-items: center; gap: 6px; }
.icon-folder::before { content: '📁'; }
.icon-file::before { content: '📄'; }
.ctx-menu {
  position: fixed; background: white; border: 1px solid #e4e7ed; border-radius: 4px;
  box-shadow: 0 2px 12px rgba(0,0,0,0.1); z-index: 9999; min-width: 120px;
}
.ctx-item { padding: 8px 16px; cursor: pointer; }
.ctx-item:hover { background: #f5f7fa; }
</style>
```

- [ ] **Step 3: 浏览器验证**

操作：
1. 点击右上角 "+ 新建" 按钮 → 输入 "test"
2. Expected: 树中新增 "test.xlsx"，右侧自动打开（空表）
3. 在某 cell 输入内容 → 保存
4. 刷新页面 → 重新打开 test.xlsx → 内容仍在
5. 右键某文件夹节点 → 选 "新建文件" → 输入名称 → 在该文件夹下创建

- [ ] **Step 4: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/stores/fileStore.ts frontend/src/components/FileTreePanel.vue
git commit -m "feat(scenarios): scenario 6 - create new empty xlsx file"
```

---

### Task 5.7: 重命名 / 删除 / 新建文件夹

**Files:**
- Modify: `frontend/src/stores/fileStore.ts`
- Modify: `frontend/src/components/FileTreePanel.vue`

- [ ] **Step 1: 在 fileStore 增加 rename / remove / createFolder action**

修改 `frontend/src/stores/fileStore.ts` actions 增加（位置在 `createNewSheet` 上方）:
```ts
    async createFolder(parentId: string | null, name: string) {
      await fileApi.createFolder(parentId, name)
      await this.refreshTree()
    },

    async renameNode(id: string, newName: string) {
      await fileApi.renameFile(id, newName)
      if (this.currentFileId === id) {
        this.currentFileName = newName
      }
      await this.refreshTree()
    },

    async removeNode(id: string) {
      await fileApi.deleteFile(id)
      if (this.currentFileId === id) {
        if (this.univerInstance) this.univerInstance.disposeUnit(id)
        this.currentFileId = null
        this.currentFileName = null
        this.dirty = false
      }
      await this.refreshTree()
    },
```

- [ ] **Step 2: 扩展 FileTreePanel 右键菜单**

修改 `frontend/src/components/FileTreePanel.vue` 的 `<script setup>` 末尾增加:
```ts
async function renameFromContext() {
  closeMenu()
  if (!contextNode.value) return
  try {
    const { value } = await ElMessageBox.prompt('新名称', '重命名', {
      inputValue: contextNode.value.name,
      inputPattern: /^[^\\/:*?"<>|]+$/,
      inputErrorMessage: '名称含非法字符',
    })
    await store.renameNode(contextNode.value.id, value)
    ElMessage.success('已重命名')
  } catch (e: any) {
    if (e === 'cancel' || e === 'close') return
    ElMessage.error(`重命名失败：${e.message ?? e}`)
  }
}

async function removeFromContext() {
  closeMenu()
  if (!contextNode.value) return
  try {
    await ElMessageBox.confirm(
      `确认删除「${contextNode.value.name}」${contextNode.value.type === 'folder' ? '及其全部内容' : ''}？`,
      '危险操作',
      { type: 'warning' },
    )
    await store.removeNode(contextNode.value.id)
    ElMessage.success('已删除')
  } catch (e: any) {
    if (e === 'cancel' || e === 'close') return
    ElMessage.error(`删除失败：${e.message ?? e}`)
  }
}

async function newFolderFromContext() {
  closeMenu()
  const parentId = contextNode.value?.type === 'folder'
    ? contextNode.value.id
    : (contextNode.value?.parentId ?? null)
  try {
    const { value } = await ElMessageBox.prompt('文件夹名', '新建文件夹', {
      inputPattern: /^[^\\/:*?"<>|]+$/,
      inputErrorMessage: '名称含非法字符',
    })
    await store.createFolder(parentId, value)
    ElMessage.success('已创建')
  } catch (e: any) {
    if (e === 'cancel' || e === 'close') return
    ElMessage.error(`创建失败：${e.message ?? e}`)
  }
}
```

修改 `<template>` 中右键菜单部分：
```vue
    <div v-if="menuVisible" class="ctx-menu" :style="menuStyle" @click.stop>
      <div class="ctx-item" @click="newFileFromContext">新建文件</div>
      <div class="ctx-item" @click="newFolderFromContext">新建文件夹</div>
      <div class="ctx-item" @click="renameFromContext">重命名</div>
      <div class="ctx-item ctx-item-danger" @click="removeFromContext">删除</div>
    </div>
```

style 增加：
```css
.ctx-item-danger { color: #f56c6c; }
```

修改顶部 "+ 新建" 按钮分裂为两个：
```vue
    <div class="tree-panel-header">
      <span>文件</span>
      <div class="header-actions">
        <el-button size="small" text @click="newFile(null)">+ 文件</el-button>
        <el-button size="small" text @click="newFolderFromContext">+ 文件夹</el-button>
      </div>
    </div>
```

> 注：顶部 "+ 文件夹" 调 `newFolderFromContext` 时 `contextNode.value` 应为 null（在根目录创建）。在调用前 reset：

```ts
function newFolderAtRoot() {
  contextNode.value = null
  newFolderFromContext()
}
```

template 改用 `newFolderAtRoot`:
```vue
<el-button size="small" text @click="newFolderAtRoot">+ 文件夹</el-button>
```

- [ ] **Step 3: 浏览器验证**

操作：
1. 顶部 "+ 文件夹" → 输入 "测试组" → Expected: 根下出现新文件夹
2. 右键 "测试组" → "新建文件" → 在文件夹下创建文件
3. 右键文件 → "重命名" → 改名 → 树更新
4. 右键文件夹 → "删除" → 确认 → Expected: 文件夹及子项全部消失
5. 删除当前打开文件 → Expected: 编辑器变空，filename 清空

- [ ] **Step 4: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/stores/fileStore.ts frontend/src/components/FileTreePanel.vue
git commit -m "feat(scenarios): rename, delete, create folder via context menu"
```

---

### Task 5.8: window.beforeunload 守卫

**Files:**
- Modify: `frontend/src/components/UniverHost.vue`

- [ ] **Step 1: 在 UniverHost 注册 beforeunload**

修改 `frontend/src/components/UniverHost.vue` 在 `onMounted` 末尾增加:
```ts
function onBeforeUnload(e: BeforeUnloadEvent) {
  if (store.dirty) {
    e.preventDefault()
    e.returnValue = '当前文件有未保存的改动，确定离开吗？'
  }
}
window.addEventListener('beforeunload', onBeforeUnload)
```

`onUnmounted` 增加：
```ts
window.removeEventListener('beforeunload', onBeforeUnload)
```

- [ ] **Step 2: 浏览器验证**

操作：
1. 打开文件 → 编辑 → 不保存
2. 关闭浏览器标签 / 刷新
3. Expected: 浏览器原生提示"是否离开"

- [ ] **Step 3: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/src/components/UniverHost.vue
git commit -m "feat(scenarios): beforeunload guard for unsaved changes"
```

---

## 阶段 6：验收（3 任务）

### Task 6.1: 多 fixture round-trip 测试扩展

**Files:**
- Modify: `frontend/tests/xlsxConverter.spec.ts`
- Create: `frontend/public/fixtures/conditional-format.xlsx`
- Create: `frontend/public/fixtures/multi-sheet.xlsx`
- Create: `frontend/public/fixtures/large-10k-rows.xlsx`

- [ ] **Step 1: 用桌面 Excel 准备 3 个补充 fixture**

手动准备 3 个文件放到 `frontend/public/fixtures/`：
- `conditional-format.xlsx` — 含条件格式（色阶/数据条）
- `multi-sheet.xlsx` — 含 3 个 worksheet
- `large-10k-rows.xlsx` — 1 万行数据（A 列 row_1..row_10000，B 列随机数）

- [ ] **Step 2: 扩展测试 fixture 列表**

修改 `frontend/tests/xlsxConverter.spec.ts` 的 fixtures 数组:
```ts
  const fixtures = [
    'simple-100x10.xlsx',
    'merged-cells.xlsx',
    'formulas-sum-vlookup.xlsx',
    'conditional-format.xlsx',
    'multi-sheet.xlsx',
    'large-10k-rows.xlsx',
  ]
```

- [ ] **Step 3: 运行测试**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm test -- tests/xlsxConverter.spec.ts
```

Expected: 6 个测试全部 PASS。如有失败：记录失败的 fixture 与具体差异，触发 R1 fallback 评估（可能需要修改 `workbookEquivalent` 容忍策略，或更换 converter 实现）。

- [ ] **Step 4: Commit**

```bash
cd /Users/churuikai/Desktop/online_excel
git add frontend/tests/xlsxConverter.spec.ts frontend/public/fixtures/
git commit -m "test: extend round-trip coverage to 6 fixtures"
```

---

### Task 6.2: 离线启动 + dist 无外链验证

**Files:**（无文件改动，纯验证）

- [ ] **Step 1: 完整构建**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm run build
```

Expected: 构建成功，`dist/` 目录生成。

- [ ] **Step 2: dist 无外链 grep**

```bash
grep -rE 'https?://(?!localhost)' dist/ | grep -v '\.map' | grep -v 'sourceMappingURL' || echo '✅ no external URLs'
```

Expected: 输出 `✅ no external URLs` 或类似。

> 如果有外链：检查是哪个依赖引入。可能的源头：element-plus 的图标 svg url、Univer 的某个 worker 的远程 fallback。修复后重 build。

- [ ] **Step 3: 离线 preview 验证**

```bash
cd /Users/churuikai/Desktop/online_excel/frontend
npm run preview
```

操作：
1. 浏览器打开 http://localhost:4173
2. DevTools → Network → 勾选 "Offline"
3. 刷新页面
4. Expected: 应用仍然能加载、文件树仍能渲染、能打开文件、能编辑、能下载（下载会从 IndexedDB mock 拿，无网络也行）

- [ ] **Step 4: Commit（如有修复）**

如有修复，commit。否则跳过。

---

### Task 6.3: 手动检查清单走一遍

**Files:**（无文件改动，纯检查）

按 spec §8.3 + 本 plan 的所有场景，逐项验证。建议在新窗口/无痕浏览器跑：

- [ ] **打开 / 编辑 / 保存 / 上传 / 下载 / 新建** 6 个场景全部跑通
- [ ] 切换文件未保存时正确弹守卫（保存 / 放弃 / 取消三选项）
- [ ] `Ctrl+S` 保存正常
- [ ] 关闭标签页未保存时浏览器原生提示
- [ ] 至少 3 个真实业务 .xlsx：上传 → 编辑 → 下载 → Excel 桌面打开**视觉一致**
- [ ] 离线启动：上一 task 已验
- [ ] 重命名 / 删除 / 新建文件夹 工作正常
- [ ] 删除当前打开文件后编辑器清空
- [ ] DevTools Console 无 error 输出（warning 可接受）
- [ ] 失败注入测试：`localStorage.setItem('mockFailRate', '0.5'); location.reload()` 后操作时部分失败 toast 正确显示
- [ ] `localStorage.removeItem('mockFailRate'); location.reload()` 恢复正常
- [ ] `?reset=1` URL 参数清空数据后重新 seed

完成后 commit:
```bash
cd /Users/churuikai/Desktop/online_excel
git commit --allow-empty -m "chore: pass full manual acceptance checklist"
```

---

## 完成标准

全部 task 通过 + 全部测试 PASS + 手动清单全部 ✓ → **MVP 完成**，可进入下一阶段：

- 接入真实 FastAPI 后端（按 spec §7.2 契约实现）
- 启用真实文件锁机制（spec §6 公共机制扩展）
- 性能调优 / 错误监控等

不在本计划范围内。
