# 在线 Excel 系统 — 设计文档

| 字段 | 内容 |
|---|---|
| 日期 | 2026-05-10 |
| 状态 | Draft（待用户 review） |
| 范围 | 前端 + Mock 后端契约；真实后端实现不在本文档范围 |
| 后续 | 通过 review 后转入 implementation plan |

---

## 1. 背景与目标

构建一个浏览器内的在线 Excel 系统：

- **左侧文件树**（element-plus `el-tree`）+ **右侧 Excel 编辑区**
- **单人编辑**：同一时刻一个文件只允许一人打开（无多人协同）
- **后端职责极简**：只负责存取 .xlsx 二进制文件 + 文件树元数据；不解析 Excel 内部结构

本阶段：前端完整功能 + Mock 后端（MSW + IndexedDB 持久化）。真实 FastAPI 后端实现按 §4 契约即可。

## 2. 需求决策（已通过澄清确认）

| 维度 | 决策 |
|---|---|
| Excel 兼容程度 | 完整 Excel 克隆（公式 / 格式 / 合并 / 冻结 / 筛选 / 条件格式 / 批注） |
| 不需要 | 图表、数据透视表、打印 / 导出 PDF |
| License | 仅开源 / 免费（不接受商业授权） |
| 前端框架 | Vue 3 + Vite + TypeScript |
| 文件交换格式 | 原生 .xlsx 双向（与桌面 Excel 互通） |
| 后端 HTTP 契约 | **仅传输 .xlsx 二进制和文件元数据 JSON**；IWorkbookData JSON 永不进入 HTTP body |
| 后端实现语言 | Python FastAPI（**本阶段先 mock**） |
| 单文件规模 | ≤ 1 万行 |
| 单人编辑 | 文件锁机制本阶段 mock，metadata 字段预留 |
| 依赖来源 | 全部 npm 本地安装，禁用任何 CDN |

## 3. 调研结论与方案选择

完整调研记录见会话历史。结论：在 "完全开源 + Vue 3 + 完整 Excel 克隆 + .xlsx 双向" 的硬约束下，唯一可行路径是：

**前端 = [Univer OSS](https://github.com/dream-num/univer)（Apache-2.0） + [Luckyexcel](https://github.com/zwight/Luckyexcel)（MIT）**

- Univer 是原 Luckysheet 团队的下一代产品，开源 Excel 赛道唯一仍在重度投入的项目
- Univer OSS 唯一缺口是 .xlsx I/O（在 Pro 商业版），由 Luckyexcel（社区维护，2025-06 仍活跃）补齐
- Vue 3 一等公民：`@univerjs/ui-adapter-vue3` 官方适配
- 不需要的图表 / 透视表 / 打印恰好是 Univer Pro 的核心增值——刚好绕开

**排除的方案**：

| 方案 | 排除原因 |
|---|---|
| 老 Luckysheet | 已存档停更，jQuery 体系 |
| Fortune-sheet (React) + Vue 包装 | React-in-Vue 技术债，社区被 Univer 接管 |
| x-spreadsheet / wolf-table | npm 0.0.1 早期阶段，已停滞 |
| Jspreadsheet CE / RevoGrid OSS | .xlsx I/O 在 Pro 付费版 |
| Handsontable | 2019 年起非开源，商用付费 |
| OnlyOffice 社区版 | AGPL 传染性 + 需 Document Server，对"后端只存文件"的简单架构过重 |

**风险与退路**：Luckyexcel 是小众库（43 stars），偏门 .xlsx 可能丢格式。退路：将 Luckyexcel 隔离在 `xlsxConverter` 抽象后面，可替换为 (a) SheetJS 自写转换层 或 (b) 后端 Python openpyxl，前端零改动。

---

## 4. 系统架构总览

```
┌──────────────────────────────────────────────────────────────────────┐
│                  浏览器 (Vue 3 + Vite + TypeScript)                   │
│                                                                      │
│  ┌────────────────────┐  ┌──────────────────────────────────────┐    │
│  │  FileTreePanel     │  │      EditorPanel                      │    │
│  │  - element-plus    │  │  ┌────────────────────────────────┐   │    │
│  │    el-tree         │  │  │  ToolbarBar                    │   │    │
│  │  - 新建/重命名/删除│  │  │  上传/下载/保存 + 状态指示      │   │    │
│  │  - 上传 .xlsx      │  │  └────────────────────────────────┘   │    │
│  │  - 选中触发 open   │  │  ┌────────────────────────────────┐   │    │
│  └────────────────────┘  │  │   UniverHost (mount div)       │   │    │
│           │              │  │   - Univer OSS Sheets Core     │   │    │
│           ↓              │  │   - Vue3 Adapter Plugin        │   │    │
│  ┌──────────────────────────────────────────────────────────────┐    │
│  │                fileStore (Pinia)                              │    │
│  │  状态: currentFileId / dirty / saveState / univerInstance    │    │
│  │  动作: openFile / save / upload / download / createNew       │    │
│  └──────────────────────────────────────────────────────────────┘    │
│           │                          │                               │
│           ↓                          ↓                               │
│  ┌────────────────────┐  ┌──────────────────────────────────────┐    │
│  │  api/fileApi.ts    │  │  utils/xlsxConverter.ts (抽象接口)   │    │
│  │  axios + REST      │  │  实现: utils/xlsxConverter.luckyexcel│    │
│  └────────────────────┘  └──────────────────────────────────────┘    │
└──────────────────────────────────────────────────────────────────────┘
            ↓
┌──────────────────────────────────────────────────────────────────────┐
│        Mock 后端 (MSW handlers + IndexedDB 持久化)                    │
│        7 个 RESTful 端点 — 仅交换 .xlsx 二进制 + 元数据 JSON          │
│        未来：替换为 FastAPI + 文件系统/对象存储，契约不变             │
└──────────────────────────────────────────────────────────────────────┘
```

**核心设计原则**

1. **后端只懂 .xlsx 二进制 + 元数据**——所有 IWorkbookData JSON 仅在浏览器内存
2. **`xlsxConverter` 是可替换抽象**——Luckyexcel 出问题不影响其他模块
3. **`fileStore` 是组件间通信的唯一通路**——组件不直接调 API，Univer 实例由 store 持有生命周期
4. **Vue 3 + Univer 框架无关 API + Vue3 Adapter**——不依赖 React-in-Vue 之类妥协

---

## 5. 模块边界与文件结构

### 5.1 目录结构

```
online_excel/
├── frontend/                              # Vue 3 + Vite + TypeScript
│   ├── src/
│   │   ├── main.ts                        # 入口
│   │   ├── App.vue                        # 顶层布局（左树 + 右编辑器）
│   │   ├── components/
│   │   │   ├── FileTreePanel.vue          # 左侧文件树
│   │   │   ├── EditorPanel.vue            # 右侧编辑区容器
│   │   │   ├── ToolbarBar.vue             # 上传/下载/保存 + 状态指示
│   │   │   └── UniverHost.vue             # Univer 实例生命周期封装
│   │   ├── stores/
│   │   │   └── fileStore.ts               # Pinia store
│   │   ├── api/
│   │   │   ├── fileApi.ts                 # axios + REST 客户端
│   │   │   └── types.ts                   # FileTreeNode / FileMetadata
│   │   ├── utils/
│   │   │   ├── xlsxConverter.ts           # 抽象接口 + 默认导出
│   │   │   └── xlsxConverter.luckyexcel.ts # Luckyexcel 实现
│   │   └── mocks/
│   │       ├── handlers.ts                # MSW request handlers
│   │       ├── browser.ts                 # MSW worker 引导
│   │       └── seed.ts                    # IndexedDB 初始数据
│   ├── public/
│   │   ├── mockServiceWorker.js           # MSW 生成的本地 SW
│   │   └── fixtures/                      # 样本 .xlsx
│   ├── tests/
│   │   ├── xlsxConverter.spec.ts
│   │   └── fileStore.spec.ts
│   ├── vite.config.ts
│   └── package.json
└── docs/superpowers/specs/
    └── 2026-05-10-online-excel-design.md
```

### 5.2 单向依赖图

```
   App.vue
     ├── FileTreePanel ─────────────┐
     └── EditorPanel ──────┐        │
           ├── ToolbarBar  │        │
           └── UniverHost  │        │
                           ↓        ↓
                        fileStore (Pinia)
                           │        │
                           ↓        ↓
                    xlsxConverter   fileApi
                           │        │
                           │        ↓
                           │     axios → MSW → 未来 FastAPI
                           ↓
                  Luckyexcel (.xlsx ↔ Univer IWorkbookData)
```

### 5.3 三个核心抽象

**`XlsxConverter` 接口**（让 Luckyexcel 可替换）：
```ts
export interface XlsxConverter {
  toUniver(blob: Blob, fileName: string): Promise<IWorkbookData>
  toXlsx(snapshot: IWorkbookData, fileName: string): Promise<Blob>
}
export const xlsxConverter: XlsxConverter = createLuckyexcelConverter()
```
> `fileName` 在 toUniver 用于构造 File 对象（Luckyexcel API 需要）；在 toXlsx 用于嵌入 .xlsx 元数据。

**`FileApi` 接口**（让 mock 和真实后端可互换）：
```ts
export interface FileApi {
  // —— 元数据 ops（JSON）——
  listTree(): Promise<FileTreeNode[]>
  renameFile(id: string, newName: string): Promise<FileMetadata>
  deleteFile(id: string): Promise<void>
  createFolder(parentId: string | null, name: string): Promise<FileMetadata>

  // —— 内容 ops（.xlsx 二进制）——
  getFile(id: string): Promise<Blob>
  saveFile(id: string, blob: Blob): Promise<void>
  uploadFile(parentId: string | null, name: string, blob: Blob): Promise<FileMetadata>
}
```

**`fileStore`（Pinia）状态**：
```ts
state: {
  treeNodes: FileTreeNode[]
  currentFileId: string | null
  currentFileName: string | null
  dirty: boolean
  saveState: 'idle' | 'saving' | 'saved' | 'error'
  univerInstance: Univer | null     // 由 UniverHost 注入
  univerAPI: FUniver | null
}
```

### 5.4 划分理由

- 三个抽象层完全解耦：converter 不知 api，api 不知 univer，组件不直接调 api
- UniverHost 是唯一持 `Univer` 实例的地方，避免组件分散持有导致内存泄漏；切文件由 store 协调
- 测试聚焦：`xlsxConverter.spec.ts` 验证 round-trip（最大风险点）；`fileStore.spec.ts` 验证状态机

### 5.5 全本地化约束（无 CDN）

**统一原则**：所有 JS/CSS/字体/Worker/图标/mock 数据**全部走 npm 安装到 `node_modules` 或放在 `public/`**，构建产物**自包含**，离线环境也能跑。

| 依赖 | 来源 | 备注 |
|---|---|---|
| Univer 全套 `@univerjs/*` | npm | 默认本地 |
| `@univerjs/ui-adapter-vue3` | npm | 同上 |
| `@univerjs/design/lib/index.css` | npm 路径 import | 不要换成 CDN |
| Univer Formula Worker | Vite `new URL('./worker.ts', import.meta.url)` | Vite 自动 bundle 成本地；**禁止 `workerURL: 'https://...'`** |
| Luckyexcel converter | `npm i @mertdeveci55/univer-import-export xlsx` | 注：原计划用 `@zwight/luckyexcel`，但其嵌套 Univer 0.6.x 与项目 0.22 冲突；R1 fallback 切换为 mertdeveci55（zwight 的 fork，peerDep 模式，honor 项目 Univer 版本）。运行时实际依赖 `xlsx`（SheetJS，未声明），需显式安装 |
| Vue 3 / Vite / TypeScript | npm | 默认本地 |
| element-plus + 图标 | npm + `import 'element-plus/dist/index.css'` | 不要用 unpkg CDN 示例 |
| MSW | npm + `npx msw init public/` | service worker 文件本地 |
| 测试框架 (Vitest) | npm | 本地 |

**vite.config.ts 中明确禁掉**：CDN 外链注入 (`vite-plugin-cdn-import` 等)、`build.rollupOptions.external` 把依赖踢给 CDN。

**验证命令**：
```bash
npm ci --prefer-offline
npm run build
npm run preview         # 断网状态下应能正常打开
# BSD/macOS 兼容版（不用 PCRE 负向预查）：列出所有 http(s) URL，肉眼扫一遍
grep -rEoh 'https?://[^"'\'' )<>]+' frontend/dist/ \
  --include='*.js' --include='*.css' --include='*.html' \
  | grep -v 'localhost' \
  | grep -v 'sourceMappingURL' \
  | sort -u
# 期望：仅 XML 命名空间 / 文档链接 / 占位示例 等元数据，无运行时 CDN/外部资源引用
```

> NOTE: 旧版 `grep -rE 'https?://(?!localhost)' ...` 用了 PCRE 负向预查 `(?!...)`，BSD/macOS `grep -E` 不支持，命令静默非 0 退出，被 `|| echo '✅ ...'` 吞掉，造成假阴性"clean"。上面是可移植替代。

### 5.6 后端契约纯粹性铁律

`getFile` / `saveFile` / `uploadFile` 三个端点的请求/响应体**永远是 .xlsx 二进制**。新建空文件也走 `uploadFile`：前端用 Univer 造空 `IWorkbookData` → `toXlsx` → 上传。后端永不感知 Excel 内部结构。

---

## 6. 关键数据流

### 6.1 公共机制 1：dirty 检测

UniverHost 挂载后注册一次：
```ts
const ignoreInitial = ref(true)
commandService.onCommandExecuted((info) => {
  // 仅 MUTATION 算改数据；OPERATION/COMMAND 不算
  if (info.type === CommandType.MUTATION && !ignoreInitial.value) {
    fileStore.dirty = true
  }
})
// createUnit 完成后 nextTick → ignoreInitial = false
// 切文件时同样 reset
```

### 6.2 公共机制 2：切换/关闭前的 dirty 守卫

```ts
async function guardDirty(): Promise<'continue' | 'cancel'>
// 弹三选一: [保存] [放弃] [取消]
```
触发位置：切文件、上传后自动跳转、`window.beforeunload`

### 6.3 场景调用链

**场景 1 — 打开文件**
```
FileTreePanel.click(nodeId)
  ↓
fileStore.openFile(nodeId)
  1. if ((await guardDirty()) === 'cancel') return          # 用户取消则中止
  2. blob = await fileApi.getFile(nodeId)                   # .xlsx 二进制
  3. workbookData = await xlsxConverter.toUniver(blob)
  4. if (currentFileId) univer.disposeUnit(currentFileId)
  5. univer.createUnit(UNIVER_SHEET, { ...workbookData, id: nodeId })
  6. store: currentFileId, dirty=false, saveState='idle'
  7. nextTick → ignoreInitial=false
```

**场景 2 — 编辑**
```
用户操作 cell → Univer 派发 SET_RANGE_VALUES_MUTATION 等
  → commandService 监听 → store.dirty = true
  → ToolbarBar 显示「未保存」红点
```

**场景 3 — 保存（手动 / Ctrl+S）**
```
fileStore.save()
  1. saveState='saving'
  2. snapshot = univerAPI.getActiveWorkbook()!.save()
  3. blob = await xlsxConverter.toXlsx(snapshot, fileName)
  4. await fileApi.saveFile(currentFileId, blob)
  5. dirty=false, saveState='saved' (3 秒回 idle)
     失败: saveState='error', dirty 保持 true
```

**场景 4 — 上传 .xlsx**
```
ToolbarBar 上传 → <input type="file" accept=".xlsx">
fileStore.upload(parentId, file)
  1. 校验：后缀 .xlsx、size < 100MB
  2. metadata = await fileApi.uploadFile(parentId, file.name, file)
     # 直接传 File（继承 Blob），前端不预解析
  3. await refreshTree()
  4. await openFile(metadata.id)        # 自动打开（走 guardDirty）
```

**场景 5 — 下载 .xlsx**
```
fileStore.download()
  1. if (dirty):
       choice = 弹「先保存？」[是/否/取消]
         是   → await save()，继续 step 2
         否   → 继续 step 2（用户明确选择下载后端版本）
         取消 → return
  2. blob = await fileApi.getFile(currentFileId)   # 总是从后端拿
  3. 浏览器原生下载（URL.createObjectURL + a.click，不引入 file-saver）
```
> 下载从后端拿避免"看到的"和"下载到的"不一致——后端文件即真相。

**场景 6 — 新建空文件**
```
FileTreePanel 右键 → 新建 → 输入名称
fileStore.createNewSheet(parentId, name)
  1. emptyWorkbook = { sheetOrder:['s1'], sheets:{ s1:{ rowCount:100, columnCount:26 } } }
  2. blob = await xlsxConverter.toXlsx(emptyWorkbook, name)
  3. metadata = await fileApi.uploadFile(parentId, name, blob)
  4. refreshTree() + openFile(metadata.id)
```

---

## 7. Mock 后端契约

### 7.1 数据模型

```ts
type NodeType = 'folder' | 'file'

interface FileTreeNode {
  id: string                   // uuid v4
  name: string
  type: NodeType
  parentId: string | null
  children?: FileTreeNode[]    // 仅 folder：完整子树（项目规模小，不做懒加载）
}

interface FileMetadata {
  id: string
  name: string
  type: NodeType
  parentId: string | null
  size: number                 // bytes，文件夹为 0
  mtime: string                // ISO 8601
  lockOwner: string | null     // 预留：未来文件锁
  lockExpires: string | null   // 预留
}

interface ErrorBody {
  error: string                // 如 "FILE_NOT_FOUND"
  message: string
}
```

### 7.2 7 个端点

| # | 方法 | 路径 | 请求体 | 响应（成功） | 错误码 |
|---|---|---|---|---|---|
| 1 | GET | `/api/files/tree` | — | `200` `FileTreeNode[]` | `500` |
| 2 | GET | `/api/files/:id` | — | `200` `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`；Header `Content-Disposition: attachment; filename="..."` | `404 FILE_NOT_FOUND` / `423 FILE_LOCKED` |
| 3 | PUT | `/api/files/:id` | xlsx 二进制 | `204` | `404` / `413 FILE_TOO_LARGE` / `423` |
| 4 | POST | `/api/files` | `multipart/form-data`：`parentId`、`file` (.xlsx) | `201` `FileMetadata` | `409 NAME_CONFLICT` / `413` |
| 5 | POST | `/api/folders` | JSON `{ parentId, name }` | `201` `FileMetadata` (type=folder) | `409` / `400 INVALID_NAME` |
| 6 | PATCH | `/api/files/:id` | JSON `{ name }` | `200` `FileMetadata` | `404` / `409` / `400` |
| 7 | DELETE | `/api/files/:id` | — | `204`（folder 递归删除） | `404` / `423` |

### 7.3 Mock 阶段实现要点

- **存储**：浏览器 IndexedDB（用 `idb` 库），跨刷新持久化
- **Seed**：首次加载 fetch `public/fixtures/*.xlsx` 进 IndexedDB
- **延迟模拟**：handlers 加 `await delay(50~200ms)`
- **失败注入**：`localStorage.setItem('mockFailRate', '0.1')` 让 10% 请求 500
- **文件锁字段预留**：handlers 暂返回 `lockOwner: null`

### 7.4 切换到真实 FastAPI

```ts
// fileApi.ts
const USE_MOCK = import.meta.env.VITE_USE_MOCK !== 'false'
if (USE_MOCK) await import('./mocks/browser').then(m => m.startMock())
```
真实后端实现这 7 个端点的相同契约，前端零改动。

---

## 8. 测试策略与风险

### 8.1 测试金字塔

| 层 | 工具 | 覆盖 | 何时跑 |
|---|---|---|---|
| 单元 | Vitest | `xlsxConverter`（最大风险点）/ `fileStore` 状态机 / 纯函数 utils | 每次提交 |
| 集成 | Vitest + MSW | `fileStore` × `fileApi` × MSW handlers 联调 | 每次提交 |
| E2E（可选） | Playwright | 5 个用户场景在浏览器跑通 | 合并主分支前 |
| 手动检查清单 | — | UI 视觉、3+ 真实业务 .xlsx round-trip、离线启动 | 阶段验收 |

### 8.2 `xlsxConverter` round-trip 必跑 case

```ts
describe('xlsxConverter round-trip', () => {
  it.each([
    'simple-100x10.xlsx',
    'merged-cells.xlsx',
    'formulas-sum-vlookup.xlsx',
    'conditional-format.xlsx',
    'multi-sheet.xlsx',
    'large-10k-rows.xlsx',
  ])('%s: blob → IWorkbookData → blob 后内容等价', async (fixture) => {
    const original = await loadFixture(fixture)
    const data = await converter.toUniver(original)
    const exported = await converter.toXlsx(data, fixture)
    const reparsed = await converter.toUniver(exported)
    expect(workbookEquivalent(data, reparsed)).toBe(true)
  })
})
```
> `workbookEquivalent` 用 deep equal 但忽略 mtime 元字段、Univer 内部 ID、空 cellData 字段差异。

### 8.3 手动检查清单（每个 PR）

- [ ] 6 个场景在浏览器跑通：打开 / 编辑 / 保存 / 上传 / 下载 / 新建
- [ ] 切文件未保存时正确弹守卫
- [ ] `Ctrl+S` 保存正常
- [ ] 关闭标签页未保存浏览器原生提示
- [ ] 至少 3 个真实业务 .xlsx 上传 → 编辑 → 下载 → Excel 桌面打开**视觉一致**
- [ ] 离线启动：`npm run build && npm run preview` 后断网刷新仍可工作
- [ ] dist 外链审计（BSD/macOS 兼容；输出仅含元数据，无运行时 CDN）：
  ```bash
  grep -rEoh 'https?://[^"'\'' )<>]+' frontend/dist/ \
    --include='*.js' --include='*.css' --include='*.html' \
    | grep -v 'localhost' | grep -v 'sourceMappingURL' | sort -u
  ```

### 8.4 风险登记

| # | 风险 | 影响 | 概率 | 缓解 |
|---|---|---|---|---|
| R1 | Luckyexcel 在偏门 .xlsx 文件上转换失败/丢格式 | 高 | 中 | 阶段 1 立刻 PoC：5+ 真实业务文件 round-trip。失败则 fallback 到 SheetJS 自写转换层（接口不变） |
| R2 | Univer 切文件 disposeUnit + createUnit 出现渲染残留 | 中 | 低 | 退到"每文件重建 Univer 实例"模式（skill 文档明确推荐多实例） |
| R3 | `ignoreInitial` 时序未覆盖某些初始 mutation | 中 | 中 | 测试覆盖：打开后 dirty 必须 false。仍漏改用 setTimeout(100ms) 兜底 |
| R4 | `workbook.save()` 大文件 JSON 序列化卡顿 | 低 | 低 | 项目限定 ≤ 1 万行，profile 验证；必要时 toXlsx 放 Web Worker |
| R5 | IndexedDB 配额或损坏导致 mock 数据丢失 | 低 | 低 | mock 仅开发用；提供 `?reset=1` 清空重 seed |
| R6 | 切换真实 FastAPI 时契约偏差 | 高 | 中 | §7 契约文档作为后端开发的 source of truth；MSW handlers 也作为参考实现 |
| R7 | element-plus 与 Univer 样式冲突（z-index / 字体 / scoped） | 中 | 中 | UniverHost 用独立 `<div ref="container">` 包裹；初期遇到再调样式 |

---

## 9. 阶段 1 PoC 验收标准

**在写其他代码之前必须完成**：

```
Day 1-2: 最小可行 PoC
  - Vite + Vue 3 + TypeScript 脚手架
  - 集成 Univer OSS + Vue3 Adapter（按官方 vue3-vite 模板）
  - 集成 Luckyexcel
  - 用 1 个真实 .xlsx 跑通：fetch → toUniver → 显示 → 编辑 → toXlsx → 下载 → Excel 桌面打开
```

**通过条件**：下载后的 .xlsx 在桌面 Microsoft Excel / WPS 中打开**视觉一致、可继续编辑**。

**不通过的处理**：触发 R1 fallback 决策——评估 SheetJS 自写转换层 或 后端 Python openpyxl，更新本文档后重新 PoC。

---

## 10. 范围之外（明确不做）

- 多人实时协同
- 图表 / 数据透视表 / 打印 / 导出 PDF
- 权限系统（用户、组织、ACL）
- 版本历史 / 快照
- 全文搜索
- 移动端适配
- 真实文件锁实现（仅预留字段）
- 真实后端 FastAPI 实现（仅给契约）
