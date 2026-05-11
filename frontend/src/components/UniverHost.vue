<script setup lang="ts">
import { onMounted, onUnmounted, ref } from 'vue'
import { LocaleType, LogLevel, Univer, UniverInstanceType, merge, type IWorkbookData } from '@univerjs/core'
import { defaultTheme } from '@univerjs/themes'
import { UniverRenderEnginePlugin } from '@univerjs/engine-render'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverUIPlugin } from '@univerjs/ui'
import { UniverVue3AdapterPlugin } from '@univerjs/ui-adapter-vue3'
import { UniverDocsPlugin } from '@univerjs/docs'
import { UniverDocsUIPlugin } from '@univerjs/docs-ui'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import { UniverSheetsFormulaUIPlugin } from '@univerjs/sheets-formula-ui'
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt'
import { UniverSheetsNumfmtUIPlugin } from '@univerjs/sheets-numfmt-ui'
import { FUniver } from '@univerjs/core/facade'

// zh-CN locale dictionaries — must merge per-plugin packs or LocaleService 报 not initialized
import DesignZhCN from '@univerjs/design/locale/zh-CN'
import UIZhCN from '@univerjs/ui/locale/zh-CN'
import DocsUIZhCN from '@univerjs/docs-ui/locale/zh-CN'
import SheetsZhCN from '@univerjs/sheets/locale/zh-CN'
import SheetsUIZhCN from '@univerjs/sheets-ui/locale/zh-CN'
import SheetsFormulaZhCN from '@univerjs/sheets-formula/locale/zh-CN'
import SheetsFormulaUIZhCN from '@univerjs/sheets-formula-ui/locale/zh-CN'
import SheetsNumfmtUIZhCN from '@univerjs/sheets-numfmt-ui/locale/zh-CN'

import '@univerjs/design/lib/index.css'
import '@univerjs/ui/lib/index.css'
import '@univerjs/docs-ui/lib/index.css'
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
let dirtyDisposable: { dispose: () => void } | null = null

/**
 * 构造一个新的 Univer 实例（含 plugins / dirty listener）。
 * 可选：附带 workbookData 时，会立刻 createUnit。
 *
 * dirty 检测策略（spec §6.1）：
 *   - 监听 facade 的 SheetValueChanged：只在 SetRangeValuesMutation 等
 *     真正改值的命令被 dispatch 时触发（用户编辑 / 外部 API 调用）
 *   - createUnit 走数据模型构造路径，不会派发这些 mutation，
 *     因此初始化期间不会触发 markDirty
 *   - 比 commandService.onCommandExecuted 干净：后者在 init 期间被
 *     大量底层 mutation（formula / skeleton / selections）误触发
 */
function buildUniver(workbookData?: IWorkbookData) {
  const u = new Univer({
    theme: defaultTheme,
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: merge(
        {},
        DesignZhCN,
        UIZhCN,
        DocsUIZhCN,
        SheetsZhCN,
        SheetsUIZhCN,
        SheetsFormulaZhCN,
        SheetsFormulaUIZhCN,
        SheetsNumfmtUIZhCN,
      ),
    },
    logLevel: LogLevel.WARN,
  })
  u.registerPlugins([
    [UniverRenderEnginePlugin],
    [UniverFormulaEnginePlugin],
    [UniverUIPlugin, { container: container.value! }],
    [UniverVue3AdapterPlugin],
    [UniverDocsPlugin],          // docs core, required by sheets-ui's EditorBridgeService
    [UniverDocsUIPlugin],        // provides IEditorService
    [UniverSheetsPlugin],
    [UniverSheetsUIPlugin],
    [UniverSheetsFormulaPlugin],
    [UniverSheetsFormulaUIPlugin],
    [UniverSheetsNumfmtPlugin],
    [UniverSheetsNumfmtUIPlugin],
  ])

  if (workbookData) {
    u.createUnit(UniverInstanceType.UNIVER_SHEET, workbookData)
  }

  const api = FUniver.newAPI(u)

  const disposable = api.addEvent(api.Event.SheetValueChanged, () => {
    store.markDirty()
  })

  return { univer: u, api, disposable }
}

/**
 * Spec §8.4 R2 fallback：完全重建 Univer 实例。
 * 解决 disposeUnit 后 SheetsSelectionsService / RefSelectionsRenderService 残留 stale unit id 的崩溃。
 */
function recreate(workbookData: IWorkbookData) {
  dirtyDisposable?.dispose()
  univer?.dispose()
  dirtyDisposable = null
  univer = null

  const built = buildUniver(workbookData)
  univer = built.univer
  dirtyDisposable = built.disposable
  store.bindUniver(built.univer, built.api, recreate)
}

function onKeydown(e: KeyboardEvent) {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
    e.preventDefault()
    if (store.currentFileId) {
      store.save().catch(err => console.error('[Ctrl+S save] failed:', err))
    }
  }
}

function onBeforeUnload(e: BeforeUnloadEvent) {
  if (store.dirty) {
    e.preventDefault()
    e.returnValue = '当前文件有未保存的改动，确定离开吗？'
  }
}

onMounted(() => {
  const built = buildUniver()
  univer = built.univer
  dirtyDisposable = built.disposable
  store.bindUniver(built.univer, built.api, recreate)

  window.addEventListener('keydown', onKeydown)
  window.addEventListener('beforeunload', onBeforeUnload)
})

onUnmounted(() => {
  dirtyDisposable?.dispose()
  univer?.dispose()
  window.removeEventListener('keydown', onKeydown)
  window.removeEventListener('beforeunload', onBeforeUnload)
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
