<script setup lang="ts">
import { onMounted, onUnmounted, ref } from 'vue'
import { CommandType, ICommandService, LocaleType, LogLevel, Univer, merge } from '@univerjs/core'
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
let mutationDisposable: { dispose: () => void } | null = null

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
  univer = new Univer({
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
  univer.registerPlugins([
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

  const api = FUniver.newAPI(univer)
  store.bindUniver(univer, api)

  // 监听 mutation 设 dirty（spec §6.1）
  const commandService = (univer as any).__getInjector().get(ICommandService)
  mutationDisposable = commandService.onCommandExecuted((info: { type: number }) => {
    if (info.type === CommandType.MUTATION) {
      store.markDirty()
    }
  })

  window.addEventListener('keydown', onKeydown)
  window.addEventListener('beforeunload', onBeforeUnload)
})

onUnmounted(() => {
  mutationDisposable?.dispose()
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
