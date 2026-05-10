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
import LuckyExcel from '@zwight/luckyexcel'

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
  await new Promise<void>((resolve, reject) => {
    LuckyExcel.transformExcelToUniver(
      file,
      (workbookData /* IWorkbookData */) => {
        univerAPI?.disposeUnit(currentUnitId)
        currentUnitId = (workbookData as { id?: string }).id || 'fixture-workbook'
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
    LuckyExcel.transformUniverToExcel({
      snapshot,
      fileName: 'poc-export.xlsx',
      success: () => resolve(),
      error: (err: Error) => reject(err),
    })
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
