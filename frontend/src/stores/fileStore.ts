import { defineStore } from 'pinia'
import type { Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/facade'
// Side-effect import: augments FUniver with sheet-related methods (getActiveWorkbook, etc.)
import '@univerjs/sheets/facade'
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
  _ignoreInitial: boolean
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
    _ignoreInitial: true,
  }),

  actions: {
    /** 由 UniverHost 在 onMounted 调用 */
    bindUniver(instance: Univer, api: FUniver) {
      this.univerInstance = instance
      this.univerAPI = api
    },

    /** 切换文件时由 store 设为 true，加载完后 nextTick 设回 false */
    setIgnoreInitial(value: boolean) {
      this._ignoreInitial = value
    },

    /** 由 UniverHost 在 commandService 监听到 mutation 时调用 */
    markDirty() {
      if (this._ignoreInitial) return
      this.dirty = true
    },

    async refreshTree() {
      this.treeNodes = await fileApi.listTree()
    },

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
        this.univerAPI?.disposeUnit(this.currentFileId)
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

    // 后续补充：rename/delete/createFolder (5.7)
  },
})

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
