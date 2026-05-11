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
    store.setIgnoreInitial(false)
    store.markDirty()
    expect(store.dirty).toBe(true)
  })

  it('markDirty 在 _ignoreInitial=true 时不改 dirty', () => {
    const store = useFileStore()
    // 初始 _ignoreInitial 为 true，markDirty 应直接 return
    store.markDirty()
    expect(store.dirty).toBe(false)
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
