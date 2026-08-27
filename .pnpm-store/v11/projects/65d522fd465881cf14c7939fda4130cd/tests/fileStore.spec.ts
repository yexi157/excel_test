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

  it('openFile 完成后状态正确', async () => {
    const store = useFileStore()
    store.treeNodes = [
      { id: 'f1', name: 'a.xlsx', type: 'file', parentId: null },
    ]
    const fakeUniver = {} as any
    const fakeAPI = {} as any
    const mockRecreate = vi.fn()
    store.bindUniver(fakeUniver, fakeAPI, mockRecreate)

    // mock fileApi.getFile + xlsxConverter.toUniver
    const { fileApi } = await import('@/api/fileApi')
    const { xlsxConverter } = await import('@/utils/xlsxConverter')
    vi.mocked(fileApi.getFile = vi.fn().mockResolvedValue(new Blob(['x'])))
    vi.mocked(xlsxConverter.toUniver = vi.fn().mockResolvedValue({ id: 'f1', sheetOrder: ['s1'], sheets: { s1: {} } }))

    await store.openFile('f1')
    expect(store.currentFileId).toBe('f1')
    expect(store.currentFileName).toBe('a.xlsx')
    expect(store.dirty).toBe(false)
    expect(mockRecreate).toHaveBeenCalledWith(expect.objectContaining({ id: 'f1' }))
  })
})
