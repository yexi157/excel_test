// @vitest-environment node
// 用 node 环境而非 jsdom：jsdom 把 globalThis.File / FormData 替换为自己的实现，
// 与 Node 24 内置 undici v7 的 multipart parser 不兼容（webidl File brand check 失败）。
// 本套测试只需 IndexedDB（fake-indexeddb 在 node 下也工作），不需要 DOM。
import { describe, it, expect, beforeAll, afterAll, beforeEach } from 'vitest'
import { setupServer } from 'msw/node'
import { handlers } from '@/mocks/handlers'
import { createFileApi } from '@/api/fileApi'
import { resetDB } from '@/mocks/db'
import axios from 'axios'

// node 端 axios 需要绝对 URL；handlers 用相对路径 '/api/...'，
// MSW 用 globalThis.location（setup.ts 已 stub 为 http://localhost/）解析为绝对 URL，
// 故 baseURL origin 必须与 setup.ts 中 location 一致。
// 使用 axios 默认 http adapter（node 下原生 multipart 序列化与 undici parser 兼容良好）。
const api = createFileApi(axios.create({
  baseURL: 'http://localhost/api',
}))

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
