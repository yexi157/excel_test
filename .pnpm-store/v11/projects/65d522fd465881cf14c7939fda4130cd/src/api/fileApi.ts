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

async function decodeErrorBody(raw: unknown): Promise<ErrorBody | undefined> {
  if (raw == null) return undefined
  if (typeof raw === 'string') {
    try { return JSON.parse(raw) as ErrorBody } catch { return undefined }
  }
  if (typeof Blob !== 'undefined' && raw instanceof Blob) {
    try { return JSON.parse(await raw.text()) as ErrorBody } catch { return undefined }
  }
  if (raw instanceof ArrayBuffer) {
    try { return JSON.parse(new TextDecoder().decode(raw)) as ErrorBody } catch { return undefined }
  }
  // node Buffer / Uint8Array 等 ArrayBufferView
  if (ArrayBuffer.isView(raw)) {
    try {
      const view = raw as ArrayBufferView
      const ab = view.buffer.slice(view.byteOffset, view.byteOffset + view.byteLength) as ArrayBuffer
      return JSON.parse(new TextDecoder().decode(ab)) as ErrorBody
    } catch { return undefined }
  }
  // 已经是 JSON 对象
  if (typeof raw === 'object') return raw as ErrorBody
  return undefined
}

async function wrap(err: unknown): Promise<ApiError> {
  if (axios.isAxiosError(err)) {
    const ax = err as AxiosError<ErrorBody>
    const status = ax.response?.status ?? 0
    // responseType: 'blob' / 'arraybuffer' 路径下错误响应是二进制/字符串，需要 decode
    const eb = await decodeErrorBody(ax.response?.data)
    if (eb?.error) {
      return new ApiError(eb.error as any, status, eb.message ?? ax.message)
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
      } catch (e) { throw await wrap(e) }
    },

    async renameFile(id, newName) {
      try {
        const res = await client.patch<FileMetadata>(`/files/${id}`, { name: newName })
        return res.data
      } catch (e) { throw await wrap(e) }
    },

    async deleteFile(id) {
      try {
        await client.delete(`/files/${id}`)
      } catch (e) { throw await wrap(e) }
    },

    async createFolder(parentId, name) {
      try {
        const res = await client.post<FileMetadata>('/folders', { parentId, name })
        return res.data
      } catch (e) { throw await wrap(e) }
    },

    async getFile(id) {
      try {
        // 用 arraybuffer 而非 blob：node 端 axios 没有 Blob 适配，会回落成 string；
        // arraybuffer 在浏览器和 node 端都返回二进制（Buffer/ArrayBuffer），统一包成 Blob。
        const res = await client.get<ArrayBuffer>(`/files/${id}`, { responseType: 'arraybuffer' })
        return new Blob([res.data as BlobPart], { type: XLSX_MIME })
      } catch (e) { throw await wrap(e) }
    },

    async saveFile(id, blob) {
      try {
        await client.put(`/files/${id}`, blob, {
          headers: { 'Content-Type': XLSX_MIME },
        })
      } catch (e) { throw await wrap(e) }
    },

    async uploadFile(parentId, name, blob) {
      try {
        const form = new FormData()
        if (parentId) form.append('parentId', parentId)
        form.append('file', new File([blob], name, { type: XLSX_MIME }))
        const res = await client.post<FileMetadata>('/files', form)
        return res.data
      } catch (e) { throw await wrap(e) }
    },
  }
}

export const fileApi = createFileApi()
