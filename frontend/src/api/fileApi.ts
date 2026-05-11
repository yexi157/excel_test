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
