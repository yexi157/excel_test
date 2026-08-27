// File system 类型（spec §7.1）

export type NodeType = 'folder' | 'file'

export interface FileTreeNode {
  id: string                   // uuid v4
  name: string
  type: NodeType
  parentId: string | null
  children?: FileTreeNode[]    // 仅 folder
}

export interface FileMetadata {
  id: string
  name: string
  type: NodeType
  parentId: string | null
  size: number                 // bytes，folder 为 0
  mtime: string                // ISO 8601
  lockOwner: string | null     // 预留：未来文件锁
  lockExpires: string | null   // 预留
}

export interface ErrorBody {
  error: string                // 机器可读，如 "FILE_NOT_FOUND"
  message: string
}

// 业务异常分类
export type ApiErrorCode =
  | 'FILE_NOT_FOUND'
  | 'FILE_LOCKED'
  | 'FILE_TOO_LARGE'
  | 'NAME_CONFLICT'
  | 'INVALID_NAME'
  | 'NETWORK'
  | 'UNKNOWN'

export class ApiError extends Error {
  constructor(public code: ApiErrorCode, public httpStatus: number, message: string) {
    super(message)
    this.name = 'ApiError'
  }
}
