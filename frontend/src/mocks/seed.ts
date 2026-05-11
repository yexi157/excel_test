import { getDB } from './db'
import type { FileMetadata } from '@/api/types'

const SEED_FIXTURES = [
  'simple-100x10.xlsx',
  'merged-cells.xlsx',
  'formulas-sum-vlookup.xlsx',
]

function uuid(): string {
  return crypto.randomUUID()
}

function nowIso(): string {
  return new Date().toISOString()
}

export async function seedIfEmpty(): Promise<void> {
  const db = await getDB()
  const count = await db.count('metadata')
  if (count > 0) return

  // 创建一个根文件夹
  const rootId = uuid()
  const root: FileMetadata = {
    id: rootId,
    name: '示例文件夹',
    type: 'folder',
    parentId: null,
    size: 0,
    mtime: nowIso(),
    lockOwner: null,
    lockExpires: null,
  }
  await db.put('metadata', root)

  // 注入示例文件
  for (const name of SEED_FIXTURES) {
    const response = await fetch(`/fixtures/${name}`)
    if (!response.ok) {
      console.warn(`[seed] fixture ${name} fetch failed: ${response.status}`)
      continue
    }
    const blob = await response.blob()
    const id = uuid()
    const meta: FileMetadata = {
      id,
      name,
      type: 'file',
      parentId: rootId,
      size: blob.size,
      mtime: nowIso(),
      lockOwner: null,
      lockExpires: null,
    }
    await db.put('metadata', meta)
    await db.put('blobs', blob, id)
  }
}
