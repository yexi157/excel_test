import { openDB, type DBSchema, type IDBPDatabase } from 'idb'
import type { FileMetadata } from '@/api/types'

interface OnlineExcelDB extends DBSchema {
  metadata: {
    key: string                    // file/folder id
    value: FileMetadata
    indexes: { 'by-parentId': string }
  }
  blobs: {
    key: string                    // file id（仅 type=file）
    value: Blob                    // .xlsx 二进制
  }
}

let dbPromise: Promise<IDBPDatabase<OnlineExcelDB>> | null = null

export function getDB(): Promise<IDBPDatabase<OnlineExcelDB>> {
  if (!dbPromise) {
    dbPromise = openDB<OnlineExcelDB>('online-excel-mock', 1, {
      upgrade(db) {
        const meta = db.createObjectStore('metadata', { keyPath: 'id' })
        meta.createIndex('by-parentId', 'parentId')
        db.createObjectStore('blobs')
      },
    })
  }
  return dbPromise
}

export async function resetDB() {
  const db = await getDB()
  await db.clear('metadata')
  await db.clear('blobs')
}
