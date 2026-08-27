import { http, HttpResponse, delay } from 'msw'
import { getDB } from './db'
import type { FileMetadata, FileTreeNode } from '@/api/types'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const MAX_SIZE = 100 * 1024 * 1024  // 100MB

function uuid() { return crypto.randomUUID() }
function nowIso() { return new Date().toISOString() }

async function maybeFail(): Promise<Response | null> {
  await delay(50 + Math.random() * 150)
  const rate = Number(localStorage.getItem('mockFailRate') ?? '0')
  if (rate > 0 && Math.random() < rate) {
    return HttpResponse.json(
      { error: 'INJECTED_FAILURE', message: 'Random failure injection' },
      { status: 500 }
    )
  }
  return null
}

// 把扁平 metadata 列表组装成树
async function buildTree(): Promise<FileTreeNode[]> {
  const db = await getDB()
  const all = await db.getAll('metadata')
  const byId = new Map<string, FileTreeNode>()
  for (const m of all) {
    byId.set(m.id, { id: m.id, name: m.name, type: m.type, parentId: m.parentId, ...(m.type === 'folder' ? { children: [] } : {}) })
  }
  const roots: FileTreeNode[] = []
  for (const node of byId.values()) {
    if (node.parentId === null) {
      roots.push(node)
    } else {
      const parent = byId.get(node.parentId)
      if (parent && parent.children) parent.children.push(node)
    }
  }
  return roots
}

async function nameConflict(parentId: string | null, name: string, excludeId?: string): Promise<boolean> {
  const db = await getDB()
  const all = await db.getAll('metadata')
  return all.some(m => m.parentId === parentId && m.name === name && m.id !== excludeId)
}

export const handlers = [
  // 1. GET /api/files/tree
  http.get('/api/files/tree', async () => {
    const fail = await maybeFail()
    if (fail) return fail
    const tree = await buildTree()
    return HttpResponse.json(tree)
  }),

  // 2. GET /api/files/:id
  http.get('/api/files/:id', async ({ params }) => {
    const fail = await maybeFail()
    if (fail) return fail
    const db = await getDB()
    const meta = await db.get('metadata', params.id as string)
    if (!meta || meta.type !== 'file') {
      return HttpResponse.json({ error: 'FILE_NOT_FOUND', message: 'File not found' }, { status: 404 })
    }
    const blob = await db.get('blobs', params.id as string)
    if (!blob) {
      return HttpResponse.json({ error: 'FILE_NOT_FOUND', message: 'Blob missing' }, { status: 404 })
    }
    return new HttpResponse(blob, {
      status: 200,
      headers: {
        'Content-Type': XLSX_MIME,
        'Content-Disposition': `attachment; filename="${encodeURIComponent(meta.name)}"`,
      },
    })
  }),

  // 3. PUT /api/files/:id
  http.put('/api/files/:id', async ({ request, params }) => {
    const fail = await maybeFail()
    if (fail) return fail
    const db = await getDB()
    const meta = await db.get('metadata', params.id as string)
    if (!meta || meta.type !== 'file') {
      return HttpResponse.json({ error: 'FILE_NOT_FOUND', message: 'File not found' }, { status: 404 })
    }
    const blob = await request.blob()
    if (blob.size > MAX_SIZE) {
      return HttpResponse.json({ error: 'FILE_TOO_LARGE', message: 'Max 100MB' }, { status: 413 })
    }
    await db.put('blobs', blob, meta.id)
    meta.size = blob.size
    meta.mtime = nowIso()
    await db.put('metadata', meta)
    return new HttpResponse(null, { status: 204 })
  }),

  // 4. POST /api/files (multipart upload)
  http.post('/api/files', async ({ request }) => {
    const fail = await maybeFail()
    if (fail) return fail
    const form = await request.formData()
    const parentIdRaw = form.get('parentId')
    const parentId = (typeof parentIdRaw === 'string' && parentIdRaw) ? parentIdRaw : null
    const file = form.get('file')
    if (!(file instanceof File)) {
      return HttpResponse.json({ error: 'INVALID_NAME', message: 'No file in request' }, { status: 400 })
    }
    if (file.size > MAX_SIZE) {
      return HttpResponse.json({ error: 'FILE_TOO_LARGE', message: 'Max 100MB' }, { status: 413 })
    }
    if (await nameConflict(parentId, file.name)) {
      return HttpResponse.json({ error: 'NAME_CONFLICT', message: 'Name already exists' }, { status: 409 })
    }
    const id = uuid()
    const meta: FileMetadata = {
      id, name: file.name, type: 'file', parentId, size: file.size,
      mtime: nowIso(), lockOwner: null, lockExpires: null,
    }
    const db = await getDB()
    await db.put('metadata', meta)
    await db.put('blobs', file, id)
    return HttpResponse.json(meta, { status: 201 })
  }),

  // 5. POST /api/folders
  http.post('/api/folders', async ({ request }) => {
    const fail = await maybeFail()
    if (fail) return fail
    const body = await request.json() as { parentId: string | null; name: string }
    if (!body.name || /[\\/:*?"<>|]/.test(body.name)) {
      return HttpResponse.json({ error: 'INVALID_NAME', message: 'Invalid folder name' }, { status: 400 })
    }
    if (await nameConflict(body.parentId, body.name)) {
      return HttpResponse.json({ error: 'NAME_CONFLICT', message: 'Name already exists' }, { status: 409 })
    }
    const id = uuid()
    const meta: FileMetadata = {
      id, name: body.name, type: 'folder', parentId: body.parentId,
      size: 0, mtime: nowIso(), lockOwner: null, lockExpires: null,
    }
    const db = await getDB()
    await db.put('metadata', meta)
    return HttpResponse.json(meta, { status: 201 })
  }),

  // 6. PATCH /api/files/:id
  http.patch('/api/files/:id', async ({ request, params }) => {
    const fail = await maybeFail()
    if (fail) return fail
    const body = await request.json() as { name: string }
    if (!body.name || /[\\/:*?"<>|]/.test(body.name)) {
      return HttpResponse.json({ error: 'INVALID_NAME', message: 'Invalid name' }, { status: 400 })
    }
    const db = await getDB()
    const meta = await db.get('metadata', params.id as string)
    if (!meta) {
      return HttpResponse.json({ error: 'FILE_NOT_FOUND', message: 'Not found' }, { status: 404 })
    }
    if (await nameConflict(meta.parentId, body.name, meta.id)) {
      return HttpResponse.json({ error: 'NAME_CONFLICT', message: 'Name already exists' }, { status: 409 })
    }
    meta.name = body.name
    meta.mtime = nowIso()
    await db.put('metadata', meta)
    return HttpResponse.json(meta)
  }),

  // 7. DELETE /api/files/:id (递归删除 folder)
  http.delete('/api/files/:id', async ({ params }) => {
    const fail = await maybeFail()
    if (fail) return fail
    const db = await getDB()
    const root = await db.get('metadata', params.id as string)
    if (!root) {
      return HttpResponse.json({ error: 'FILE_NOT_FOUND', message: 'Not found' }, { status: 404 })
    }
    // BFS 收集所有需要删除的 id
    const toDelete: string[] = [root.id]
    const queue = [root.id]
    while (queue.length) {
      const parentId = queue.shift()!
      const all = await db.getAll('metadata')
      for (const m of all) {
        if (m.parentId === parentId) {
          toDelete.push(m.id)
          if (m.type === 'folder') queue.push(m.id)
        }
      }
    }
    for (const id of toDelete) {
      await db.delete('metadata', id)
      await db.delete('blobs', id)
    }
    return new HttpResponse(null, { status: 204 })
  }),
]
