import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import path from 'node:path'
import ExcelJS from '@zwight/exceljs'
import { xlsxConverter } from '@/utils/xlsxConverter'
import type { ICellData, IStyleData, IWorkbookData, IWorksheetData } from '@univerjs/core'

function loadFixtureBlob(name: string): Blob {
  const buf = readFileSync(path.resolve(__dirname, '../public/fixtures', name))
  // jsdom 环境下 Blob 接受 BufferSource
  return new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
}

type CellMatrix = Record<string, Record<string, ICellData | undefined>>

/**
 * 空 cell 归一：mertdeveci55 round-trip 会把 `{v:null}` 变成 `{v:"",t:1,s:"style_X"}`，
 * 二者语义一致。比较前先归一掉。
 *
 * 判定为空：没有 formula (`f`)，且 v 为 null / undefined / ""。
 */
function isEmptyCell(cell: ICellData | undefined): boolean {
  if (!cell) return true
  if (cell.f) return false
  const v = cell.v
  return v === null || v === undefined || v === ''
}

/** 把单元格归一成可比较的稳定形式。空 cell → 'EMPTY'；非空 → 仅保留语义字段（v、f、t、s）。 */
function normalizeCell(cell: ICellData | undefined): string {
  if (isEmptyCell(cell)) return 'EMPTY'
  return JSON.stringify({ v: cell!.v, f: cell!.f, t: cell!.t, s: cell!.s })
}

function gatherCellKeys(matrix: CellMatrix | undefined): Set<string> {
  const keys = new Set<string>()
  if (!matrix) return keys
  for (const [r, row] of Object.entries(matrix)) {
    if (!row) continue
    for (const c of Object.keys(row)) keys.add(`${r},${c}`)
  }
  return keys
}

/**
 * 等价比较：忽略 mtime / Univer 内部 ID / 空 cellData 字段差异。
 * 比较关键内容：
 *   - sheetOrder 长度
 *   - 每个 sheet 的 name
 *   - 每个 sheet 的 mergeData
 *   - 每个 sheet 的非空 cellData（按 normalizeCell 归一后逐格比较）
 */
function workbookEquivalent(a: IWorkbookData, b: IWorkbookData): boolean {
  if (a.sheetOrder.length !== b.sheetOrder.length) return false
  for (let i = 0; i < a.sheetOrder.length; i++) {
    const sheetA = a.sheets[a.sheetOrder[i]] as IWorksheetData | undefined
    const sheetB = b.sheets[b.sheetOrder[i]] as IWorksheetData | undefined
    if (!sheetA || !sheetB) return false
    if (sheetA.name !== sheetB.name) return false
    if (JSON.stringify(sheetA.mergeData ?? []) !== JSON.stringify(sheetB.mergeData ?? [])) {
      return false
    }
    const cellsA = (sheetA.cellData ?? {}) as CellMatrix
    const cellsB = (sheetB.cellData ?? {}) as CellMatrix
    const allKeys = new Set<string>([...gatherCellKeys(cellsA), ...gatherCellKeys(cellsB)])
    for (const key of allKeys) {
      const [r, c] = key.split(',')
      const cellA = cellsA[r]?.[c]
      const cellB = cellsB[r]?.[c]
      if (normalizeCell(cellA) !== normalizeCell(cellB)) return false
    }
  }
  return true
}

describe('xlsxConverter round-trip', () => {
  const fixtures = [
    'simple-100x10.xlsx',
    'merged-cells.xlsx',
    'formulas-sum-vlookup.xlsx',
    'conditional-format.xlsx',
    'multi-sheet.xlsx',
    'large-10k-rows.xlsx',
  ]

  it.each(fixtures)('%s: blob → IWorkbookData → blob 后内容等价', async (name) => {
    const original = loadFixtureBlob(name)
    const data = await xlsxConverter.toUniver(original, name)
    const exported = await xlsxConverter.toXlsx(data, name)
    const reparsed = await xlsxConverter.toUniver(exported, name)
    expect(workbookEquivalent(data, reparsed)).toBe(true)
  })

  it('粘贴单元格的关闭文本装饰在保存后不会变成下划线和删除线', async () => {
    const data = await xlsxConverter.toUniver(loadFixtureBlob('simple-100x10.xlsx'), 'simple-100x10.xlsx')
    const sheet = data.sheets[data.sheetOrder[0]] as IWorksheetData
    const disabledCell = sheet.cellData?.[0]?.[0] as ICellData
    const enabledCell = sheet.cellData?.[1]?.[0] as ICellData

    data.styles.pastedDisabled = { bl: 0, it: 0, ul: { s: 0 }, st: { s: 0 } }
    disabledCell.s = 'pastedDisabled'
    enabledCell.s = { ...(enabledCell.s as IStyleData), bl: 0, it: 0, ul: { s: 1 }, st: { s: 1 } }

    const exported = await xlsxConverter.toXlsx(data, 'pasted-row.xlsx')
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(await exported.arrayBuffer())
    const exportedSheet = workbook.worksheets[0]

    expect(exportedSheet.getCell('A1').font?.underline).toBeFalsy()
    expect(exportedSheet.getCell('A1').font?.strike).toBeFalsy()
    expect(exportedSheet.getCell('A2').font?.underline).toBe(true)
    expect(exportedSheet.getCell('A2').font?.strike).toBe(true)
  })
})
