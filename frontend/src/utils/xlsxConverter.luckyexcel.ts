import type { ICellData, IWorkbookData } from '@univerjs/core'
import LuckyExcel from '@zwight/luckyexcel'
import type { XlsxConverter } from './xlsxConverter'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

function toArrayBuffer(value: ArrayBuffer | ArrayBufferView): ArrayBuffer {
  if (value instanceof ArrayBuffer) return value

  return value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength) as ArrayBuffer
}

function adaptSnapshotForExport(snapshot: IWorkbookData): IWorkbookData {
  let workbookChanged = false
  const sheets = { ...snapshot.sheets }

  for (const sheetId of snapshot.sheetOrder) {
    const sheet = snapshot.sheets[sheetId]
    if (!sheet?.cellData) continue

    let sheetChanged = false
    const cellData = { ...sheet.cellData }

    for (const [rowIndex, sourceRow] of Object.entries(sheet.cellData)) {
      if (!sourceRow) continue

      let rowChanged = false
      const row = { ...sourceRow }
      for (const [columnIndex, sourceCell] of Object.entries(sourceRow)) {
        const cell = sourceCell as ICellData | undefined
        if (!cell?.f) continue

        // @zwight/luckyexcel 1.1.6 reads `si` as the formula text during export.
        row[Number(columnIndex)] = { ...cell, f: null, si: cell.f }
        rowChanged = true
      }

      if (rowChanged) {
        cellData[Number(rowIndex)] = row
        sheetChanged = true
      }
    }

    if (sheetChanged) {
      sheets[sheetId] = { ...sheet, cellData }
      workbookChanged = true
    }
  }

  return workbookChanged ? { ...snapshot, sheets } : snapshot
}

export function createLuckyexcelConverter(): XlsxConverter {
  return {
    async toUniver(blob: Blob, fileName: string): Promise<IWorkbookData> {
      const file = new File([blob], fileName)
      return new Promise<IWorkbookData>((resolve, reject) => {
        LuckyExcel.transformExcelToUniver(
          file,
          (workbookData: IWorkbookData) => resolve(workbookData),
          (err: Error) => reject(err),
        )
      })
    },

    async toXlsx(snapshot: IWorkbookData, fileName: string): Promise<Blob> {
      return new Promise<Blob>((resolve, reject) => {
        LuckyExcel.transformUniverToExcel({
          snapshot: adaptSnapshotForExport(snapshot),
          fileName,
          getBuffer: true,
          success: (buffer?: ArrayBuffer | ArrayBufferView) => {
            if (!buffer) {
              reject(new Error('LuckyExcel returned empty buffer'))
              return
            }
            resolve(new Blob([toArrayBuffer(buffer)], { type: XLSX_MIME }))
          },
          error: (err: Error) => reject(err),
        })
      })
    },
  }
}
